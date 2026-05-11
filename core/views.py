from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login
from django.contrib import messages
from django.utils import timezone
from django.db.models import Avg, Count, Max, Q
from django.contrib.auth import get_user_model
import json

from .models import Store, Visit, VisitPhoto, VisitCheckItem, StoreSKUPlan
from .forms import VisitPhotoForm, VisitNotesForm, AnalyticsFilterForm, UserRegistrationForm

User = get_user_model()


def register(request):
    if request.user.is_authenticated:
        return redirect('dashboard')

    if request.method == 'POST':
        form = UserRegistrationForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)
            user.role = 'supervisor'
            user.save()
            messages.success(request, 'Регистрация успешна! Войдите в аккаунт.')
            return redirect('login')
    else:
        form = UserRegistrationForm()

    return render(request, 'registration/register.html', {'form': form})


@login_required
def dashboard(request):
    today = timezone.localdate()
    visits = Visit.objects.filter(supervisor=request.user, date=today).select_related('store')
    return render(request, 'dashboard.html', {'visits': visits, 'today': today})


@login_required
def store_list(request):
    q = request.GET.get('q', '').strip()
    stores = Store.objects.filter(is_active=True)
    if q:
        stores = stores.filter(Q(name__icontains=q) | Q(network__icontains=q))
    return render(request, 'stores/list.html', {'stores': stores, 'q': q})


@login_required
def visit_new(request):
    store_id = request.GET.get('store')
    store = get_object_or_404(Store, pk=store_id, is_active=True)

    if request.method == 'POST':
        visit = Visit.objects.create(supervisor=request.user, store=store)
        plan_skus = StoreSKUPlan.objects.filter(store=store).select_related('sku')
        check_items = [
            VisitCheckItem(visit=visit, sku=plan.sku, status='not_checked')
            for plan in plan_skus
        ]
        VisitCheckItem.objects.bulk_create(check_items)
        return redirect('visit_detail', pk=visit.pk)

    return render(request, 'visits/new.html', {'store': store})


@login_required
def visit_detail(request, pk):
    visit = get_object_or_404(Visit, pk=pk, supervisor=request.user)
    photo_form = VisitPhotoForm()
    notes_form = VisitNotesForm(instance=visit)

    if request.method == 'POST':
        action = request.POST.get('action')

        if action == 'upload_photo':
            photo_form = VisitPhotoForm(request.POST, request.FILES)
            if photo_form.is_valid():
                photo = photo_form.save(commit=False)
                photo.visit = visit
                photo.save()
                messages.success(request, 'Фото загружено.')
                return redirect('visit_detail', pk=pk)

        elif action == 'save_notes':
            notes_form = VisitNotesForm(request.POST, instance=visit)
            if notes_form.is_valid():
                notes_form.save()
                messages.success(request, 'Заметки сохранены.')
                return redirect('visit_detail', pk=pk)

        elif action == 'update_check':
            item_id = request.POST.get('item_id')
            status = request.POST.get('status')
            if item_id and status in dict(VisitCheckItem.STATUS_CHOICES):
                VisitCheckItem.objects.filter(
                    pk=item_id, visit=visit
                ).update(status=status)
            return redirect('visit_detail', pk=pk)

        elif action == 'complete_visit':
            items = visit.check_items.all()
            total = items.count()
            present = items.filter(status='present').count()
            visit.completion_pct = (present / total * 100) if total > 0 else 0
            visit.status = 'completed'
            visit.save()
            return redirect('visit_report', pk=pk)

    check_items = visit.check_items.select_related('sku').all()
    total = check_items.count()
    checked = check_items.exclude(status='not_checked').count()
    photos = visit.photos.all()
    active_tab = request.GET.get('tab', 'photos')

    return render(request, 'visits/detail.html', {
        'visit': visit,
        'photo_form': photo_form,
        'notes_form': notes_form,
        'check_items': check_items,
        'total': total,
        'checked': checked,
        'photos': photos,
        'active_tab': active_tab,
    })


@login_required
def visit_report(request, pk):
    visit = get_object_or_404(Visit, pk=pk, supervisor=request.user)
    check_items = visit.check_items.select_related('sku').all()
    photos = visit.photos.all()
    return render(request, 'visits/report.html', {
        'visit': visit,
        'check_items': check_items,
        'photos': photos,
    })


@login_required
def visit_list(request):
    visits = Visit.objects.filter(supervisor=request.user).select_related('store')
    return render(request, 'visits/list.html', {'visits': visits})


@login_required
def analytics(request):
    if request.user.is_superuser:
        return supervisor_analytics(request)
    else:
        return personal_analytics(request)


def personal_analytics(request):
    form = AnalyticsFilterForm(request.GET or None)
    date_from = None
    date_to = None

    if form.is_valid():
        date_from = form.cleaned_data.get('date_from')
        date_to = form.cleaned_data.get('date_to')

    # Только визиты текущего пользователя
    visits_qs = Visit.objects.filter(supervisor=request.user, status='completed')
    if date_from:
        visits_qs = visits_qs.filter(date__gte=date_from)
    if date_to:
        visits_qs = visits_qs.filter(date__lte=date_to)

    # Статистика по магазинам
    store_stats = (
        visits_qs
        .values('store__id', 'store__name', 'store__network')
        .annotate(
            visit_count=Count('id'),
            avg_pct=Avg('completion_pct'),
            last_visit=Max('date'),
        )
        .order_by('avg_pct')
    )

    # KPI метрики
    total_visits = visits_qs.count()
    avg_completion = visits_qs.aggregate(avg=Avg('completion_pct'))['avg'] or 0
    total_stores = visits_qs.values('store').distinct().count()

    # График по дням
    daily_stats = (
        visits_qs
        .extra(select={'day': 'DATE(date)'})
        .values('day')
        .annotate(count=Count('id'), avg_pct=Avg('completion_pct'))
        .order_by('day')
    )
    daily_labels = [str(stat['day']) for stat in daily_stats]
    daily_counts = [stat['count'] for stat in daily_stats]
    daily_pcts = [round(stat['avg_pct'] or 0, 1) for stat in daily_stats]

    # График по магазинам
    store_labels = [s['store__name'] for s in store_stats]
    store_pcts = [round(s['avg_pct'] or 0, 1) for s in store_stats]

    # Распределение по % выполнения
    ranges = [
        {'min': 0, 'max': 25, 'label': '0-25%'},
        {'min': 25, 'max': 50, 'label': '25-50%'},
        {'min': 50, 'max': 75, 'label': '50-75%'},
        {'min': 75, 'max': 100, 'label': '75-100%'},
    ]
    dist_labels = []
    dist_counts = []
    for r in ranges:
        count = visits_qs.filter(completion_pct__gte=r['min'], completion_pct__lt=r['max']).count()
        if r['min'] == 75:
            count = visits_qs.filter(completion_pct__gte=r['min'], completion_pct__lte=100).count()
        dist_labels.append(r['label'])
        dist_counts.append(count)

    return render(request, 'analytics/personal.html', {
        'form': form,
        'store_stats': store_stats,
        'date_from': date_from,
        'date_to': date_to,
        'total_visits': total_visits,
        'avg_completion': round(avg_completion, 1),
        'total_stores': total_stores,
        'daily_labels': json.dumps(daily_labels),
        'daily_counts': json.dumps(daily_counts),
        'daily_pcts': json.dumps(daily_pcts),
        'store_labels': json.dumps(store_labels),
        'store_pcts': json.dumps(store_pcts),
        'dist_labels': json.dumps(dist_labels),
        'dist_counts': json.dumps(dist_counts),
    })


def supervisor_analytics(request):
    form = AnalyticsFilterForm(request.GET or None)
    date_from = None
    date_to = None

    if form.is_valid():
        date_from = form.cleaned_data.get('date_from')
        date_to = form.cleaned_data.get('date_to')

    # Все завершённые визиты
    visits_qs = Visit.objects.filter(status='completed')
    if date_from:
        visits_qs = visits_qs.filter(date__gte=date_from)
    if date_to:
        visits_qs = visits_qs.filter(date__lte=date_to)

    # Статистика по магазинам
    store_stats = (
        visits_qs
        .values('store__id', 'store__name', 'store__network')
        .annotate(
            visit_count=Count('id'),
            avg_pct=Avg('completion_pct'),
            last_visit=Max('date'),
        )
        .order_by('avg_pct')
    )

    # Статистика по мерчандайзерам
    supervisor_stats = (
        visits_qs
        .values('supervisor__id', 'supervisor__first_name', 'supervisor__last_name', 'supervisor__username')
        .annotate(
            visit_count=Count('id'),
            avg_pct=Avg('completion_pct'),
            last_visit=Max('date'),
        )
        .order_by('-visit_count')
    )

    # KPI метрики
    total_visits = visits_qs.count()
    avg_completion = visits_qs.aggregate(avg=Avg('completion_pct'))['avg'] or 0
    total_supervisors = User.objects.filter(is_superuser=False).count()
    total_stores = Store.objects.filter(is_active=True).count()

    # График по дням
    daily_stats = (
        visits_qs
        .extra(select={'day': 'DATE(date)'})
        .values('day')
        .annotate(count=Count('id'), avg_pct=Avg('completion_pct'))
        .order_by('day')
    )
    daily_labels = [str(stat['day']) for stat in daily_stats]
    daily_counts = [stat['count'] for stat in daily_stats]
    daily_pcts = [round(stat['avg_pct'] or 0, 1) for stat in daily_stats]

    # График по магазинам
    store_labels = [s['store__name'] for s in store_stats]
    store_pcts = [round(s['avg_pct'] or 0, 1) for s in store_stats]

    # График по мерчандайзерам
    supervisor_labels = [
        f"{s['supervisor__first_name'] or s['supervisor__username']}"
        for s in supervisor_stats
    ]
    supervisor_visits = [s['visit_count'] for s in supervisor_stats]
    supervisor_pcts = [round(s['avg_pct'] or 0, 1) for s in supervisor_stats]

    # Распределение по % выполнения
    ranges = [
        {'min': 0, 'max': 25, 'label': '0-25%'},
        {'min': 25, 'max': 50, 'label': '25-50%'},
        {'min': 50, 'max': 75, 'label': '50-75%'},
        {'min': 75, 'max': 100, 'label': '75-100%'},
    ]
    dist_labels = []
    dist_counts = []
    for r in ranges:
        count = visits_qs.filter(completion_pct__gte=r['min'], completion_pct__lt=r['max']).count()
        if r['min'] == 75:
            count = visits_qs.filter(completion_pct__gte=r['min'], completion_pct__lte=100).count()
        dist_labels.append(r['label'])
        dist_counts.append(count)

    return render(request, 'analytics/supervisor.html', {
        'form': form,
        'store_stats': store_stats,
        'supervisor_stats': supervisor_stats,
        'date_from': date_from,
        'date_to': date_to,
        'total_visits': total_visits,
        'avg_completion': round(avg_completion, 1),
        'total_supervisors': total_supervisors,
        'total_stores': total_stores,
        'daily_labels': json.dumps(daily_labels),
        'daily_counts': json.dumps(daily_counts),
        'daily_pcts': json.dumps(daily_pcts),
        'store_labels': json.dumps(store_labels),
        'store_pcts': json.dumps(store_pcts),
        'supervisor_labels': json.dumps(supervisor_labels),
        'supervisor_visits': json.dumps(supervisor_visits),
        'supervisor_pcts': json.dumps(supervisor_pcts),
        'dist_labels': json.dumps(dist_labels),
        'dist_counts': json.dumps(dist_counts),
    })
