from django.db import models
from django.contrib.auth.models import AbstractUser


class User(AbstractUser):
    ROLE_CHOICES = [('supervisor', 'Супервайзер')]
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='supervisor')

    class Meta:
        verbose_name = 'Пользователь'
        verbose_name_plural = 'Пользователи'


class Store(models.Model):
    name = models.CharField(max_length=200, verbose_name='Название')
    address = models.CharField(max_length=500, verbose_name='Адрес')
    network = models.CharField(max_length=200, verbose_name='Торговая сеть')
    is_active = models.BooleanField(default=True, verbose_name='Активен')

    class Meta:
        verbose_name = 'Магазин'
        verbose_name_plural = 'Магазины'
        ordering = ['name']

    def __str__(self):
        return f'{self.name} ({self.network})'


class SKU(models.Model):
    name = models.CharField(max_length=300, verbose_name='Название')
    barcode = models.CharField(max_length=100, unique=True, verbose_name='Штрихкод')
    category = models.CharField(max_length=200, verbose_name='Категория')

    class Meta:
        verbose_name = 'SKU'
        verbose_name_plural = 'SKU'
        ordering = ['category', 'name']

    def __str__(self):
        return f'{self.name} ({self.barcode})'


class StoreSKUPlan(models.Model):
    store = models.ForeignKey(Store, on_delete=models.CASCADE, related_name='sku_plans', verbose_name='Магазин')
    sku = models.ForeignKey(SKU, on_delete=models.CASCADE, related_name='store_plans', verbose_name='SKU')

    class Meta:
        verbose_name = 'План SKU магазина'
        verbose_name_plural = 'Планы SKU магазинов'
        unique_together = ('store', 'sku')

    def __str__(self):
        return f'{self.store.name} — {self.sku.name}'


class Visit(models.Model):
    STATUS_CHOICES = [
        ('in_progress', 'В процессе'),
        ('completed', 'Завершён'),
    ]

    supervisor = models.ForeignKey(User, on_delete=models.CASCADE, related_name='visits', verbose_name='Супервайзер')
    store = models.ForeignKey(Store, on_delete=models.CASCADE, related_name='visits', verbose_name='Магазин')
    date = models.DateField(auto_now_add=True, verbose_name='Дата')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='in_progress', verbose_name='Статус')
    completion_pct = models.FloatField(null=True, blank=True, verbose_name='% выполнения')
    notes = models.TextField(blank=True, verbose_name='Заметки')

    class Meta:
        verbose_name = 'Визит'
        verbose_name_plural = 'Визиты'
        ordering = ['-date']

    def __str__(self):
        return f'{self.supervisor.get_full_name() or self.supervisor.username} — {self.store.name} ({self.date})'


class VisitPhoto(models.Model):
    visit = models.ForeignKey(Visit, on_delete=models.CASCADE, related_name='photos', verbose_name='Визит')
    image = models.ImageField(upload_to='photos/%Y/%m/%d/', verbose_name='Фото')
    comment = models.CharField(max_length=500, blank=True, verbose_name='Комментарий')
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Дата загрузки')

    class Meta:
        verbose_name = 'Фото визита'
        verbose_name_plural = 'Фото визитов'
        ordering = ['created_at']

    def __str__(self):
        return f'Фото для визита #{self.visit_id}'


class VisitCheckItem(models.Model):
    STATUS_CHOICES = [
        ('present', 'Есть'),
        ('absent', 'Нет'),
        ('not_checked', 'Не проверено'),
    ]

    visit = models.ForeignKey(Visit, on_delete=models.CASCADE, related_name='check_items', verbose_name='Визит')
    sku = models.ForeignKey(SKU, on_delete=models.CASCADE, related_name='check_items', verbose_name='SKU')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='not_checked', verbose_name='Статус')

    class Meta:
        verbose_name = 'Элемент чек-листа'
        verbose_name_plural = 'Элементы чек-листа'
        ordering = ['sku__category', 'sku__name']

    def __str__(self):
        return f'{self.sku.name} — {self.get_status_display()}'
