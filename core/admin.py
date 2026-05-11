from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from .models import User, Store, SKU, StoreSKUPlan, Visit, VisitPhoto, VisitCheckItem


@admin.register(User)
class UserAdmin(BaseUserAdmin):
    fieldsets = BaseUserAdmin.fieldsets + (
        ('Роль', {'fields': ('role',)}),
    )
    list_display = ('username', 'email', 'first_name', 'last_name', 'role', 'is_active')


class StoreSKUPlanInline(admin.TabularInline):
    model = StoreSKUPlan
    extra = 1
    autocomplete_fields = ['sku']


@admin.register(Store)
class StoreAdmin(admin.ModelAdmin):
    list_display = ('name', 'network', 'address', 'is_active')
    list_filter = ('is_active', 'network')
    search_fields = ('name', 'network', 'address')
    inlines = [StoreSKUPlanInline]


@admin.register(SKU)
class SKUAdmin(admin.ModelAdmin):
    list_display = ('name', 'barcode', 'category')
    list_filter = ('category',)
    search_fields = ('name', 'barcode', 'category')


@admin.register(StoreSKUPlan)
class StoreSKUPlanAdmin(admin.ModelAdmin):
    list_display = ('store', 'sku')
    list_filter = ('store',)
    autocomplete_fields = ['store', 'sku']


class VisitPhotoInline(admin.TabularInline):
    model = VisitPhoto
    extra = 0
    readonly_fields = ('created_at',)


class VisitCheckItemInline(admin.TabularInline):
    model = VisitCheckItem
    extra = 0
    autocomplete_fields = ['sku']


@admin.register(Visit)
class VisitAdmin(admin.ModelAdmin):
    list_display = ('supervisor', 'store', 'date', 'status', 'completion_pct')
    list_filter = ('status', 'date', 'store')
    search_fields = ('supervisor__username', 'store__name')
    readonly_fields = ('date', 'completion_pct')
    inlines = [VisitPhotoInline, VisitCheckItemInline]


@admin.register(VisitPhoto)
class VisitPhotoAdmin(admin.ModelAdmin):
    list_display = ('visit', 'comment', 'created_at')
    readonly_fields = ('created_at',)


@admin.register(VisitCheckItem)
class VisitCheckItemAdmin(admin.ModelAdmin):
    list_display = ('visit', 'sku', 'status')
    list_filter = ('status',)
