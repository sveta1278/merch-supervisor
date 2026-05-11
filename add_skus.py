import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'merch_supervisor.settings')
django.setup()

from core.models import SKU, Store, StoreSKUPlan

skus_data = [
    {'name': 'Молоко 1л (3.2%)', 'barcode': '4602748001063', 'category': 'Молочные'},
    {'name': 'Хлеб пшеничный', 'barcode': '4602748001064', 'category': 'Хлеб'},
    {'name': 'Йогурт 150г', 'barcode': '4602748001065', 'category': 'Молочные'},
    {'name': 'Масло сливочное 180г', 'barcode': '4602748001066', 'category': 'Молочные'},
    {'name': 'Сыр плавленый', 'barcode': '4602748001067', 'category': 'Молочные'},
]

print("Добавляю SKU:")
for sku_data in skus_data:
    sku, created = SKU.objects.get_or_create(
        barcode=sku_data['barcode'],
        defaults={
            'name': sku_data['name'],
            'category': sku_data['category'],
        }
    )
    status = '✓ Создан' if created else '— уже существует'
    print(f"  {sku.name} ({sku.barcode}) {status}")

print(f"\nВсего SKU в БД: {SKU.objects.count()}")

# Добавляю все SKU в план каждого магазина
print("\nДобавляю SKU в план магазинов:")
stores = Store.objects.all()
skus = SKU.objects.all()

for store in stores:
    for sku in skus:
        plan, created = StoreSKUPlan.objects.get_or_create(
            store=store,
            sku=sku,
        )
    print(f"  {store.name}: {skus.count()} товаров в плане")

print(f"\nВсего планов: {StoreSKUPlan.objects.count()}")
