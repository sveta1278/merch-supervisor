import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'merch_supervisor.settings')
django.setup()

from core.models import Store

stores_data = [
    {'name': 'Декабристов 10', 'address': 'ул. Декабристов д.10', 'network': 'Пятёрочка'},
    {'name': 'Декабристов 12', 'address': 'ул. Декабристов д.12', 'network': 'Пятёрочка'},
    {'name': 'Декабристов 14', 'address': 'ул. Декабристов д.14', 'network': 'Пятёрочка'},
    {'name': 'Декабристов 16', 'address': 'ул. Декабристов д.16', 'network': 'Пятёрочка'},
]

for store_data in stores_data:
    store, created = Store.objects.get_or_create(
        address=store_data['address'],
        defaults={
            'name': store_data['name'],
            'network': store_data['network'],
            'is_active': True,
        }
    )
    status = '✓ Создан' if created else '— уже существует'
    print(f"{store_data['name']} ({store_data['address']}) {status}")

print(f"\nВсего магазинов в БД: {Store.objects.count()}")
