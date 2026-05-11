# Merch Supervisor

Адаптивный веб-сервис мониторинга мерчандайзинга на Django 5.2.

## Функциональность

- Регистрация и авторизация пользователей (мерчандайзер / администратор)
- Управление торговыми точками и товарными позициями (SKU)
- Создание визитов с автоматической генерацией чек-листа
- Фотофиксация выкладки с галереей
- AJAX-обновление статусов SKU без перезагрузки страницы
- Расчёт процента выполнения плана при завершении визита
- Двухуровневая аналитика с графиками Chart.js:
  - личная статистика для мерчандайзера
  - сводная аналитика по всем сотрудникам для администратора
- Адаптивный интерфейс Bootstrap 5 с нижней навигацией для мобильных устройств

## Стек

| Компонент | Технология |
|---|---|
| Backend | Python 3.12, Django 5.2 |
| БД | PostgreSQL / SQLite (для разработки) |
| Frontend | Bootstrap 5, Chart.js 4 |
| Загрузка фото | Pillow |
| Переменные окружения | python-dotenv |

## Быстрый старт

```bash
# Клонировать репозиторий
git clone https://github.com/sveta1278/merch-supervisor.git
cd merch-supervisor

# Создать виртуальное окружение
python3 -m venv venv
source venv/bin/activate

# Установить зависимости
pip install -r requirements.txt

# Настроить переменные окружения
cp .env.example .env
# Отредактировать .env при необходимости

# Применить миграции
python manage.py migrate

# Создать суперпользователя
python manage.py createsuperuser

# Запустить сервер
python manage.py runserver
```

Открыть в браузере: http://127.0.0.1:8000

## Структура проекта

```
merch_supervisor/
├── core/                  # Основное приложение
│   ├── models.py          # 7 моделей данных
│   ├── views.py           # Представления
│   ├── forms.py           # Формы регистрации и загрузки фото
│   ├── urls.py            # URL-маршруты
│   ├── admin.py           # Django Admin
│   └── templatetags/      # Кастомные фильтры шаблонов
├── templates/             # HTML-шаблоны
├── static/css/            # Стили
├── merch_supervisor/      # Настройки проекта
└── requirements.txt
```

## База данных

| Таблица | Назначение |
|---|---|
| `core_user` | Пользователи (расширяет AbstractUser) |
| `core_store` | Торговые точки |
| `core_sku` | Товарные позиции |
| `core_storeskuplan` | Плановый ассортимент магазина |
| `core_visit` | Визиты мерчандайзеров |
| `core_visitphoto` | Фотографии выкладки |
| `core_visitcheckitem` | Результаты проверки SKU |

## Переменные окружения

Скопировать `.env.example` в `.env` и заполнить:

```
SECRET_KEY=your-secret-key
DEBUG=True
USE_SQLITE=True
```

Для PostgreSQL установить `USE_SQLITE=False` и добавить параметры подключения.
