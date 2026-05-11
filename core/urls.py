from django.urls import path
from . import views

urlpatterns = [
    path('register/', views.register, name='register'),
    path('', views.dashboard, name='dashboard'),
    path('stores/', views.store_list, name='store_list'),
    path('visits/', views.visit_list, name='visit_list'),
    path('visits/new/', views.visit_new, name='visit_new'),
    path('visits/<int:pk>/', views.visit_detail, name='visit_detail'),
    path('visits/<int:pk>/report/', views.visit_report, name='visit_report'),
    path('analytics/', views.analytics, name='analytics'),
]
