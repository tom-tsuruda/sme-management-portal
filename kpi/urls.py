from django.urls import path

from . import views

app_name = "kpi"

urlpatterns = [
    path("", views.kpi_home, name="kpi_home"),
    path("monthly/", views.monthly_kpi_list, name="monthly_kpi_list"),
    path("monthly/<str:kpi_id>/", views.monthly_kpi_detail, name="monthly_kpi_detail"),
    path("manufacturing/", views.manufacturing_kpi_list, name="manufacturing_kpi_list"),
    path("manufacturing/<str:kpi_id>/", views.manufacturing_kpi_detail, name="manufacturing_kpi_detail"),
]