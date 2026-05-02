from django.urls import path

from . import views

app_name = "organizations"

urlpatterns = [
    path("", views.organization_home, name="organization_home"),
    path("departments/", views.department_list, name="department_list"),
    path("positions/", views.position_list, name="position_list"),
    path("employees/", views.employee_list, name="employee_list"),
    path("employees/<str:employee_id>/", views.employee_detail, name="employee_detail"),
]