from django.urls import path

from . import views

app_name = "organizations"

urlpatterns = [
    path("", views.organization_home, name="organization_home"),

    path("departments/", views.department_list, name="department_list"),
    path("departments/create/", views.department_create, name="department_create"),
    path("departments/<str:department_id>/edit/", views.department_edit, name="department_edit"),
    path("departments/<str:department_id>/delete/", views.department_delete, name="department_delete"),

    path("positions/", views.position_list, name="position_list"),
    path("positions/create/", views.position_create, name="position_create"),
    path("positions/<str:position_id>/edit/", views.position_edit, name="position_edit"),
    path("positions/<str:position_id>/delete/", views.position_delete, name="position_delete"),

    path("employees/", views.employee_list, name="employee_list"),
    path("employees/create/", views.employee_create, name="employee_create"),
    path("employees/<str:employee_id>/edit/", views.employee_edit, name="employee_edit"),
    path("employees/<str:employee_id>/delete/", views.employee_delete, name="employee_delete"),
    path("employees/<str:employee_id>/", views.employee_detail, name="employee_detail"),
]