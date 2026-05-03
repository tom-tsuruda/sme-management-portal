from django.urls import path

from . import views

app_name = "tasks"

urlpatterns = [
    path("", views.task_list, name="task_list"),
    path("create/", views.create_task, name="create_task"),
    path("update-status/", views.update_status, name="update_status"),
    path("<str:task_id>/edit/", views.edit_task, name="edit_task"),
    path("<str:task_id>/", views.task_detail, name="task_detail"),
    path("<str:task_id>/update-status/", views.update_status_from_detail, name="update_status_from_detail"),
]