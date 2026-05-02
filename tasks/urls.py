from django.urls import path

from . import views

app_name = "tasks"

urlpatterns = [
    path("", views.task_list, name="task_list"),
    path("update-status/", views.update_status, name="update_status"),
    path("create/", views.create_task, name="create_task"),
]