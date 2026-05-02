from django.urls import path

from . import views

app_name = "manufacturing"

urlpatterns = [
    path("", views.management_item_list, name="management_item_list"),
    path("<str:item_id>/", views.management_item_detail, name="management_item_detail"),
    path("<str:item_id>/update-status/", views.update_status, name="update_status"),
    path("<str:item_id>/create-task/", views.create_task_from_management_item, name="create_task_from_management_item"),
]