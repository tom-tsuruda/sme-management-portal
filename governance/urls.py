from django.urls import path

from . import views

app_name = "governance"

urlpatterns = [
    path("", views.governance_item_list, name="governance_item_list"),
    path("<str:item_id>/", views.governance_item_detail, name="governance_item_detail"),
    path("<str:item_id>/update-status/", views.update_status, name="update_status"),
    path("<str:item_id>/create-task/", views.create_task_from_governance_item, name="create_task_from_governance_item"),
]