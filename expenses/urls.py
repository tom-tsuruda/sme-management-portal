from django.urls import path

from . import views

app_name = "expenses"

urlpatterns = [
    path("", views.expense_list, name="expense_list"),
    path("create/", views.expense_create, name="expense_create"),
    path("export-csv/", views.export_csv, name="export_csv"),
    path("<str:expense_id>/", views.expense_detail, name="expense_detail"),
    path("<str:expense_id>/update-status/", views.update_status, name="update_status"),
    path("<str:expense_id>/upload-attachment/", views.upload_attachment, name="upload_attachment"),
]