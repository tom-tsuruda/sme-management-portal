from django.urls import path

from . import views

app_name = "documents"

urlpatterns = [
    path("", views.document_list, name="document_list"),
    path("<str:document_id>/", views.document_detail, name="document_detail"),
    path("<str:document_id>/template/", views.download_template, name="download_template"),
    path("<str:document_id>/completed/", views.download_completed_document, name="download_completed_document"),
    path("<str:document_id>/upload-completed/", views.upload_completed_document, name="upload_completed_document"),
    path("<str:document_id>/update-status/", views.update_status, name="update_status"),
    path("<str:document_id>/create-task/", views.create_task_from_document, name="create_task_from_document"),
]