from django.urls import path

from . import views

app_name = "workflows"

urlpatterns = [
    path("", views.request_list, name="request_list"),
    path("create/", views.request_create, name="request_create"),
    path("<str:request_id>/", views.request_detail, name="request_detail"),
    path("<str:request_id>/update-status/", views.update_status, name="update_status"),
    path("<str:request_id>/upload-attachment/", views.upload_attachment, name="upload_attachment"),
]