from django.urls import path

from . import views

app_name = "notifications"

urlpatterns = [
    path("", views.notification_list, name="notification_list"),
    path("<str:notification_id>/", views.notification_detail, name="notification_detail"),
    path("<str:notification_id>/mark-read/", views.mark_read, name="mark_read"),
    path("<str:notification_id>/mark-unread/", views.mark_unread, name="mark_unread"),
]