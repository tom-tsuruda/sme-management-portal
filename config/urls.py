from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path("", include("dashboard.urls")),
    path("admin/", admin.site.urls),
    path("documents/", include("documents.urls")),
    path("tasks/", include("tasks.urls")),
    path("questionnaires/", include("questionnaires.urls")),
]