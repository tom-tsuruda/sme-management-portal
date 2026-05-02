from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path("", include("dashboard.urls")),
    path("admin/", admin.site.urls),
    path("documents/", include("documents.urls")),
    path("tasks/", include("tasks.urls")),
    path("questionnaires/", include("questionnaires.urls")),
    path("workflows/", include("workflows.urls")),
    path("expenses/", include("expenses.urls")),
    path("organizations/", include("organizations.urls")),
    path("manufacturing/", include("manufacturing.urls")),
    path("kpi/", include("kpi.urls")),
    path("governance/", include("governance.urls")),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)