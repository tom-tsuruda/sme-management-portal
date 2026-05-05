from django.urls import path

from . import views

app_name = "accounting"

urlpatterns = [
    path("", views.accounting_home, name="accounting_home"),

    path("sales/", views.sales_list, name="sales_list"),
    path("sales/create/", views.sales_create, name="sales_create"),
    path("sales/<str:sale_id>/edit/", views.sales_edit, name="sales_edit"),

    path("receivables/", views.receivables_list, name="receivables_list"),
    path("receivables/create/", views.receivable_create, name="receivable_create"),
    path("receivables/<str:receivable_id>/edit/", views.receivable_edit, name="receivable_edit"),

    path("payables/", views.payables_list, name="payables_list"),
    path("payables/create/", views.payable_create, name="payable_create"),
    path("payables/<str:payable_id>/edit/", views.payable_edit, name="payable_edit"),

    path("balance-sheet/", views.balance_sheet_edit, name="balance_sheet_edit"),
    path("cashflow/", views.cashflow_schedule, name="cashflow_schedule"),
]