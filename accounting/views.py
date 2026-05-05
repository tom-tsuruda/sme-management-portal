from django.shortcuts import redirect, render

from .services import (
    calculate_accounting_summary,
    calculate_cashflow_schedule,
    create_payable,
    create_receivable,
    create_sale,
    find_balance_sheet_by_month,
    load_payables,
    load_receivables,
    load_sales,
    save_balance_sheet,
)


SALE_STATUS_CHOICES = [
    "未請求",
    "請求済",
    "入金済",
]

RECEIVABLE_STATUS_CHOICES = [
    "未回収",
    "一部回収",
    "回収済",
]

PAYABLE_STATUS_CHOICES = [
    "未払",
    "支払予定",
    "支払済",
]


def accounting_home(request):
    target_month = request.GET.get("target_month", "").strip()
    summary = calculate_accounting_summary(target_month)

    return render(request, "accounting/accounting_home.html", summary)


def sales_list(request):
    sales = load_sales()

    return render(request, "accounting/sales_list.html", {
        "sales": sales,
    })


def sales_create(request):
    if request.method == "POST":
        sale_date = request.POST.get("sale_date", "").strip()
        customer = request.POST.get("customer", "").strip()
        title = request.POST.get("title", "").strip()
        amount = request.POST.get("amount", "").strip()
        tax_amount = request.POST.get("tax_amount", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "").strip()
        memo = request.POST.get("memo", "").strip()

        if title:
            create_sale(
                sale_date=sale_date,
                customer=customer,
                title=title,
                amount=amount,
                tax_amount=tax_amount,
                due_date=due_date,
                status=status,
                memo=memo,
            )

            return redirect("accounting:sales_list")

    return render(request, "accounting/sales_form.html", {
        "sale": None,
        "status_choices": SALE_STATUS_CHOICES,
        "page_title": "売上登録",
        "submit_label": "登録する",
    })

def sales_edit(request, sale_id):
    sale = find_sale_by_id(sale_id)

    if not sale:
        return render(request, "accounting/sales_form.html", {
            "sale": None,
            "status_choices": SALE_STATUS_CHOICES,
            "page_title": "売上が見つかりません",
            "submit_label": "保存する",
        })

    if request.method == "POST":
        sale_date = request.POST.get("sale_date", "").strip()
        customer = request.POST.get("customer", "").strip()
        title = request.POST.get("title", "").strip()
        amount = request.POST.get("amount", "").strip()
        tax_amount = request.POST.get("tax_amount", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "").strip()
        memo = request.POST.get("memo", "").strip()

        if title:
            update_sale(
                sale_id=sale_id,
                sale_date=sale_date,
                customer=customer,
                title=title,
                amount=amount,
                tax_amount=tax_amount,
                due_date=due_date,
                status=status,
                memo=memo,
            )

            return redirect("accounting:sales_list")

    return render(request, "accounting/sales_form.html", {
        "sale": sale,
        "status_choices": SALE_STATUS_CHOICES,
        "page_title": "売上編集",
        "submit_label": "保存する",
    })

def receivables_list(request):
    receivables = load_receivables()

    return render(request, "accounting/receivables_list.html", {
        "receivables": receivables,
    })


def receivable_create(request):
    if request.method == "POST":
        invoice_date = request.POST.get("invoice_date", "").strip()
        customer = request.POST.get("customer", "").strip()
        title = request.POST.get("title", "").strip()
        amount = request.POST.get("amount", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "").strip()
        collected_date = request.POST.get("collected_date", "").strip()
        memo = request.POST.get("memo", "").strip()

        if title:
            create_receivable(
                invoice_date=invoice_date,
                customer=customer,
                title=title,
                amount=amount,
                due_date=due_date,
                status=status,
                collected_date=collected_date,
                memo=memo,
            )

            return redirect("accounting:receivables_list")

    return render(request, "accounting/receivable_form.html", {
        "status_choices": RECEIVABLE_STATUS_CHOICES,
    })


def payables_list(request):
    payables = load_payables()

    return render(request, "accounting/payables_list.html", {
        "payables": payables,
    })


def payable_create(request):
    if request.method == "POST":
        purchase_date = request.POST.get("purchase_date", "").strip()
        vendor = request.POST.get("vendor", "").strip()
        title = request.POST.get("title", "").strip()
        amount = request.POST.get("amount", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "").strip()
        paid_date = request.POST.get("paid_date", "").strip()
        memo = request.POST.get("memo", "").strip()

        if title:
            create_payable(
                purchase_date=purchase_date,
                vendor=vendor,
                title=title,
                amount=amount,
                due_date=due_date,
                status=status,
                paid_date=paid_date,
                memo=memo,
            )

            return redirect("accounting:payables_list")

    return render(request, "accounting/payable_form.html", {
        "status_choices": PAYABLE_STATUS_CHOICES,
    })


def balance_sheet_edit(request):
    target_month = request.GET.get("target_month", "").strip()

    if not target_month:
        target_month = request.POST.get("target_month", "").strip()

    balance_sheet = find_balance_sheet_by_month(target_month) if target_month else None

    if request.method == "POST":
        target_month = request.POST.get("target_month", "").strip()
        cash_balance = request.POST.get("cash_balance", "").strip()
        inventory_amount = request.POST.get("inventory_amount", "").strip()
        fixed_assets = request.POST.get("fixed_assets", "").strip()
        loan_balance = request.POST.get("loan_balance", "").strip()
        memo = request.POST.get("memo", "").strip()

        if target_month:
            save_balance_sheet(
                target_month=target_month,
                cash_balance=cash_balance,
                inventory_amount=inventory_amount,
                fixed_assets=fixed_assets,
                loan_balance=loan_balance,
                memo=memo,
            )

            return redirect(f"/accounting/?target_month={target_month}")

    return render(request, "accounting/balance_sheet_form.html", {
        "balance_sheet": balance_sheet,
        "target_month": target_month,
    })

def cashflow_schedule(request):
    target_month = request.GET.get("target_month", "").strip()
    context = calculate_cashflow_schedule(target_month)

    return render(request, "accounting/cashflow_schedule.html", context)