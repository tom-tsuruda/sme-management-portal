from django.shortcuts import redirect, render

from .services import (
    attach_manufacturing_kpi_comparisons,
    attach_monthly_kpi_comparisons,
    create_manufacturing_kpi,
    create_monthly_kpi,
    find_manufacturing_kpi_by_id_with_comparison,
    find_monthly_kpi_by_id_with_comparison,
    get_latest_manufacturing_kpi_with_comparison,
    get_latest_monthly_kpi_with_comparison,
    load_manufacturing_kpis,
    load_monthly_kpis,
)


def kpi_home(request):
    latest_monthly_kpi = get_latest_monthly_kpi_with_comparison()
    latest_manufacturing_kpi = get_latest_manufacturing_kpi_with_comparison()

    monthly_kpis = attach_monthly_kpi_comparisons(load_monthly_kpis())
    manufacturing_kpis = attach_manufacturing_kpi_comparisons(load_manufacturing_kpis())

    context = {
        "latest_monthly_kpi": latest_monthly_kpi,
        "latest_manufacturing_kpi": latest_manufacturing_kpi,
        "monthly_kpi_count": len(monthly_kpis),
        "manufacturing_kpi_count": len(manufacturing_kpis),
        "monthly_kpis": monthly_kpis[:5],
        "manufacturing_kpis": manufacturing_kpis[:5],
    }

    return render(request, "kpi/kpi_home.html", context)


def monthly_kpi_list(request):
    kpis = attach_monthly_kpi_comparisons(load_monthly_kpis())

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []

        for row in kpis:
            target_text = " ".join([
                str(row.get("id", "")),
                str(row.get("year_month", "")),
                str(row.get("comment", "")),
                str(row.get("comparison_comment", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(row)

        kpis = filtered

    return render(request, "kpi/monthly_kpi_list.html", {
        "kpis": kpis,
        "keyword": keyword,
    })


def monthly_kpi_detail(request, kpi_id):
    kpi = find_monthly_kpi_by_id_with_comparison(kpi_id)

    return render(request, "kpi/monthly_kpi_detail.html", {
        "kpi": kpi,
    })


def monthly_kpi_create(request):
    if request.method == "POST":
        year_month = request.POST.get("year_month", "").strip()
        sales_amount = request.POST.get("sales_amount", "").strip()
        gross_profit = request.POST.get("gross_profit", "").strip()
        operating_profit = request.POST.get("operating_profit", "").strip()
        cash_balance = request.POST.get("cash_balance", "").strip()
        accounts_receivable = request.POST.get("accounts_receivable", "").strip()
        accounts_payable = request.POST.get("accounts_payable", "").strip()
        expense_total = request.POST.get("expense_total", "").strip()
        new_orders = request.POST.get("new_orders", "").strip()
        order_backlog = request.POST.get("order_backlog", "").strip()
        comment = request.POST.get("comment", "").strip()

        if year_month:
            kpi_id = create_monthly_kpi(
                year_month=year_month,
                sales_amount=sales_amount,
                gross_profit=gross_profit,
                operating_profit=operating_profit,
                cash_balance=cash_balance,
                accounts_receivable=accounts_receivable,
                accounts_payable=accounts_payable,
                expense_total=expense_total,
                new_orders=new_orders,
                order_backlog=order_backlog,
                comment=comment,
            )

            return redirect("kpi:monthly_kpi_detail", kpi_id=kpi_id)

    return render(request, "kpi/monthly_kpi_form.html")


def manufacturing_kpi_list(request):
    kpis = attach_manufacturing_kpi_comparisons(load_manufacturing_kpis())

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []

        for row in kpis:
            target_text = " ".join([
                str(row.get("id", "")),
                str(row.get("year_month", "")),
                str(row.get("department", "")),
                str(row.get("comment", "")),
                str(row.get("comparison_comment", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(row)

        kpis = filtered

    return render(request, "kpi/manufacturing_kpi_list.html", {
        "kpis": kpis,
        "keyword": keyword,
    })


def manufacturing_kpi_detail(request, kpi_id):
    kpi = find_manufacturing_kpi_by_id_with_comparison(kpi_id)

    return render(request, "kpi/manufacturing_kpi_detail.html", {
        "kpi": kpi,
    })


def manufacturing_kpi_create(request):
    if request.method == "POST":
        year_month = request.POST.get("year_month", "").strip()
        department = request.POST.get("department", "").strip()
        production_volume = request.POST.get("production_volume", "").strip()
        defect_count = request.POST.get("defect_count", "").strip()
        defect_rate = request.POST.get("defect_rate", "").strip()
        yield_rate = request.POST.get("yield_rate", "").strip()
        on_time_delivery_rate = request.POST.get("on_time_delivery_rate", "").strip()
        equipment_availability = request.POST.get("equipment_availability", "").strip()
        downtime_hours = request.POST.get("downtime_hours", "").strip()
        accident_count = request.POST.get("accident_count", "").strip()
        near_miss_count = request.POST.get("near_miss_count", "").strip()
        quality_claim_count = request.POST.get("quality_claim_count", "").strip()
        energy_usage = request.POST.get("energy_usage", "").strip()
        comment = request.POST.get("comment", "").strip()

        if year_month:
            kpi_id = create_manufacturing_kpi(
                year_month=year_month,
                department=department,
                production_volume=production_volume,
                defect_count=defect_count,
                defect_rate=defect_rate,
                yield_rate=yield_rate,
                on_time_delivery_rate=on_time_delivery_rate,
                equipment_availability=equipment_availability,
                downtime_hours=downtime_hours,
                accident_count=accident_count,
                near_miss_count=near_miss_count,
                quality_claim_count=quality_claim_count,
                energy_usage=energy_usage,
                comment=comment,
            )

            return redirect("kpi:manufacturing_kpi_detail", kpi_id=kpi_id)

    return render(request, "kpi/manufacturing_kpi_form.html")