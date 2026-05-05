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
    update_manufacturing_kpi,
    update_monthly_kpi,
)


def find_by_year_month(rows, year_month):
    if not year_month:
        return None

    for row in rows:
        if str(row.get("year_month", "")) == str(year_month):
            return row

    return None


def kpi_home(request):
    selected_month = request.GET.get("selected_month", "").strip()

    monthly_kpis = attach_monthly_kpi_comparisons(load_monthly_kpis())
    manufacturing_kpis = attach_manufacturing_kpi_comparisons(load_manufacturing_kpis())

    available_months = sorted(
        {
            str(row.get("year_month", ""))
            for row in monthly_kpis + manufacturing_kpis
            if str(row.get("year_month", "")).strip()
        },
        reverse=True,
    )

    if selected_month:
        latest_monthly_kpi = find_by_year_month(monthly_kpis, selected_month)
        latest_manufacturing_kpi = find_by_year_month(manufacturing_kpis, selected_month)
    else:
        latest_monthly_kpi = get_latest_monthly_kpi_with_comparison()
        latest_manufacturing_kpi = get_latest_manufacturing_kpi_with_comparison()

        if latest_monthly_kpi:
            selected_month = latest_monthly_kpi.get("year_month", "")
        elif latest_manufacturing_kpi:
            selected_month = latest_manufacturing_kpi.get("year_month", "")

    monthly_display_rows = monthly_kpis
    manufacturing_display_rows = manufacturing_kpis

    if selected_month:
        monthly_display_rows = [
            row for row in monthly_kpis
            if str(row.get("year_month", "")) == selected_month
        ]
        manufacturing_display_rows = [
            row for row in manufacturing_kpis
            if str(row.get("year_month", "")) == selected_month
        ]

    return render(request, "kpi/kpi_home.html", {
        "latest_monthly_kpi": latest_monthly_kpi,
        "latest_manufacturing_kpi": latest_manufacturing_kpi,
        "monthly_kpi_count": len(monthly_kpis),
        "manufacturing_kpi_count": len(manufacturing_kpis),
        "monthly_kpis": monthly_display_rows,
        "manufacturing_kpis": manufacturing_display_rows,
        "available_months": available_months,
        "selected_month": selected_month,
    })


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

    return render(request, "kpi/monthly_kpi_form.html", {
        "kpi": None,
        "page_title": "月次経営KPI登録",
        "submit_label": "登録する",
    })


def monthly_kpi_edit(request, kpi_id):
    kpi = find_monthly_kpi_by_id_with_comparison(kpi_id)

    if not kpi:
        return render(request, "kpi/monthly_kpi_detail.html", {
            "kpi": None,
        })

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
            update_monthly_kpi(
                kpi_id=kpi_id,
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

    return render(request, "kpi/monthly_kpi_form.html", {
        "kpi": kpi,
        "page_title": "月次経営KPI編集",
        "submit_label": "保存する",
    })


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

    return render(request, "kpi/manufacturing_kpi_form.html", {
        "kpi": None,
        "page_title": "製造KPI登録",
        "submit_label": "登録する",
    })


def manufacturing_kpi_edit(request, kpi_id):
    kpi = find_manufacturing_kpi_by_id_with_comparison(kpi_id)

    if not kpi:
        return render(request, "kpi/manufacturing_kpi_detail.html", {
            "kpi": None,
        })

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
            update_manufacturing_kpi(
                kpi_id=kpi_id,
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

    return render(request, "kpi/manufacturing_kpi_form.html", {
        "kpi": kpi,
        "page_title": "製造KPI編集",
        "submit_label": "保存する",
    })