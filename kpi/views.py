from django.shortcuts import redirect, render

from .services import (
    attach_manufacturing_kpi_comparisons,
    attach_monthly_kpi_comparisons,
    create_manufacturing_kpi,
    create_monthly_kpi,
    find_manufacturing_kpi_by_id_with_comparison,
    find_monthly_kpi_by_id_with_comparison,
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


def to_number(value):
    text = str(value or "").replace(",", "").replace("%", "").strip()

    if text == "":
        return 0

    try:
        return float(text)
    except Exception:
        return 0


def format_amount(value):
    try:
        number = float(value)

        if number.is_integer():
            return f"{int(number):,}"

        return f"{number:,.1f}".rstrip("0").rstrip(".")
    except Exception:
        return ""


def format_percent(value):
    try:
        number = float(value)
        return f"{number:,.1f}%"
    except Exception:
        return ""


def safe_rate(numerator, denominator):
    denominator = to_number(denominator)

    if denominator == 0:
        return 0

    return to_number(numerator) / denominator * 100


def average(values):
    numbers = [to_number(value) for value in values if str(value or "").strip() != ""]

    if not numbers:
        return 0

    return sum(numbers) / len(numbers)


def filter_by_period(rows, start_month, end_month):
    if not start_month or not end_month:
        return rows

    if start_month > end_month:
        start_month, end_month = end_month, start_month

    return [
        row for row in rows
        if start_month <= str(row.get("year_month", "")) <= end_month
    ]


def build_period_label(start_month, end_month):
    if not start_month and not end_month:
        return ""

    if start_month == end_month:
        return start_month

    return f"{start_month}〜{end_month}"


def build_monthly_period_summary(rows, period_label):
    if not rows:
        return None

    if len(rows) == 1:
        summary = rows[0].copy()
        summary["year_month"] = period_label or summary.get("year_month", "")
        return summary

    sales_amount = sum(to_number(row.get("sales_amount", "")) for row in rows)
    gross_profit = sum(to_number(row.get("gross_profit", "")) for row in rows)
    operating_profit = sum(to_number(row.get("operating_profit", "")) for row in rows)

    latest_row = sorted(rows, key=lambda row: str(row.get("year_month", "")))[-1]
    cash_balance = to_number(latest_row.get("cash_balance", ""))

    gross_profit_rate = safe_rate(gross_profit, sales_amount)
    operating_profit_rate = safe_rate(operating_profit, sales_amount)

    return {
        "id": "",
        "year_month": period_label,
        "sales_amount": sales_amount,
        "gross_profit": gross_profit,
        "operating_profit": operating_profit,
        "cash_balance": cash_balance,
        "sales_amount_display": format_amount(sales_amount),
        "gross_profit_display": format_amount(gross_profit),
        "operating_profit_display": format_amount(operating_profit),
        "cash_balance_display": format_amount(cash_balance),
        "gross_profit_rate_display": format_percent(gross_profit_rate),
        "operating_profit_rate_display": format_percent(operating_profit_rate),
    }


def build_manufacturing_period_summary(rows, period_label):
    if not rows:
        return None

    if len(rows) == 1:
        summary = rows[0].copy()
        summary["year_month"] = period_label or summary.get("year_month", "")
        return summary

    production_volume = sum(to_number(row.get("production_volume", "")) for row in rows)
    defect_count = sum(to_number(row.get("defect_count", "")) for row in rows)
    downtime_hours = sum(to_number(row.get("downtime_hours", "")) for row in rows)
    accident_count = sum(to_number(row.get("accident_count", "")) for row in rows)
    near_miss_count = sum(to_number(row.get("near_miss_count", "")) for row in rows)
    quality_claim_count = sum(to_number(row.get("quality_claim_count", "")) for row in rows)
    energy_usage = sum(to_number(row.get("energy_usage", "")) for row in rows)

    defect_rate = safe_rate(defect_count, production_volume)
    yield_rate = average(row.get("yield_rate", "") for row in rows)
    on_time_delivery_rate = average(row.get("on_time_delivery_rate", "") for row in rows)
    equipment_availability = average(row.get("equipment_availability", "") for row in rows)

    return {
        "id": "",
        "year_month": period_label,
        "department": "全体",
        "production_volume": production_volume,
        "defect_count": defect_count,
        "defect_rate": defect_rate,
        "yield_rate": yield_rate,
        "on_time_delivery_rate": on_time_delivery_rate,
        "equipment_availability": equipment_availability,
        "downtime_hours": downtime_hours,
        "accident_count": accident_count,
        "near_miss_count": near_miss_count,
        "quality_claim_count": quality_claim_count,
        "energy_usage": energy_usage,
        "production_volume_display": format_amount(production_volume),
        "defect_count_display": format_amount(defect_count),
        "defect_rate_display": format_percent(defect_rate),
        "yield_rate_display": format_percent(yield_rate),
        "on_time_delivery_rate_display": format_percent(on_time_delivery_rate),
        "equipment_availability_display": format_percent(equipment_availability),
        "downtime_hours_display": format_amount(downtime_hours),
        "accident_count_display": format_amount(accident_count),
        "near_miss_count_display": format_amount(near_miss_count),
        "quality_claim_count_display": format_amount(quality_claim_count),
        "energy_usage_display": format_amount(energy_usage),
    }


def kpi_home(request):
    monthly_kpis = attach_monthly_kpi_comparisons(load_monthly_kpis())
    manufacturing_kpis = attach_manufacturing_kpi_comparisons(load_manufacturing_kpis())

    available_months = sorted(
        {
            str(row.get("year_month", ""))
            for row in monthly_kpis + manufacturing_kpis
            if str(row.get("year_month", "")).strip()
        }
    )

    latest_month = available_months[-1] if available_months else ""

    start_month = request.GET.get("start_month", "").strip() or latest_month
    end_month = request.GET.get("end_month", "").strip() or latest_month

    if start_month and end_month and start_month > end_month:
        start_month, end_month = end_month, start_month

    period_label = build_period_label(start_month, end_month)

    monthly_display_rows = filter_by_period(monthly_kpis, start_month, end_month)
    manufacturing_display_rows = filter_by_period(manufacturing_kpis, start_month, end_month)

    latest_monthly_kpi = build_monthly_period_summary(
        monthly_display_rows,
        period_label,
    )
    latest_manufacturing_kpi = build_manufacturing_period_summary(
        manufacturing_display_rows,
        period_label,
    )

    return render(request, "kpi/kpi_home.html", {
        "latest_monthly_kpi": latest_monthly_kpi,
        "latest_manufacturing_kpi": latest_manufacturing_kpi,
        "monthly_kpi_count": len(monthly_kpis),
        "manufacturing_kpi_count": len(manufacturing_kpis),
        "monthly_kpis": monthly_display_rows,
        "manufacturing_kpis": manufacturing_display_rows,
        "available_months": available_months,
        "start_month": start_month,
        "end_month": end_month,
        "period_label": period_label,
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