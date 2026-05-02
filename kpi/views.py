from django.shortcuts import render

# Create your views here.
from django.shortcuts import render

from .services import (
    find_manufacturing_kpi_by_id,
    find_monthly_kpi_by_id,
    get_latest_manufacturing_kpi,
    get_latest_monthly_kpi,
    load_manufacturing_kpis,
    load_monthly_kpis,
)


def kpi_home(request):
    latest_monthly_kpi = get_latest_monthly_kpi()
    latest_manufacturing_kpi = get_latest_manufacturing_kpi()

    monthly_kpis = load_monthly_kpis()
    manufacturing_kpis = load_manufacturing_kpis()

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
    kpis = load_monthly_kpis()

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []

        for row in kpis:
            target_text = " ".join([
                str(row.get("id", "")),
                str(row.get("year_month", "")),
                str(row.get("comment", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(row)

        kpis = filtered

    return render(request, "kpi/monthly_kpi_list.html", {
        "kpis": kpis,
        "keyword": keyword,
    })


def monthly_kpi_detail(request, kpi_id):
    kpi = find_monthly_kpi_by_id(kpi_id)

    return render(request, "kpi/monthly_kpi_detail.html", {
        "kpi": kpi,
    })


def manufacturing_kpi_list(request):
    kpis = load_manufacturing_kpis()

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []

        for row in kpis:
            target_text = " ".join([
                str(row.get("id", "")),
                str(row.get("year_month", "")),
                str(row.get("department", "")),
                str(row.get("comment", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(row)

        kpis = filtered

    return render(request, "kpi/manufacturing_kpi_list.html", {
        "kpis": kpis,
        "keyword": keyword,
    })


def manufacturing_kpi_detail(request, kpi_id):
    kpi = find_manufacturing_kpi_by_id(kpi_id)

    return render(request, "kpi/manufacturing_kpi_detail.html", {
        "kpi": kpi,
    })