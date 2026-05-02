from django.shortcuts import redirect, render

from .services import (
    create_task_from_management_item as create_management_task,
    find_management_item_by_id,
    load_management_items,
    update_management_item_status,
)


MANAGEMENT_STATUS_CHOICES = [
    "未整備",
    "要確認",
    "確認中",
    "整備済",
    "要改善",
    "期限超過",
    "対象外",
]


def management_item_list(request):
    all_items = load_management_items()
    items = all_items

    keyword = request.GET.get("q", "").strip()
    area = request.GET.get("area", "").strip()
    status = request.GET.get("status", "").strip()

    if keyword:
        filtered_items = []

        for item in items:
            target_text = " ".join([
                str(item.get("id", "")),
                str(item.get("area", "")),
                str(item.get("category", "")),
                str(item.get("item_name", "")),
                str(item.get("description", "")),
                str(item.get("owner_department", "")),
                str(item.get("owner", "")),
                str(item.get("risk_level", "")),
                str(item.get("status", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_items.append(item)

        items = filtered_items

    if area:
        items = [
            item for item in items
            if str(item.get("area", "")) == area
        ]

    if status:
        items = [
            item for item in items
            if str(item.get("status", "")) == status
        ]

    areas = sorted(set([
        item.get("area", "")
        for item in all_items
        if item.get("area", "")
    ]))

    not_ready_count = len([
        item for item in all_items
        if item.get("status") == "未整備"
    ])

    need_check_count = len([
        item for item in all_items
        if item.get("status") == "要確認"
    ])

    need_improvement_count = len([
        item for item in all_items
        if item.get("status") == "要改善"
    ])

    completed_count = len([
        item for item in all_items
        if item.get("status") == "整備済"
    ])

    overdue_count = len([
        item for item in all_items
        if item.get("status") == "期限超過"
    ])

    high_risk_count = len([
        item for item in all_items
        if item.get("risk_level") == "高"
    ])

    context = {
        "items": items,
        "keyword": keyword,
        "area": area,
        "status": status,
        "areas": areas,
        "status_choices": MANAGEMENT_STATUS_CHOICES,
        "total_count": len(all_items),
        "display_count": len(items),
        "not_ready_count": not_ready_count,
        "need_check_count": need_check_count,
        "need_improvement_count": need_improvement_count,
        "completed_count": completed_count,
        "overdue_count": overdue_count,
        "high_risk_count": high_risk_count,
    }

    return render(request, "manufacturing/management_item_list.html", context)


def management_item_detail(request, item_id):
    item = find_management_item_by_id(item_id)

    return render(request, "manufacturing/management_item_detail.html", {
        "item": item,
        "status_choices": MANAGEMENT_STATUS_CHOICES,
    })


def update_status(request, item_id):
    if request.method == "POST":
        status = request.POST.get("status", "").strip()

        if status:
            update_management_item_status(item_id, status)

    return redirect("manufacturing:management_item_detail", item_id=item_id)


def create_task_from_management_item(request, item_id):
    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        owner = request.POST.get("owner", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        priority = request.POST.get("priority", "中").strip()

        create_management_task(
            item_id=item_id,
            task_name=task_name,
            owner=owner,
            due_date=due_date,
            priority=priority,
        )

    return redirect("manufacturing:management_item_detail", item_id=item_id)