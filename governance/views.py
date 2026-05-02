from django.shortcuts import render

# Create your views here.
from django.shortcuts import redirect, render

from .services import (
    create_task_from_governance_item as create_governance_task,
    find_governance_item_by_id,
    load_governance_items,
    update_governance_item_status,
)


GOVERNANCE_STATUS_CHOICES = [
    "未整備",
    "要確認",
    "確認中",
    "整備済",
    "要改定",
    "期限超過",
    "対象外",
]


def governance_item_list(request):
    all_items = load_governance_items()
    items = all_items

    keyword = request.GET.get("q", "").strip()
    category = request.GET.get("category", "").strip()
    status = request.GET.get("status", "").strip()

    if keyword:
        filtered_items = []

        for item in items:
            target_text = " ".join([
                str(item.get("id", "")),
                str(item.get("category", "")),
                str(item.get("item_name", "")),
                str(item.get("required_level", "")),
                str(item.get("owner_department", "")),
                str(item.get("owner", "")),
                str(item.get("status", "")),
                str(item.get("risk_level", "")),
                str(item.get("remarks", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_items.append(item)

        items = filtered_items

    if category:
        items = [
            item for item in items
            if str(item.get("category", "")) == category
        ]

    if status:
        items = [
            item for item in items
            if str(item.get("status", "")) == status
        ]

    categories = sorted(set([
        item.get("category", "")
        for item in all_items
        if item.get("category", "")
    ]))

    not_ready_count = len([
        item for item in all_items
        if item.get("status") == "未整備"
    ])

    need_check_count = len([
        item for item in all_items
        if item.get("status") == "要確認"
    ])

    need_revision_count = len([
        item for item in all_items
        if item.get("status") == "要改定"
    ])

    overdue_count = len([
        item for item in all_items
        if item.get("status") == "期限超過"
    ])

    high_risk_count = len([
        item for item in all_items
        if item.get("risk_level") == "高"
    ])

    required_count = len([
        item for item in all_items
        if item.get("required_level") in ["法定必須", "法定必須級"]
    ])

    context = {
        "items": items,
        "keyword": keyword,
        "category": category,
        "status": status,
        "categories": categories,
        "status_choices": GOVERNANCE_STATUS_CHOICES,
        "total_count": len(all_items),
        "display_count": len(items),
        "not_ready_count": not_ready_count,
        "need_check_count": need_check_count,
        "need_revision_count": need_revision_count,
        "overdue_count": overdue_count,
        "high_risk_count": high_risk_count,
        "required_count": required_count,
    }

    return render(request, "governance/governance_item_list.html", context)


def governance_item_detail(request, item_id):
    item = find_governance_item_by_id(item_id)

    return render(request, "governance/governance_item_detail.html", {
        "item": item,
        "status_choices": GOVERNANCE_STATUS_CHOICES,
    })


def update_status(request, item_id):
    if request.method == "POST":
        status = request.POST.get("status", "").strip()

        if status:
            update_governance_item_status(item_id, status)

    return redirect("governance:governance_item_detail", item_id=item_id)


def create_task_from_governance_item(request, item_id):
    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        owner = request.POST.get("owner", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        priority = request.POST.get("priority", "中").strip()

        create_governance_task(
            item_id=item_id,
            task_name=task_name,
            owner=owner,
            due_date=due_date,
            priority=priority,
        )

    return redirect("governance:governance_item_detail", item_id=item_id)