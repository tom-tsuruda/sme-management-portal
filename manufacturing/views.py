from django.shortcuts import redirect, render

from notifications.services import create_notification

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


def get_notification_priority(item, status):
    risk_level = item.get("risk_level", "")

    if status in ["未整備", "要改善", "期限超過"]:
        return "高"

    if risk_level == "高":
        return "高"

    if status == "要確認":
        return "中"

    return "低"


def should_create_status_notification(status):
    return status in [
        "未整備",
        "要確認",
        "要改善",
        "期限超過",
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
            update_result = update_management_item_status(item_id, status)

            if update_result:
                item = find_management_item_by_id(item_id)

                if item and should_create_status_notification(status):
                    priority = get_notification_priority(item, status)

                    create_notification(
                        title="製造管理項目の状態が更新されました",
                        message=(
                            f"製造管理項目「{item.get('item_name', '')}」の状態が"
                            f"「{status}」に更新されました。"
                            f"領域：{item.get('area', '')}。"
                            f"担当部署：{item.get('owner_department', '')}。"
                            f"リスク：{item.get('risk_level', '')}。"
                        ),
                        target_user=item.get("owner", "") or item.get("owner_department", "") or "製造責任者",
                        category="製造管理",
                        priority=priority,
                        related_type="manufacturing",
                        related_id=item_id,
                    )

    return redirect("manufacturing:management_item_detail", item_id=item_id)


def create_task_from_management_item(request, item_id):
    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        owner = request.POST.get("owner", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        priority = request.POST.get("priority", "中").strip()

        task_id = create_management_task(
            item_id=item_id,
            task_name=task_name,
            owner=owner,
            due_date=due_date,
            priority=priority,
        )

        item = find_management_item_by_id(item_id)

        if task_id and item:
            create_notification(
                title="製造管理タスクが作成されました",
                message=(
                    f"製造管理項目「{item.get('item_name', '')}」から"
                    f"タスクが作成されました。"
                    f"タスクID：{task_id}。"
                    f"担当者：{owner or item.get('owner', '')}。"
                    f"期限：{due_date or '未設定'}。"
                ),
                target_user=owner or item.get("owner", "") or "製造責任者",
                category="製造管理",
                priority=priority or item.get("risk_level", "中"),
                related_type="tasks",
                related_id=task_id,
            )

    return redirect("manufacturing:management_item_detail", item_id=item_id)