from django.shortcuts import redirect, render

from notifications.services import create_notification

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


def get_notification_priority(item, status):
    risk_level = item.get("risk_level", "")
    required_level = item.get("required_level", "")

    if status in ["未整備", "要改定", "期限超過"]:
        return "高"

    if risk_level == "高":
        return "高"

    if required_level in ["法定必須", "法定必須級"]:
        return "高"

    if status == "要確認":
        return "中"

    return "低"


def should_create_status_notification(status):
    return status in [
        "未整備",
        "要確認",
        "要改定",
        "期限超過",
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
            update_result = update_governance_item_status(item_id, status)

            if update_result:
                item = find_governance_item_by_id(item_id)

                if item and should_create_status_notification(status):
                    priority = get_notification_priority(item, status)

                    create_notification(
                        title="ガバナンス文書の状態が更新されました",
                        message=(
                            f"ガバナンス文書「{item.get('item_name', '')}」の状態が"
                            f"「{status}」に更新されました。"
                            f"カテゴリ：{item.get('category', '')}。"
                            f"重要度：{item.get('required_level', '')}。"
                            f"主管部署：{item.get('owner_department', '')}。"
                            f"リスク：{item.get('risk_level', '')}。"
                        ),
                        target_user=item.get("owner", "") or item.get("owner_department", "") or "管理者",
                        category="ガバナンス",
                        priority=priority,
                        related_type="governance",
                        related_id=item_id,
                    )

    return redirect("governance:governance_item_detail", item_id=item_id)


def create_task_from_governance_item(request, item_id):
    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        owner = request.POST.get("owner", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        priority = request.POST.get("priority", "中").strip()

        task_id = create_governance_task(
            item_id=item_id,
            task_name=task_name,
            owner=owner,
            due_date=due_date,
            priority=priority,
        )

        item = find_governance_item_by_id(item_id)

        if task_id and item:
            create_notification(
                title="ガバナンスタスクが作成されました",
                message=(
                    f"ガバナンス文書「{item.get('item_name', '')}」から"
                    f"タスクが作成されました。"
                    f"タスクID：{task_id}。"
                    f"担当者：{owner or item.get('owner', '')}。"
                    f"期限：{due_date or '未設定'}。"
                ),
                target_user=owner or item.get("owner", "") or "管理者",
                category="ガバナンス",
                priority=priority or item.get("risk_level", "中"),
                related_type="tasks",
                related_id=task_id,
            )

    return redirect("governance:governance_item_detail", item_id=item_id)