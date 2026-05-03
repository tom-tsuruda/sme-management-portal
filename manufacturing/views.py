from django.shortcuts import redirect, render

from notifications.services import create_notification
from organizations.services import find_employee_by_id, load_employees

from .services import (
    create_monitoring_record as create_manufacturing_monitoring_record,
    create_task_from_management_item as create_management_task,
    find_management_item_by_id,
    find_monitoring_records_by_item_id,
    load_management_items,
    update_management_item_status,
    create_incident as create_manufacturing_incident,
    load_incidents,
    find_incident_by_id,
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

MONITORING_RESULT_CHOICES = [
    "OK",
    "要確認",
    "要改善",
    "NG",
    "異常",
]

INCIDENT_SEVERITY_CHOICES = [
    "重大",
    "高",
    "中",
    "低",
]

INCIDENT_STATUS_CHOICES = [
    "未対応",
    "対応中",
    "是正済",
    "再発防止確認中",
    "完了",
    "対象外",
]

INCIDENT_AREA_CHOICES = [
    "品質管理",
    "安全衛生",
    "設備保全",
    "化学物質",
    "環境管理",
    "研究開発",
    "製造管理",
    "IT/OT",
]

def resolve_employee_name(employee_id, fallback=""):
    employee = find_employee_by_id(employee_id) if employee_id else None

    if employee:
        return employee.get("employee_name", "")

    return fallback


def get_notification_priority(item, status):
    risk_level = item.get("risk_level", "")

    if status in ["未整備", "要改善", "期限超過"]:
        return "高"

    if risk_level == "高":
        return "高"

    if status == "要確認":
        return "中"

    return "低"


def get_monitoring_notification_priority(item, result):
    if result in ["要改善", "NG", "異常"]:
        return "高"

    if result == "要確認":
        return "中"

    if item.get("risk_level") == "高":
        return "中"

    return "低"

def get_incident_notification_priority(severity):
    if severity in ["重大", "高"]:
        return "高"

    if severity == "中":
        return "中"

    return "低"


def should_create_incident_notification(severity):
    return severity in ["重大", "高", "中"]

def should_create_status_notification(status):
    return status in [
        "未整備",
        "要確認",
        "要改善",
        "期限超過",
    ]


def should_create_monitoring_notification(result):
    return result in [
        "要確認",
        "要改善",
        "NG",
        "異常",
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
    employees = load_employees()
    monitoring_records = find_monitoring_records_by_item_id(item_id)

    return render(request, "manufacturing/management_item_detail.html", {
        "item": item,
        "employees": employees,
        "monitoring_records": monitoring_records,
        "status_choices": MANAGEMENT_STATUS_CHOICES,
        "monitoring_result_choices": MONITORING_RESULT_CHOICES,
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
        owner_employee_id = request.POST.get("owner_employee_id", "").strip()

        owner = resolve_employee_name(
            owner_employee_id,
            fallback=request.POST.get("owner", "").strip(),
        )

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


def create_monitoring_record(request, item_id):
    if request.method == "POST":
        checked_at = request.POST.get("checked_at", "").strip()
        checked_by_employee_id = request.POST.get("checked_by_employee_id", "").strip()

        checked_by = resolve_employee_name(
            checked_by_employee_id,
            fallback=request.POST.get("checked_by", "").strip(),
        )

        result = request.POST.get("result", "").strip()
        score = request.POST.get("score", "").strip()
        comment = request.POST.get("comment", "").strip()
        evidence_file = request.POST.get("evidence_file", "").strip()
        next_action = request.POST.get("next_action", "").strip()

        record_id = create_manufacturing_monitoring_record(
            management_item_id=item_id,
            checked_at=checked_at,
            checked_by=checked_by,
            result=result,
            score=score,
            comment=comment,
            evidence_file=evidence_file,
            next_action=next_action,
        )

        item = find_management_item_by_id(item_id)

        if record_id and item and should_create_monitoring_notification(result):
            priority = get_monitoring_notification_priority(item, result)

            create_notification(
                title="製造管理の監視記録で要対応が発生しました",
                message=(
                    f"製造管理項目「{item.get('item_name', '')}」の監視記録で"
                    f"結果が「{result}」として登録されました。"
                    f"記録ID：{record_id}。"
                    f"確認者：{checked_by or '未設定'}。"
                    f"次アクション：{next_action or '未設定'}。"
                ),
                target_user=item.get("owner", "") or checked_by or "製造責任者",
                category="製造管理",
                priority=priority,
                related_type="manufacturing",
                related_id=item_id,
            )

    return redirect("manufacturing:management_item_detail", item_id=item_id)

def incident_list(request):
    all_incidents = load_incidents()
    incidents = all_incidents

    keyword = request.GET.get("q", "").strip()
    area = request.GET.get("area", "").strip()
    status = request.GET.get("status", "").strip()
    severity = request.GET.get("severity", "").strip()

    if keyword:
        filtered_incidents = []

        for incident in incidents:
            target_text = " ".join([
                str(incident.get("id", "")),
                str(incident.get("area", "")),
                str(incident.get("incident_date", "")),
                str(incident.get("title", "")),
                str(incident.get("description", "")),
                str(incident.get("severity", "")),
                str(incident.get("detected_by", "")),
                str(incident.get("owner", "")),
                str(incident.get("status", "")),
                str(incident.get("corrective_action", "")),
                str(incident.get("preventive_action", "")),
                str(incident.get("related_management_item_id", "")),
                str(incident.get("related_task_id", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_incidents.append(incident)

        incidents = filtered_incidents

    if area:
        incidents = [
            incident for incident in incidents
            if str(incident.get("area", "")) == area
        ]

    if status:
        incidents = [
            incident for incident in incidents
            if str(incident.get("status", "")) == status
        ]

    if severity:
        incidents = [
            incident for incident in incidents
            if str(incident.get("severity", "")) == severity
        ]

    high_count = len([
        incident for incident in all_incidents
        if incident.get("severity") in ["重大", "高"]
    ])

    open_count = len([
        incident for incident in all_incidents
        if incident.get("status") in ["未対応", "対応中", "再発防止確認中"]
    ])

    context = {
        "incidents": incidents,
        "keyword": keyword,
        "area": area,
        "status": status,
        "severity": severity,
        "area_choices": INCIDENT_AREA_CHOICES,
        "status_choices": INCIDENT_STATUS_CHOICES,
        "severity_choices": INCIDENT_SEVERITY_CHOICES,
        "total_count": len(all_incidents),
        "display_count": len(incidents),
        "high_count": high_count,
        "open_count": open_count,
    }

    return render(request, "manufacturing/incident_list.html", context)


def incident_create(request):
    employees = load_employees()
    management_items = load_management_items()

    if request.method == "POST":
        area = request.POST.get("area", "").strip()
        incident_date = request.POST.get("incident_date", "").strip()
        title = request.POST.get("title", "").strip()
        description = request.POST.get("description", "").strip()
        severity = request.POST.get("severity", "中").strip()

        detected_by_employee_id = request.POST.get("detected_by_employee_id", "").strip()
        detected_by = resolve_employee_name(
            detected_by_employee_id,
            fallback=request.POST.get("detected_by", "").strip(),
        )

        owner_employee_id = request.POST.get("owner_employee_id", "").strip()
        owner = resolve_employee_name(
            owner_employee_id,
            fallback=request.POST.get("owner", "").strip(),
        )

        status = request.POST.get("status", "未対応").strip()
        corrective_action = request.POST.get("corrective_action", "").strip()
        preventive_action = request.POST.get("preventive_action", "").strip()
        related_management_item_id = request.POST.get("related_management_item_id", "").strip()
        create_task = request.POST.get("create_task") == "1"
        due_date = request.POST.get("due_date", "").strip()

        if title:
            incident_id, related_task_id = create_manufacturing_incident(
                area=area,
                incident_date=incident_date,
                title=title,
                description=description,
                severity=severity,
                detected_by=detected_by,
                owner=owner,
                status=status,
                corrective_action=corrective_action,
                preventive_action=preventive_action,
                related_management_item_id=related_management_item_id,
                create_task=create_task,
                due_date=due_date,
            )

            if should_create_incident_notification(severity):
                priority = get_incident_notification_priority(severity)

                create_notification(
                    title="製造インシデントが登録されました",
                    message=(
                        f"製造インシデント「{title}」が登録されました。"
                        f"インシデントID：{incident_id}。"
                        f"領域：{area or '未設定'}。"
                        f"重大度：{severity}。"
                        f"担当者：{owner or '未設定'}。"
                        f"関連タスク：{related_task_id or 'なし'}。"
                    ),
                    target_user=owner or "製造責任者",
                    category="製造インシデント",
                    priority=priority,
                    related_type="manufacturing_incident",
                    related_id=incident_id,
                )

            return redirect("manufacturing:incident_list")

    return render(request, "manufacturing/incident_form.html", {
        "employees": employees,
        "management_items": management_items,
        "area_choices": INCIDENT_AREA_CHOICES,
        "status_choices": INCIDENT_STATUS_CHOICES,
        "severity_choices": INCIDENT_SEVERITY_CHOICES,
    })


def incident_detail(request, incident_id):
    incident = find_incident_by_id(incident_id)

    return render(request, "manufacturing/incident_detail.html", {
        "incident": incident,
    })