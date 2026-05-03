from django.shortcuts import render

from documents.services import load_documents
from tasks.services import load_tasks
from questionnaires.services import load_diagnosis_summaries
from workflows.services import load_requests
from expenses.services import load_expenses
from manufacturing.services import load_incidents, load_management_items
from kpi.services import (
    get_latest_manufacturing_kpi_with_comparison,
    get_latest_monthly_kpi_with_comparison,
)
from governance.services import load_governance_items
from notifications.services import load_notifications, load_unread_notifications


def home(request):
    documents = load_documents()
    tasks = load_tasks()
    diagnosis_summaries = load_diagnosis_summaries()
    workflow_requests = load_requests()
    expenses = load_expenses()
    manufacturing_items = load_management_items()
    manufacturing_incidents = load_incidents()
    governance_items = load_governance_items()
    notifications = load_notifications()
    unread_notifications = load_unread_notifications()

    latest_monthly_kpi = get_latest_monthly_kpi_with_comparison()
    latest_manufacturing_kpi = get_latest_manufacturing_kpi_with_comparison()

    not_ready_statuses = [
        "未整備",
        "未確認",
        "要確認",
        "作成中",
        "レビュー中",
        "要改定",
    ]

    action_task_statuses = [
        "未着手",
        "進行中",
        "要対応",
    ]

    not_ready_documents = [
        doc for doc in documents
        if doc.get("status") in not_ready_statuses
    ]

    action_tasks = [
        task for task in tasks
        if task.get("status") in action_task_statuses
    ]

    high_priority_tasks = [
        task for task in action_tasks
        if task.get("priority") == "高"
    ]

    latest_diagnoses = sorted(
        diagnosis_summaries,
        key=lambda x: x.get("answered_at", ""),
        reverse=True,
    )[:20]

    approval_waiting_requests = [
        item for item in workflow_requests
        if item.get("status") == "承認待ち"
    ]

    returned_requests = [
        item for item in workflow_requests
        if item.get("status") == "差戻し"
    ]

    approved_requests = [
        item for item in workflow_requests
        if item.get("status") == "承認済"
    ]

    expense_applying = [
        item for item in expenses
        if item.get("status") == "申請中"
    ]

    expense_returned = [
        item for item in expenses
        if item.get("status") == "差戻し"
    ]

    expense_approved = [
        item for item in expenses
        if item.get("status") == "承認済"
    ]

    expense_settled = [
        item for item in expenses
        if item.get("status") == "精算済"
    ]

    manufacturing_not_ready = [
        item for item in manufacturing_items
        if item.get("status") == "未整備"
    ]

    manufacturing_need_check = [
        item for item in manufacturing_items
        if item.get("status") == "要確認"
    ]

    manufacturing_need_improvement = [
        item for item in manufacturing_items
        if item.get("status") == "要改善"
    ]

    manufacturing_overdue = [
        item for item in manufacturing_items
        if item.get("status") == "期限超過"
    ]

    manufacturing_high_risk = [
        item for item in manufacturing_items
        if item.get("risk_level") == "高"
    ]

    manufacturing_action_items = [
        item for item in manufacturing_items
        if item.get("status") in ["未整備", "要確認", "要改善", "期限超過"]
    ]

    manufacturing_incident_high = [
        incident for incident in manufacturing_incidents
        if incident.get("severity") in ["重大", "高"]
    ]

    manufacturing_incident_open = [
        incident for incident in manufacturing_incidents
        if incident.get("status") in ["未対応", "対応中", "再発防止確認中"]
    ]

    manufacturing_incident_critical_open = [
        incident for incident in manufacturing_incidents
        if incident.get("severity") in ["重大", "高"]
        and incident.get("status") in ["未対応", "対応中", "再発防止確認中"]
    ]

    latest_manufacturing_incidents = manufacturing_incidents[:20]

    governance_not_ready = [
        item for item in governance_items
        if item.get("status") == "未整備"
    ]

    governance_need_check = [
        item for item in governance_items
        if item.get("status") == "要確認"
    ]

    governance_need_revision = [
        item for item in governance_items
        if item.get("status") == "要改定"
    ]

    governance_overdue = [
        item for item in governance_items
        if item.get("status") == "期限超過"
    ]

    governance_high_risk = [
        item for item in governance_items
        if item.get("risk_level") == "高"
    ]

    governance_required = [
        item for item in governance_items
        if item.get("required_level") in ["法定必須", "法定必須級"]
    ]

    governance_action_items = [
        item for item in governance_items
        if item.get("status") in ["未整備", "要確認", "要改定", "期限超過"]
    ]

    high_priority_unread_notifications = [
        notification for notification in unread_notifications
        if notification.get("priority") == "高"
    ]

    context = {
        "total_documents": len(documents),
        "not_ready_document_count": len(not_ready_documents),
        "not_ready_documents": not_ready_documents[:20],

        "total_tasks": len(tasks),
        "action_task_count": len(action_tasks),
        "high_priority_task_count": len(high_priority_tasks),
        "action_tasks": action_tasks[:20],

        "latest_diagnoses": latest_diagnoses,

        "total_request_count": len(workflow_requests),
        "approval_waiting_request_count": len(approval_waiting_requests),
        "returned_request_count": len(returned_requests),
        "approved_request_count": len(approved_requests),

        "total_expense_count": len(expenses),
        "expense_applying_count": len(expense_applying),
        "expense_returned_count": len(expense_returned),
        "expense_approved_count": len(expense_approved),
        "expense_settled_count": len(expense_settled),

        "manufacturing_total_count": len(manufacturing_items),
        "manufacturing_not_ready_count": len(manufacturing_not_ready),
        "manufacturing_need_check_count": len(manufacturing_need_check),
        "manufacturing_need_improvement_count": len(manufacturing_need_improvement),
        "manufacturing_overdue_count": len(manufacturing_overdue),
        "manufacturing_high_risk_count": len(manufacturing_high_risk),
        "latest_manufacturing_items": manufacturing_action_items[:20],

        "manufacturing_incident_total_count": len(manufacturing_incidents),
        "manufacturing_incident_high_count": len(manufacturing_incident_high),
        "manufacturing_incident_open_count": len(manufacturing_incident_open),
        "manufacturing_incident_critical_open_count": len(manufacturing_incident_critical_open),
        "latest_manufacturing_incidents": latest_manufacturing_incidents,

        "governance_total_count": len(governance_items),
        "governance_not_ready_count": len(governance_not_ready),
        "governance_need_check_count": len(governance_need_check),
        "governance_need_revision_count": len(governance_need_revision),
        "governance_overdue_count": len(governance_overdue),
        "governance_high_risk_count": len(governance_high_risk),
        "governance_required_count": len(governance_required),
        "latest_governance_items": governance_action_items[:20],

        "notification_total_count": len(notifications),
        "unread_notification_count": len(unread_notifications),
        "high_priority_unread_notification_count": len(high_priority_unread_notifications),
        "latest_notifications": notifications[:20],

        "latest_monthly_kpi": latest_monthly_kpi,
        "latest_manufacturing_kpi": latest_manufacturing_kpi,
    }

    return render(request, "dashboard/home.html", context)