from django.shortcuts import render

from documents.services import load_documents
from tasks.services import load_tasks
from questionnaires.services import load_diagnosis_summaries


def home(request):
    documents = load_documents()
    tasks = load_tasks()

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

    total_documents = len(documents)

    not_ready_documents = [
        doc for doc in documents
        if doc.get("status") in not_ready_statuses
    ]

    total_tasks = len(tasks)

    action_tasks = [
        task for task in tasks
        if task.get("status") in action_task_statuses
    ]

    high_priority_tasks = [
        task for task in action_tasks
        if task.get("priority") == "高"
    ]

    diagnosis_summaries = load_diagnosis_summaries()
    latest_diagnoses = sorted(
        diagnosis_summaries,
        key=lambda x: x.get("answered_at", ""),
        reverse=True,
    )[:5]

    context = {
        "total_documents": total_documents,
        "not_ready_document_count": len(not_ready_documents),
        "total_tasks": total_tasks,
        "action_task_count": len(action_tasks),
        "high_priority_task_count": len(high_priority_tasks),
        "not_ready_documents": not_ready_documents[:5],
        "action_tasks": action_tasks[:5],
        "latest_diagnoses": latest_diagnoses,
    }

    return render(request, "dashboard/home.html", context)