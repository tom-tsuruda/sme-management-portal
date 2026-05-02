from django.shortcuts import render

from documents.services import load_documents
from tasks.services import load_tasks


def home(request):
    documents = load_documents()
    tasks = load_tasks()

    total_documents = len(documents)
    not_ready_documents = len([
        doc for doc in documents
        if doc.get("status") in ["未整備", "未確認", "要確認"]
    ])

    total_tasks = len(tasks)
    action_tasks = len([
        task for task in tasks
        if task.get("status") in ["未着手", "進行中", "要対応"]
    ])

    context = {
        "total_documents": total_documents,
        "not_ready_documents": not_ready_documents,
        "total_tasks": total_tasks,
        "action_tasks": action_tasks,
    }

    return render(request, "dashboard/home.html", context)