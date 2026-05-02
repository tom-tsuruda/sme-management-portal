from urllib.parse import urlencode

from django.shortcuts import redirect, render

from .services import load_tasks, update_task_status
from .services import add_task, load_tasks, update_task_status

def create_task(request):
    """
    画面から新しいタスクを追加する。
    """
    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        category = request.POST.get("category", "").strip()
        owner = request.POST.get("owner", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "未着手").strip()
        priority = request.POST.get("priority", "中").strip()
        related_document_id = request.POST.get("related_document_id", "").strip()

        if task_name:
            add_task(
                task_name=task_name,
                category=category,
                owner=owner,
                due_date=due_date,
                status=status,
                priority=priority,
                related_document_id=related_document_id,
            )

    return redirect("/tasks/")
def task_list(request):
    all_tasks = load_tasks()
    tasks = all_tasks

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered_tasks = []

        for task in tasks:
            target_text = " ".join([
                str(task.get("id", "")),
                str(task.get("task_name", "")),
                str(task.get("category", "")),
                str(task.get("owner", "")),
                str(task.get("due_date", "")),
                str(task.get("status", "")),
                str(task.get("priority", "")),
                str(task.get("related_document_id", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_tasks.append(task)

        tasks = filtered_tasks

    return render(request, "tasks/task_list.html", {
        "tasks": tasks,
        "keyword": keyword,
        "total_count": len(all_tasks),
        "display_count": len(tasks),
    })


def update_status(request):
    """
    タスク一覧画面から状態を更新する。
    """
    if request.method == "POST":
        task_id = request.POST.get("task_id", "")
        status = request.POST.get("status", "")
        keyword = request.POST.get("keyword", "")

        if task_id and status:
            update_task_status(task_id, status)

        if keyword:
            query = urlencode({"q": keyword})
            return redirect(f"/tasks/?{query}")

    return redirect("/tasks/")