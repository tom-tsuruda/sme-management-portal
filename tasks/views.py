from urllib.parse import urlencode

from django.shortcuts import redirect, render

from documents.services import find_document_by_id
from notifications.services import create_notification
from organizations.services import find_employee_by_id, load_employees

from .services import add_task, find_task_by_id, load_tasks, update_task_status, update_task


TASK_STATUS_CHOICES = [
    "未着手",
    "進行中",
    "完了",
    "要対応",
]


def resolve_employee_name(employee_id, fallback=""):
    employee = find_employee_by_id(employee_id) if employee_id else None

    if employee:
        return employee.get("employee_name", "")

    return fallback


def get_task_notification_priority(task, status):
    task_priority = task.get("priority", "")

    if status == "要対応":
        return "高"

    if task_priority == "高":
        return "高"

    if status == "進行中":
        return "中"

    if status == "完了":
        return "低"

    return "中"


def create_task(request):
    """
    画面から新しいタスクを追加する。
    """
    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        category = request.POST.get("category", "").strip()

        owner_employee_id = request.POST.get("owner_employee_id", "").strip()
        owner = resolve_employee_name(
            owner_employee_id,
            fallback=request.POST.get("owner", "").strip(),
        )

        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "未着手").strip()
        priority = request.POST.get("priority", "中").strip()
        related_document_id = request.POST.get("related_document_id", "").strip()

        if task_name:
            task_id = add_task(
                task_name=task_name,
                category=category,
                owner=owner,
                due_date=due_date,
                status=status,
                priority=priority,
                related_document_id=related_document_id,
            )

            if task_id:
                create_notification(
                    title="タスクが作成されました",
                    message=(
                        f"タスク「{task_name}」が作成されました。"
                        f"タスクID：{task_id}。"
                        f"カテゴリ：{category or '未設定'}。"
                        f"担当者：{owner or '未設定'}。"
                        f"期限：{due_date or '未設定'}。"
                    ),
                    target_user=owner or "タスク担当者",
                    category="タスク管理",
                    priority=priority or "中",
                    related_type="tasks",
                    related_id=task_id,
                )

    return redirect("/tasks/")


def task_list(request):
    all_tasks = load_tasks()
    tasks = all_tasks
    employees = load_employees()

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
        "employees": employees,
        "keyword": keyword,
        "total_count": len(all_tasks),
        "display_count": len(tasks),
    })


def task_detail(request, task_id):
    """
    タスク詳細画面を表示する。
    """
    task = find_task_by_id(task_id)
    related_document = None

    if task and task.get("related_document_id"):
        related_document = find_document_by_id(task.get("related_document_id"))

    return render(request, "tasks/task_detail.html", {
        "task": task,
        "related_document": related_document,
        "status_choices": TASK_STATUS_CHOICES,
    })

def edit_task(request, task_id):
    """
    タスクの内容を編集する。
    """
    task = find_task_by_id(task_id)
    employees = load_employees()

    if not task:
        return redirect("tasks:task_list")

    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        category = request.POST.get("category", "").strip()

        owner_employee_id = request.POST.get("owner_employee_id", "").strip()
        owner = resolve_employee_name(
            owner_employee_id,
            fallback=request.POST.get("owner", "").strip(),
        )

        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "").strip()
        priority = request.POST.get("priority", "").strip()
        related_document_id = request.POST.get("related_document_id", "").strip()

        if not status:
            status = task.get("status", "未着手")

        if not priority:
            priority = task.get("priority", "中")

        update_result = update_task(
            task_id=task_id,
            task_name=task_name,
            category=category,
            owner=owner,
            due_date=due_date,
            status=status,
            priority=priority,
            related_document_id=related_document_id,
        )

        if update_result:
            updated_task = find_task_by_id(task_id)

            create_notification(
                title="タスクが編集されました",
                message=(
                    f"タスク「{task_name}」が編集されました。"
                    f"タスクID：{task_id}。"
                    f"担当者：{owner or '未設定'}。"
                    f"期限：{due_date or '未設定'}。"
                    f"状態：{status}。"
                ),
                target_user=owner or "タスク担当者",
                category="タスク管理",
                priority=get_task_notification_priority(updated_task or task, status),
                related_type="tasks",
                related_id=task_id,
            )

        return redirect("tasks:task_detail", task_id=task_id)

    return render(request, "tasks/task_edit.html", {
        "task": task,
        "employees": employees,
        "status_choices": TASK_STATUS_CHOICES,
        "priority_choices": ["高", "中", "低"],
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
            update_result = update_task_status(task_id, status)

            if update_result:
                task = find_task_by_id(task_id)

                if task:
                    priority = get_task_notification_priority(task, status)

                    create_notification(
                        title="タスク状態が更新されました",
                        message=(
                            f"タスク「{task.get('task_name', '')}」の状態が"
                            f"「{status}」に更新されました。"
                            f"タスクID：{task_id}。"
                            f"担当者：{task.get('owner', '') or '未設定'}。"
                            f"期限：{task.get('due_date', '') or '未設定'}。"
                        ),
                        target_user=task.get("owner", "") or "タスク担当者",
                        category="タスク管理",
                        priority=priority,
                        related_type="tasks",
                        related_id=task_id,
                    )

        if keyword:
            query = urlencode({"q": keyword})
            return redirect(f"/tasks/?{query}")

    return redirect("/tasks/")


def update_status_from_detail(request, task_id):
    """
    タスク詳細画面から状態を更新する。
    """
    if request.method == "POST":
        status = request.POST.get("status", "")

        if task_id and status:
            update_result = update_task_status(task_id, status)

            if update_result:
                task = find_task_by_id(task_id)

                if task:
                    priority = get_task_notification_priority(task, status)

                    create_notification(
                        title="タスク状態が更新されました",
                        message=(
                            f"タスク「{task.get('task_name', '')}」の状態が"
                            f"「{status}」に更新されました。"
                            f"タスクID：{task_id}。"
                            f"担当者：{task.get('owner', '') or '未設定'}。"
                            f"期限：{task.get('due_date', '') or '未設定'}。"
                        ),
                        target_user=task.get("owner", "") or "タスク担当者",
                        category="タスク管理",
                        priority=priority,
                        related_type="tasks",
                        related_id=task_id,
                    )

    return redirect("tasks:task_detail", task_id=task_id)

def edit_task(request, task_id):
    """
    タスクの内容を編集する。
    """
    task = find_task_by_id(task_id)
    employees = load_employees()

    if not task:
        return redirect("tasks:task_list")

    if request.method == "POST":
        task_name = request.POST.get("task_name", "").strip()
        category = request.POST.get("category", "").strip()

        owner_employee_id = request.POST.get("owner_employee_id", "").strip()
        owner = resolve_employee_name(
            owner_employee_id,
            fallback=request.POST.get("owner", "").strip(),
        )

        due_date = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "").strip()
        priority = request.POST.get("priority", "").strip()
        related_document_id = request.POST.get("related_document_id", "").strip()

        if not task_name:
            task_name = task.get("task_name", "")

        if not category:
            category = task.get("category", "")

        if not owner:
            owner = task.get("owner", "")

        if not status:
            status = task.get("status", "未着手")

        if not priority:
            priority = task.get("priority", "中")

        update_result = update_task(
            task_id=task_id,
            task_name=task_name,
            category=category,
            owner=owner,
            due_date=due_date,
            status=status,
            priority=priority,
            related_document_id=related_document_id,
        )

        if update_result:
            updated_task = find_task_by_id(task_id)

            create_notification(
                title="タスクが編集されました",
                message=(
                    f"タスク「{task_name}」が編集されました。"
                    f"タスクID：{task_id}。"
                    f"担当者：{owner or '未設定'}。"
                    f"期限：{due_date or '未設定'}。"
                    f"状態：{status}。"
                ),
                target_user=owner or "タスク担当者",
                category="タスク管理",
                priority=get_task_notification_priority(updated_task or task, status),
                related_type="tasks",
                related_id=task_id,
            )

        return redirect("tasks:task_detail", task_id=task_id)

    return render(request, "tasks/task_edit.html", {
        "task": task,
        "employees": employees,
        "status_choices": TASK_STATUS_CHOICES,
        "priority_choices": ["高", "中", "低"],
    })