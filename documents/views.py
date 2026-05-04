from django.http import FileResponse, Http404
from django.shortcuts import redirect, render

from notifications.services import create_notification
from tasks.services import add_task

from .services import (
    find_document_by_id,
    load_documents,
    resolve_completed_file_path,
    resolve_template_file_path,
    save_completed_document_file,
    update_document_status,
)

DOCUMENT_STATUS_CHOICES = [
    "未整備",
    "作成中",
    "要確認",
    "整備済",
    "要改定",
    "不要",
    "廃止",
]


def should_create_status_notification(status):
    return status in [
        "未整備",
        "未確認",
        "要確認",
        "レビュー中",
        "要改定",
    ]


def get_notification_priority(document, status):
    required_level = document.get("required_level", "")

    if status in ["未整備", "要改定"]:
        return "高"

    if required_level in ["必須", "法定必須", "法定必須級", "重要"]:
        return "高"

    if status in ["未確認", "要確認", "レビュー中"]:
        return "中"

    return "低"


def document_list(request):
    all_documents = load_documents()
    documents = all_documents

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered_documents = []

        for doc in documents:
            target_text = " ".join([
                str(doc.get("id", "")),
                str(doc.get("category", "")),
                str(doc.get("subcategory", "")),
                str(doc.get("document_name", "")),
                str(doc.get("document_type", "")),
                str(doc.get("required_level", "")),
                str(doc.get("owner_dept", "")),
                str(doc.get("owner_department", "")),
                str(doc.get("owner", "")),
                str(doc.get("status", "")),
                str(doc.get("template_available", "")),
                str(doc.get("risk_level", "")),
                str(doc.get("related_question_ids", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_documents.append(doc)

        documents = filtered_documents

    grouped_documents = {}

    for doc in documents:
        category = doc.get("category", "") or "未分類"

        if category not in grouped_documents:
            grouped_documents[category] = []

        grouped_documents[category].append(doc)

    return render(request, "documents/document_list.html", {
        "documents": documents,
        "grouped_documents": grouped_documents,
        "keyword": keyword,
        "total_count": len(all_documents),
        "display_count": len(documents),
    })


def document_detail(request, document_id):
    document = find_document_by_id(document_id)

    if document is None:
        return render(request, "documents/document_detail.html", {
            "document": None,
            "status_choices": DOCUMENT_STATUS_CHOICES,
        })

    return render(request, "documents/document_detail.html", {
        "document": document,
        "status_choices": DOCUMENT_STATUS_CHOICES,
    })

def download_template(request, document_id):
    document = find_document_by_id(document_id)

    if document is None:
        raise Http404("文書が見つかりません。")

    template_path = resolve_template_file_path(document)

    if template_path is None:
        raise Http404("ひな形ファイルが見つかりません。")

    return FileResponse(
        open(template_path, "rb"),
        as_attachment=False,
        filename=template_path.name,
    )

def download_completed_document(request, document_id):
    document = find_document_by_id(document_id)

    if document is None:
        raise Http404("文書が見つかりません。")

    completed_path = resolve_completed_file_path(document)

    if completed_path is None:
        raise Http404("完成文書ファイルが見つかりません。")

    return FileResponse(
        open(completed_path, "rb"),
        as_attachment=False,
        filename=completed_path.name,
    )


def upload_completed_document(request, document_id):
    document = find_document_by_id(document_id)

    if document is None:
        raise Http404("文書が見つかりません。")

    if request.method == "POST":
        uploaded_file = request.FILES.get("completed_file")
        completed_by = request.POST.get("completed_by", "").strip()

        if uploaded_file:
            save_completed_document_file(
                document_id=document_id,
                uploaded_file=uploaded_file,
                completed_by=completed_by,
            )

    return redirect("documents:document_detail", document_id=document_id)

def update_status(request, document_id):
    if request.method == "POST":
        status = request.POST.get("status", "").strip()

        if status:
            update_result = update_document_status(document_id, status)

            if update_result:
                document = find_document_by_id(document_id)

                if document and should_create_status_notification(status):
                    priority = get_notification_priority(document, status)

                    create_notification(
                        title="文書ステータスが更新されました",
                        message=(
                            f"文書「{document.get('document_name', '')}」の状態が"
                            f"「{status}」に更新されました。"
                            f"カテゴリ：{document.get('category', '')}。"
                            f"主管部署：{document.get('owner_dept', '')}。"
                            f"重要度：{document.get('required_level', '')}。"
                        ),
                        target_user=document.get("owner_dept", "") or "文書管理者",
                        category="文書管理",
                        priority=priority,
                        related_type="documents",
                        related_id=document_id,
                    )

    return redirect("documents:document_detail", document_id=document_id)


def create_task_from_document(request, document_id):
    document = find_document_by_id(document_id)

    if request.method == "POST" and document:
        task_name = request.POST.get("task_name", "").strip()
        owner = request.POST.get("owner", "").strip()
        due_date = request.POST.get("due_date", "").strip()
        priority = request.POST.get("priority", "中").strip()

        if not task_name:
            task_name = f"{document.get('document_name', '')}を整備する"

        task_id = add_task(
            task_name=task_name,
            category=document.get("category", ""),
            owner=owner,
            due_date=due_date,
            status="未着手",
            priority=priority,
            related_document_id=document.get("id", ""),
        )

        if task_id:
            create_notification(
                title="文書整備タスクが作成されました",
                message=(
                    f"文書「{document.get('document_name', '')}」から"
                    f"タスクが作成されました。"
                    f"タスクID：{task_id}。"
                    f"担当者：{owner or '未設定'}。"
                    f"期限：{due_date or '未設定'}。"
                ),
                target_user=owner or document.get("owner_dept", "") or "文書管理者",
                category="文書管理",
                priority=priority or "中",
                related_type="tasks",
                related_id=task_id,
            )

    return redirect("documents:document_detail", document_id=document_id)