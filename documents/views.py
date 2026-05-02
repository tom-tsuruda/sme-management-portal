from django.shortcuts import redirect, render

from tasks.services import add_task

from .services import (
    find_document_by_id,
    load_documents,
    update_document_status,
)


DOCUMENT_STATUS_CHOICES = [
    "未整備",
    "未確認",
    "要確認",
    "作成中",
    "レビュー中",
    "整備済",
    "要改定",
    "廃止",
]


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
                str(doc.get("document_name", "")),
                str(doc.get("required_level", "")),
                str(doc.get("owner_dept", "")),
                str(doc.get("status", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_documents.append(doc)

        documents = filtered_documents

    return render(request, "documents/document_list.html", {
        "documents": documents,
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


def update_status(request, document_id):
    if request.method == "POST":
        status = request.POST.get("status", "").strip()

        if status:
            update_document_status(document_id, status)

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

        add_task(
            task_name=task_name,
            category=document.get("category", ""),
            owner=owner,
            due_date=due_date,
            status="未着手",
            priority=priority,
            related_document_id=document.get("id", ""),
        )

    return redirect("documents:document_detail", document_id=document_id)