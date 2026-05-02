from django.shortcuts import redirect, render

from core.formatters import normalize_amount

from .services import (
    create_request,
    find_attachments_by_request_id,
    find_histories_by_request_id,
    find_request_by_id,
    load_requests,
    save_request_attachment,
    update_request_status,
)


REQUEST_TYPE_CHOICES = [
    "稟議",
    "経費",
    "出張",
    "購買",
    "契約審査",
    "その他",
]

REQUEST_STATUS_CHOICES = [
    "承認待ち",
    "承認済",
    "差戻し",
    "却下",
]


def request_list(request):
    all_requests = load_requests()
    requests = all_requests

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered_requests = []

        for item in requests:
            target_text = " ".join([
                str(item.get("id", "")),
                str(item.get("request_type", "")),
                str(item.get("title", "")),
                str(item.get("applicant", "")),
                str(item.get("department", "")),
                str(item.get("amount", "")),
                str(item.get("status", "")),
                str(item.get("approver", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_requests.append(item)

        requests = filtered_requests

    return render(request, "workflows/request_list.html", {
        "requests": requests,
        "keyword": keyword,
        "total_count": len(all_requests),
        "display_count": len(requests),
    })


def request_create(request):
    if request.method == "POST":
        request_type = request.POST.get("request_type", "").strip()
        title = request.POST.get("title", "").strip()
        applicant = request.POST.get("applicant", "").strip()
        department = request.POST.get("department", "").strip()
        amount = normalize_amount(request.POST.get("amount", ""))
        approver = request.POST.get("approver", "").strip()
        description = request.POST.get("description", "").strip()

        if title:
            request_id = create_request(
                request_type=request_type,
                title=title,
                applicant=applicant,
                department=department,
                amount=amount,
                approver=approver,
                description=description,
            )
            return redirect("workflows:request_detail", request_id=request_id)

    return render(request, "workflows/request_form.html", {
        "request_type_choices": REQUEST_TYPE_CHOICES,
    })


def request_detail(request, request_id):
    workflow_request = find_request_by_id(request_id)
    histories = find_histories_by_request_id(request_id)
    attachments = find_attachments_by_request_id(request_id)

    return render(request, "workflows/request_detail.html", {
        "workflow_request": workflow_request,
        "histories": histories,
        "attachments": attachments,
        "status_choices": REQUEST_STATUS_CHOICES,
    })


def update_status(request, request_id):
    if request.method == "POST":
        status = request.POST.get("status", "").strip()
        actor = request.POST.get("actor", "").strip()
        comment = request.POST.get("comment", "").strip()

        if not actor:
            actor = "承認者"

        if status:
            update_request_status(
                request_id=request_id,
                new_status=status,
                actor=actor,
                comment=comment,
            )

    return redirect("workflows:request_detail", request_id=request_id)


def upload_attachment(request, request_id):
    if request.method == "POST":
        uploaded_file = request.FILES.get("attachment")
        uploaded_by = request.POST.get("uploaded_by", "").strip()

        if not uploaded_by:
            uploaded_by = "未設定"

        if uploaded_file:
            save_request_attachment(
                request_id=request_id,
                uploaded_file=uploaded_file,
                uploaded_by=uploaded_by,
            )

    return redirect("workflows:request_detail", request_id=request_id)