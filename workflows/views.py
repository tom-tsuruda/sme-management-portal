from django.shortcuts import redirect, render

from core.formatters import normalize_amount

from notifications.services import create_notification

from .services import (
    create_request,
    find_attachments_by_request_id,
    find_histories_by_request_id,
    find_request_by_id,
    load_requests,
    save_request_attachment,
    update_request_status,
)
from organizations.services import (
    find_department_by_id,
    find_employee_by_id,
    load_departments,
    load_employees,
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
def is_active_approver(employee):
    is_active = str(employee.get("is_active", "")).strip()
    is_approver = str(employee.get("is_approver", "")).strip()

    return is_active in ["1", "true", "True", "TRUE", "有効"] and is_approver in [
        "1",
        "true",
        "True",
        "TRUE",
        "はい",
        "○",
    ]


def get_workflow_master_context():
    employees = load_employees()
    departments = load_departments()
    approvers = [employee for employee in employees if is_active_approver(employee)]

    return {
        "employees": employees,
        "departments": departments,
        "approvers": approvers,
    }


def resolve_employee_name(employee_id, fallback=""):
    employee = find_employee_by_id(employee_id) if employee_id else None

    if employee:
        return employee.get("employee_name", "")

    return fallback


def resolve_department_name(department_id, applicant_employee_id="", fallback=""):
    department = find_department_by_id(department_id) if department_id else None

    if department:
        return department.get("department_name", "")

    applicant_employee = find_employee_by_id(applicant_employee_id) if applicant_employee_id else None

    if applicant_employee and applicant_employee.get("department_name"):
        return applicant_employee.get("department_name", "")

    return fallback

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
    master_context = get_workflow_master_context()

    if request.method == "POST":
        request_type = request.POST.get("request_type", "").strip()
        title = request.POST.get("title", "").strip()

        applicant_employee_id = request.POST.get("applicant_employee_id", "").strip()
        department_id = request.POST.get("department_id", "").strip()
        approver_employee_id = request.POST.get("approver_employee_id", "").strip()

        applicant = resolve_employee_name(
            applicant_employee_id,
            fallback=request.POST.get("applicant", "").strip(),
        )

        department = resolve_department_name(
            department_id,
            applicant_employee_id=applicant_employee_id,
            fallback=request.POST.get("department", "").strip(),
        )

        amount = normalize_amount(request.POST.get("amount", ""))

        approver = resolve_employee_name(
            approver_employee_id,
            fallback=request.POST.get("approver", "").strip(),
        )

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

            create_notification(
                title="承認待ち申請があります",
                message=(
                    f"{applicant}さんから「{title}」の申請が作成されました。"
                    f"申請ID：{request_id}。申請・承認画面で内容を確認してください。"
                ),
                target_user=approver or "承認者",
                category="申請承認",
                priority="高",
                related_type="workflows",
                related_id=request_id,
            )

            return redirect("workflows:request_detail", request_id=request_id)

    context = {
        "request_type_choices": REQUEST_TYPE_CHOICES,
    }
    context.update(master_context)

    return render(request, "workflows/request_form.html", context)


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
            update_result = update_request_status(
                request_id=request_id,
                new_status=status,
                actor=actor,
                comment=comment,
            )

            if update_result:
                workflow_request = find_request_by_id(request_id)

                if workflow_request:
                    applicant = workflow_request.get("applicant", "")
                    approver = workflow_request.get("approver", "")
                    title = workflow_request.get("title", "")
                    request_type = workflow_request.get("request_type", "")

                    if status == "承認済":
                        create_notification(
                            title="申請が承認されました",
                            message=(
                                f"申請「{title}」が{actor}さんにより承認されました。"
                                f"申請ID：{request_id}。"
                            ),
                            target_user=applicant or "申請者",
                            category="申請承認",
                            priority="中",
                            related_type="workflows",
                            related_id=request_id,
                        )

                    elif status == "差戻し":
                        create_notification(
                            title="申請が差戻しされました",
                            message=(
                                f"申請「{title}」が{actor}さんにより差戻しされました。"
                                f"申請ID：{request_id}。"
                                f"コメント：{comment or 'コメントなし'}"
                            ),
                            target_user=applicant or "申請者",
                            category="申請承認",
                            priority="高",
                            related_type="workflows",
                            related_id=request_id,
                        )

                    elif status == "却下":
                        create_notification(
                            title="申請が却下されました",
                            message=(
                                f"申請「{title}」が{actor}さんにより却下されました。"
                                f"申請ID：{request_id}。"
                                f"コメント：{comment or 'コメントなし'}"
                            ),
                            target_user=applicant or "申請者",
                            category="申請承認",
                            priority="高",
                            related_type="workflows",
                            related_id=request_id,
                        )

                    elif status == "承認待ち":
                        create_notification(
                            title="申請が承認待ちになりました",
                            message=(
                                f"{request_type}申請「{title}」が承認待ちになりました。"
                                f"申請ID：{request_id}。"
                            ),
                            target_user=approver or "承認者",
                            category="申請承認",
                            priority="高",
                            related_type="workflows",
                            related_id=request_id,
                        )

    return redirect("workflows:request_detail", request_id=request_id)


def upload_attachment(request, request_id):
    if request.method == "POST":
        uploaded_file = request.FILES.get("attachment")
        uploaded_by = request.POST.get("uploaded_by", "").strip()

        if not uploaded_by:
            uploaded_by = "未設定"

        if uploaded_file:
            attachment_id = save_request_attachment(
                request_id=request_id,
                uploaded_file=uploaded_file,
                uploaded_by=uploaded_by,
            )

            workflow_request = find_request_by_id(request_id)

            if attachment_id and workflow_request:
                title = workflow_request.get("title", "")
                approver = workflow_request.get("approver", "")

                create_notification(
                    title="申請に添付ファイルが追加されました",
                    message=(
                        f"申請「{title}」に添付ファイルが追加されました。"
                        f"申請ID：{request_id}。"
                    ),
                    target_user=approver or "承認者",
                    category="申請承認",
                    priority="低",
                    related_type="workflows",
                    related_id=request_id,
                )

    return redirect("workflows:request_detail", request_id=request_id)