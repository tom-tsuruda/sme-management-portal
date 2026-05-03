import csv

from django.http import HttpResponse
from django.shortcuts import redirect, render

from core.formatters import normalize_amount

from notifications.services import create_notification
from organizations.services import (
    find_department_by_id,
    find_employee_by_id,
    load_departments,
    load_employees,
)

from .services import (
    create_expense,
    find_attachments_by_expense_id,
    find_expense_by_id,
    find_histories_by_expense_id,
    load_expenses,
    save_expense_attachment,
    update_expense_status,
)


EXPENSE_TYPE_CHOICES = [
    "経費精算",
    "旅費精算",
    "交通費",
    "購買立替",
    "その他",
]

EXPENSE_CATEGORY_CHOICES = [
    "交通費",
    "宿泊費",
    "接待交際費",
    "会議費",
    "消耗品費",
    "通信費",
    "研修費",
    "その他",
]

PAYMENT_METHOD_CHOICES = [
    "現金",
    "クレジットカード",
    "電子マネー",
    "銀行振込",
    "立替",
    "その他",
]

EXPENSE_STATUS_CHOICES = [
    "申請中",
    "承認済",
    "差戻し",
    "却下",
    "精算済",
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


def get_expense_master_context():
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


def expense_list(request):
    all_expenses = load_expenses()
    expenses = all_expenses

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered_expenses = []

        for expense in expenses:
            target_text = " ".join([
                str(expense.get("id", "")),
                str(expense.get("expense_type", "")),
                str(expense.get("title", "")),
                str(expense.get("applicant", "")),
                str(expense.get("department", "")),
                str(expense.get("expense_date", "")),
                str(expense.get("category", "")),
                str(expense.get("amount", "")),
                str(expense.get("payment_method", "")),
                str(expense.get("status", "")),
                str(expense.get("approver", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_expenses.append(expense)

        expenses = filtered_expenses

    return render(request, "expenses/expense_list.html", {
        "expenses": expenses,
        "keyword": keyword,
        "total_count": len(all_expenses),
        "display_count": len(expenses),
    })


def expense_create(request):
    master_context = get_expense_master_context()

    if request.method == "POST":
        expense_type = request.POST.get("expense_type", "").strip()
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

        expense_date = request.POST.get("expense_date", "").strip()
        category = request.POST.get("category", "").strip()
        amount = normalize_amount(request.POST.get("amount", ""))
        payment_method = request.POST.get("payment_method", "").strip()

        approver = resolve_employee_name(
            approver_employee_id,
            fallback=request.POST.get("approver", "").strip(),
        )

        description = request.POST.get("description", "").strip()

        if title:
            expense_id = create_expense(
                expense_type=expense_type,
                title=title,
                applicant=applicant,
                department=department,
                expense_date=expense_date,
                category=category,
                amount=amount,
                payment_method=payment_method,
                approver=approver,
                description=description,
            )

            create_notification(
                title="経費申請が作成されました",
                message=(
                    f"{applicant}さんから「{title}」の経費申請が作成されました。"
                    f"経費ID：{expense_id}。金額：{amount}。"
                    f"経費・旅費精算画面で内容を確認してください。"
                ),
                target_user=approver or "経理責任者",
                category="経費精算",
                priority="高",
                related_type="expenses",
                related_id=expense_id,
            )

            uploaded_file = request.FILES.get("attachment")
            uploaded_by = applicant or "未設定"

            if uploaded_file:
                attachment_id = save_expense_attachment(
                    expense_id=expense_id,
                    uploaded_file=uploaded_file,
                    uploaded_by=uploaded_by,
                )

                if attachment_id:
                    create_notification(
                        title="経費申請に証憑が添付されました",
                        message=(
                            f"経費申請「{title}」に証憑ファイルが添付されました。"
                            f"経費ID：{expense_id}。"
                        ),
                        target_user=approver or "経理責任者",
                        category="経費精算",
                        priority="中",
                        related_type="expenses",
                        related_id=expense_id,
                    )

            return redirect("expenses:expense_detail", expense_id=expense_id)

    context = {
        "expense_type_choices": EXPENSE_TYPE_CHOICES,
        "expense_category_choices": EXPENSE_CATEGORY_CHOICES,
        "payment_method_choices": PAYMENT_METHOD_CHOICES,
    }
    context.update(master_context)

    return render(request, "expenses/expense_form.html", context)


def expense_detail(request, expense_id):
    expense = find_expense_by_id(expense_id)
    attachments = find_attachments_by_expense_id(expense_id)
    histories = find_histories_by_expense_id(expense_id)

    return render(request, "expenses/expense_detail.html", {
        "expense": expense,
        "attachments": attachments,
        "histories": histories,
        "status_choices": EXPENSE_STATUS_CHOICES,
    })


def update_status(request, expense_id):
    if request.method == "POST":
        status = request.POST.get("status", "").strip()
        actor = request.POST.get("actor", "").strip()
        comment = request.POST.get("comment", "").strip()

        if not actor:
            actor = "承認者"

        if status:
            update_result = update_expense_status(
                expense_id=expense_id,
                new_status=status,
                actor=actor,
                comment=comment,
            )

            if update_result:
                expense = find_expense_by_id(expense_id)

                if expense:
                    applicant = expense.get("applicant", "")
                    approver = expense.get("approver", "")
                    title = expense.get("title", "")
                    amount = expense.get("amount", "")

                    if status == "承認済":
                        create_notification(
                            title="経費申請が承認されました",
                            message=(
                                f"経費申請「{title}」が{actor}さんにより承認されました。"
                                f"経費ID：{expense_id}。金額：{amount}。"
                            ),
                            target_user=applicant or "申請者",
                            category="経費精算",
                            priority="中",
                            related_type="expenses",
                            related_id=expense_id,
                        )

                    elif status == "精算済":
                        create_notification(
                            title="経費申請が精算済になりました",
                            message=(
                                f"経費申請「{title}」が精算済になりました。"
                                f"経費ID：{expense_id}。金額：{amount}。"
                            ),
                            target_user=applicant or "申請者",
                            category="経費精算",
                            priority="中",
                            related_type="expenses",
                            related_id=expense_id,
                        )

                    elif status == "差戻し":
                        create_notification(
                            title="経費申請が差戻しされました",
                            message=(
                                f"経費申請「{title}」が{actor}さんにより差戻しされました。"
                                f"経費ID：{expense_id}。"
                                f"コメント：{comment or 'コメントなし'}"
                            ),
                            target_user=applicant or "申請者",
                            category="経費精算",
                            priority="高",
                            related_type="expenses",
                            related_id=expense_id,
                        )

                    elif status == "却下":
                        create_notification(
                            title="経費申請が却下されました",
                            message=(
                                f"経費申請「{title}」が{actor}さんにより却下されました。"
                                f"経費ID：{expense_id}。"
                                f"コメント：{comment or 'コメントなし'}"
                            ),
                            target_user=applicant or "申請者",
                            category="経費精算",
                            priority="高",
                            related_type="expenses",
                            related_id=expense_id,
                        )

                    elif status == "申請中":
                        create_notification(
                            title="経費申請が申請中になりました",
                            message=(
                                f"経費申請「{title}」が申請中になりました。"
                                f"経費ID：{expense_id}。"
                            ),
                            target_user=approver or "経理責任者",
                            category="経費精算",
                            priority="高",
                            related_type="expenses",
                            related_id=expense_id,
                        )

    return redirect("expenses:expense_detail", expense_id=expense_id)


def upload_attachment(request, expense_id):
    if request.method == "POST":
        uploaded_file = request.FILES.get("attachment")
        uploaded_by = request.POST.get("uploaded_by", "").strip()

        if not uploaded_by:
            uploaded_by = "未設定"

        if uploaded_file:
            attachment_id = save_expense_attachment(
                expense_id=expense_id,
                uploaded_file=uploaded_file,
                uploaded_by=uploaded_by,
            )

            expense = find_expense_by_id(expense_id)

            if attachment_id and expense:
                title = expense.get("title", "")
                approver = expense.get("approver", "")

                create_notification(
                    title="経費申請に証憑が追加されました",
                    message=(
                        f"経費申請「{title}」に証憑ファイルが追加されました。"
                        f"経費ID：{expense_id}。"
                    ),
                    target_user=approver or "経理責任者",
                    category="経費精算",
                    priority="中",
                    related_type="expenses",
                    related_id=expense_id,
                )

    return redirect("expenses:expense_detail", expense_id=expense_id)


def export_csv(request):
    expenses = load_expenses()

    export_statuses = ["承認済", "精算済"]

    target_expenses = [
        expense for expense in expenses
        if expense.get("status") in export_statuses
    ]

    response = HttpResponse(content_type="text/csv; charset=utf-8-sig")
    response["Content-Disposition"] = 'attachment; filename="expense_export.csv"'

    writer = csv.writer(response)

    writer.writerow([
        "申請ID",
        "経費区分",
        "件名",
        "申請者",
        "部署",
        "日付",
        "カテゴリ",
        "金額",
        "支払方法",
        "ステータス",
        "承認者",
        "摘要",
    ])

    for expense in target_expenses:
        writer.writerow([
            expense.get("id", ""),
            expense.get("expense_type", ""),
            expense.get("title", ""),
            expense.get("applicant", ""),
            expense.get("department", ""),
            expense.get("expense_date", ""),
            expense.get("category", ""),
            expense.get("amount", ""),
            expense.get("payment_method", ""),
            expense.get("status", ""),
            expense.get("approver", ""),
            expense.get("description", ""),
        ])

    return response