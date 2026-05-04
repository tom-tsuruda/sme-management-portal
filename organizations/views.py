from django.shortcuts import redirect, render

from .services import (
    add_department,
    add_employee,
    add_position,
    delete_department,
    delete_employee,
    delete_position,
    find_department_by_id,
    find_employee_by_id,
    find_position_by_id,
    load_departments,
    load_employees,
    load_positions,
    update_department,
    update_employee,
    update_position,
)


def organization_home(request):
    departments = load_departments()
    positions = load_positions()
    employees = load_employees()

    approvers = [
        emp for emp in employees
        if str(emp.get("is_approver")) in ["1", "True", "true", "承認者"]
    ]

    context = {
        "department_count": len(departments),
        "position_count": len(positions),
        "employee_count": len(employees),
        "approver_count": len(approvers),
    }

    return render(request, "organizations/organization_home.html", context)


def department_list(request):
    departments = load_departments()

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []
        for dept in departments:
            target_text = " ".join([
                str(dept.get("id", "")),
                str(dept.get("department_code", "")),
                str(dept.get("department_name", "")),
                str(dept.get("description", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(dept)

        departments = filtered

    return render(request, "organizations/department_list.html", {
        "departments": departments,
        "keyword": keyword,
    })


def department_create(request):
    departments = load_departments()
    employees = load_employees()

    if request.method == "POST":
        department_name = request.POST.get("department_name", "").strip()

        if department_name:
            add_department(
                department_name=department_name,
                department_code=request.POST.get("department_code", "").strip(),
                parent_department_id=request.POST.get("parent_department_id", "").strip(),
                manager_employee_id=request.POST.get("manager_employee_id", "").strip(),
                description=request.POST.get("description", "").strip(),
                is_active=request.POST.get("is_active", "0"),
            )

        return redirect("organizations:department_list")

    return render(request, "organizations/department_form.html", {
        "department": None,
        "departments": departments,
        "employees": employees,
    })


def department_edit(request, department_id):
    department = find_department_by_id(department_id)

    if not department:
        return redirect("organizations:department_list")

    departments = load_departments()
    employees = load_employees()

    if request.method == "POST":
        department_name = request.POST.get("department_name", "").strip()

        if department_name:
            update_department(
                department_id=department_id,
                department_code=request.POST.get("department_code", "").strip(),
                department_name=department_name,
                parent_department_id=request.POST.get("parent_department_id", "").strip(),
                manager_employee_id=request.POST.get("manager_employee_id", "").strip(),
                description=request.POST.get("description", "").strip(),
                is_active=request.POST.get("is_active", "0"),
            )

        return redirect("organizations:department_list")

    return render(request, "organizations/department_form.html", {
        "department": department,
        "departments": departments,
        "employees": employees,
    })


def department_delete(request, department_id):
    if request.method == "POST":
        delete_department(department_id)

    return redirect("organizations:department_list")


def position_list(request):
    positions = load_positions()

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []
        for pos in positions:
            target_text = " ".join([
                str(pos.get("id", "")),
                str(pos.get("position_code", "")),
                str(pos.get("position_name", "")),
                str(pos.get("description", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(pos)

        positions = filtered

    return render(request, "organizations/position_list.html", {
        "positions": positions,
        "keyword": keyword,
    })


def position_create(request):
    if request.method == "POST":
        position_name = request.POST.get("position_name", "").strip()

        if position_name:
            add_position(
                position_name=position_name,
                position_code=request.POST.get("position_code", "").strip(),
                rank=request.POST.get("rank", "").strip(),
                approval_limit_amount=request.POST.get("approval_limit_amount", "").strip(),
                description=request.POST.get("description", "").strip(),
                is_active=request.POST.get("is_active", "0"),
            )

        return redirect("organizations:position_list")

    return render(request, "organizations/position_form.html", {
        "position": None,
    })


def position_edit(request, position_id):
    position = find_position_by_id(position_id)

    if not position:
        return redirect("organizations:position_list")

    if request.method == "POST":
        position_name = request.POST.get("position_name", "").strip()

        if position_name:
            update_position(
                position_id=position_id,
                position_code=request.POST.get("position_code", "").strip(),
                position_name=position_name,
                rank=request.POST.get("rank", "").strip(),
                approval_limit_amount=request.POST.get("approval_limit_amount", "").strip(),
                description=request.POST.get("description", "").strip(),
                is_active=request.POST.get("is_active", "0"),
            )

        return redirect("organizations:position_list")

    return render(request, "organizations/position_form.html", {
        "position": position,
    })


def position_delete(request, position_id):
    if request.method == "POST":
        delete_position(position_id)

    return redirect("organizations:position_list")


def employee_list(request):
    employees = load_employees()

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered = []
        for emp in employees:
            target_text = " ".join([
                str(emp.get("id", "")),
                str(emp.get("employee_code", "")),
                str(emp.get("employee_name", "")),
                str(emp.get("department_name", "")),
                str(emp.get("position_name", "")),
                str(emp.get("email", "")),
                str(emp.get("role", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(emp)

        employees = filtered

    return render(request, "organizations/employee_list.html", {
        "employees": employees,
        "keyword": keyword,
    })


def employee_create(request):
    departments = load_departments()
    positions = load_positions()
    employees = load_employees()

    if request.method == "POST":
        employee_name = request.POST.get("employee_name", "").strip()

        if employee_name:
            add_employee(
                employee_name=employee_name,
                employee_code=request.POST.get("employee_code", "").strip(),
                department_id=request.POST.get("department_id", "").strip(),
                position_id=request.POST.get("position_id", "").strip(),
                email=request.POST.get("email", "").strip(),
                role=request.POST.get("role", "").strip(),
                supervisor_employee_id=request.POST.get("supervisor_employee_id", "").strip(),
                is_approver=request.POST.get("is_approver", "0"),
                is_active=request.POST.get("is_active", "0"),
            )

        return redirect("organizations:employee_list")

    return render(request, "organizations/employee_form.html", {
        "employee": None,
        "departments": departments,
        "positions": positions,
        "employees": employees,
    })


def employee_edit(request, employee_id):
    employee = find_employee_by_id(employee_id)

    if not employee:
        return redirect("organizations:employee_list")

    departments = load_departments()
    positions = load_positions()
    employees = load_employees()

    if request.method == "POST":
        employee_name = request.POST.get("employee_name", "").strip()

        if employee_name:
            update_employee(
                employee_id=employee_id,
                employee_code=request.POST.get("employee_code", "").strip(),
                employee_name=employee_name,
                department_id=request.POST.get("department_id", "").strip(),
                position_id=request.POST.get("position_id", "").strip(),
                email=request.POST.get("email", "").strip(),
                role=request.POST.get("role", "").strip(),
                supervisor_employee_id=request.POST.get("supervisor_employee_id", "").strip(),
                is_approver=request.POST.get("is_approver", "0"),
                is_active=request.POST.get("is_active", "0"),
            )

        return redirect("organizations:employee_list")

    return render(request, "organizations/employee_form.html", {
        "employee": employee,
        "departments": departments,
        "positions": positions,
        "employees": employees,
    })


def employee_delete(request, employee_id):
    if request.method == "POST":
        delete_employee(employee_id)

    return redirect("organizations:employee_list")


def employee_detail(request, employee_id):
    employee = find_employee_by_id(employee_id)

    return render(request, "organizations/employee_detail.html", {
        "employee": employee,
    })