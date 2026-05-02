from django.shortcuts import render

# Create your views here.
from django.shortcuts import render

from .services import (
    find_employee_by_id,
    load_departments,
    load_employees,
    load_positions,
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


def employee_detail(request, employee_id):
    employee = find_employee_by_id(employee_id)

    return render(request, "organizations/employee_detail.html", {
        "employee": employee,
    })