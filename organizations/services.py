from pathlib import Path
from datetime import datetime

import pandas as pd
from django.conf import settings


DEPARTMENT_COLUMNS = [
    "id",
    "department_code",
    "department_name",
    "parent_department_id",
    "manager_employee_id",
    "description",
    "is_active",
    "created_at",
    "updated_at",
]

POSITION_COLUMNS = [
    "id",
    "position_code",
    "position_name",
    "rank",
    "approval_limit_amount",
    "description",
    "is_active",
    "created_at",
    "updated_at",
]

EMPLOYEE_COLUMNS = [
    "id",
    "employee_code",
    "employee_name",
    "department_id",
    "position_id",
    "email",
    "role",
    "supervisor_employee_id",
    "is_approver",
    "is_active",
    "created_at",
    "updated_at",
]


def get_organization_excel_path():
    return Path(settings.BASE_DIR) / "data" / "organization_master.xlsx"


def ensure_organization_excel():
    excel_path = get_organization_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    departments = [
        {
            "id": "DEPT-001",
            "department_code": "D001",
            "department_name": "経営企画部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-001",
            "description": "経営方針、KPI、全社企画を管理する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "DEPT-002",
            "department_code": "D002",
            "department_name": "総務部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-002",
            "description": "社内規程、庶務、ガバナンス文書を管理する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "DEPT-003",
            "department_code": "D003",
            "department_name": "製造部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-003",
            "description": "製造工程、生産計画、設備管理を担当する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "DEPT-004",
            "department_code": "D004",
            "department_name": "品質管理部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-004",
            "description": "検査、不適合、品質記録を管理する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "DEPT-005",
            "department_code": "D005",
            "department_name": "研究開発部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-005",
            "description": "研究テーマ、試作、技術開発を担当する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "DEPT-006",
            "department_code": "D006",
            "department_name": "安全環境部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-006",
            "description": "安全衛生、環境、化学物質管理を担当する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "DEPT-007",
            "department_code": "D007",
            "department_name": "経理部",
            "parent_department_id": "",
            "manager_employee_id": "EMP-007",
            "description": "経費、支払、会計連携データを管理する部門",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    positions = [
        {
            "id": "POS-001",
            "position_code": "P001",
            "position_name": "代表取締役",
            "rank": "1",
            "approval_limit_amount": "999999999",
            "description": "最終承認者",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "POS-002",
            "position_code": "P002",
            "position_name": "部長",
            "rank": "2",
            "approval_limit_amount": "1000000",
            "description": "部門責任者",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "POS-003",
            "position_code": "P003",
            "position_name": "課長",
            "rank": "3",
            "approval_limit_amount": "300000",
            "description": "課単位の管理者",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "POS-004",
            "position_code": "P004",
            "position_name": "担当者",
            "rank": "4",
            "approval_limit_amount": "0",
            "description": "一般担当者",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    employees = [
        {
            "id": "EMP-001",
            "employee_code": "E001",
            "employee_name": "経営 太郎",
            "department_id": "DEPT-001",
            "position_id": "POS-001",
            "email": "president@example.com",
            "role": "経営者",
            "supervisor_employee_id": "",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "EMP-002",
            "employee_code": "E002",
            "employee_name": "総務 花子",
            "department_id": "DEPT-002",
            "position_id": "POS-002",
            "email": "soumu@example.com",
            "role": "管理者",
            "supervisor_employee_id": "EMP-001",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "EMP-003",
            "employee_code": "E003",
            "employee_name": "製造 一郎",
            "department_id": "DEPT-003",
            "position_id": "POS-002",
            "email": "manufacturing@example.com",
            "role": "製造責任者",
            "supervisor_employee_id": "EMP-001",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "EMP-004",
            "employee_code": "E004",
            "employee_name": "品質 次郎",
            "department_id": "DEPT-004",
            "position_id": "POS-002",
            "email": "quality@example.com",
            "role": "品質責任者",
            "supervisor_employee_id": "EMP-001",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "EMP-005",
            "employee_code": "E005",
            "employee_name": "研究 三郎",
            "department_id": "DEPT-005",
            "position_id": "POS-002",
            "email": "rd@example.com",
            "role": "研究開発責任者",
            "supervisor_employee_id": "EMP-001",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "EMP-006",
            "employee_code": "E006",
            "employee_name": "安全 四郎",
            "department_id": "DEPT-006",
            "position_id": "POS-002",
            "email": "safety@example.com",
            "role": "安全環境責任者",
            "supervisor_employee_id": "EMP-001",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "EMP-007",
            "employee_code": "E007",
            "employee_name": "経理 五郎",
            "department_id": "DEPT-007",
            "position_id": "POS-002",
            "email": "accounting@example.com",
            "role": "経理責任者",
            "supervisor_employee_id": "EMP-001",
            "is_approver": "1",
            "is_active": "1",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(departments, columns=DEPARTMENT_COLUMNS).to_excel(
            writer, sheet_name="departments", index=False
        )
        pd.DataFrame(positions, columns=POSITION_COLUMNS).to_excel(
            writer, sheet_name="positions", index=False
        )
        pd.DataFrame(employees, columns=EMPLOYEE_COLUMNS).to_excel(
            writer, sheet_name="employees", index=False
        )


def read_sheet(sheet_name, columns):
    ensure_organization_excel()
    excel_path = get_organization_excel_path()

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=columns)

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    return df[columns]


def load_departments():
    df = read_sheet("departments", DEPARTMENT_COLUMNS)
    return df.to_dict(orient="records")


def load_positions():
    df = read_sheet("positions", POSITION_COLUMNS)
    return df.to_dict(orient="records")


def load_employees():
    departments = load_departments()
    positions = load_positions()

    department_map = {
        str(dept.get("id")): dept.get("department_name", "")
        for dept in departments
    }

    position_map = {
        str(pos.get("id")): pos.get("position_name", "")
        for pos in positions
    }

    df = read_sheet("employees", EMPLOYEE_COLUMNS)
    records = df.to_dict(orient="records")

    employee_map = {
        str(emp.get("id")): emp.get("employee_name", "")
        for emp in records
    }

    for emp in records:
        emp["department_name"] = department_map.get(str(emp.get("department_id")), "")
        emp["position_name"] = position_map.get(str(emp.get("position_id")), "")
        emp["supervisor_name"] = employee_map.get(str(emp.get("supervisor_employee_id")), "")

    return records


def find_department_by_id(department_id):
    for dept in load_departments():
        if str(dept.get("id")) == str(department_id):
            return dept
    return None


def find_position_by_id(position_id):
    for pos in load_positions():
        if str(pos.get("id")) == str(position_id):
            return pos
    return None


def find_employee_by_id(employee_id):
    for emp in load_employees():
        if str(emp.get("id")) == str(employee_id):
            return emp
    return None