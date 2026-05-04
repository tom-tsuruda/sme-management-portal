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

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(columns=DEPARTMENT_COLUMNS).to_excel(
            writer, sheet_name="departments", index=False
        )
        pd.DataFrame(columns=POSITION_COLUMNS).to_excel(
            writer, sheet_name="positions", index=False
        )
        pd.DataFrame(columns=EMPLOYEE_COLUMNS).to_excel(
            writer, sheet_name="employees", index=False
        )


def read_sheet(sheet_name, columns):
    ensure_organization_excel()
    excel_path = get_organization_excel_path()

    try:
        df = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=columns)

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    df = df[columns].copy()

    for column in columns:
        df[column] = df[column].astype(str)

    return df


def write_organization_excel(departments_df, positions_df, employees_df):
    excel_path = get_organization_excel_path()

    departments_df = departments_df.copy()
    positions_df = positions_df.copy()
    employees_df = employees_df.copy()

    for column in DEPARTMENT_COLUMNS:
        if column not in departments_df.columns:
            departments_df[column] = ""

    for column in POSITION_COLUMNS:
        if column not in positions_df.columns:
            positions_df[column] = ""

    for column in EMPLOYEE_COLUMNS:
        if column not in employees_df.columns:
            employees_df[column] = ""

    departments_df = departments_df[DEPARTMENT_COLUMNS].copy()
    positions_df = positions_df[POSITION_COLUMNS].copy()
    employees_df = employees_df[EMPLOYEE_COLUMNS].copy()

    for column in DEPARTMENT_COLUMNS:
        departments_df[column] = departments_df[column].astype(str)

    for column in POSITION_COLUMNS:
        positions_df[column] = positions_df[column].astype(str)

    for column in EMPLOYEE_COLUMNS:
        employees_df[column] = employees_df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        departments_df.to_excel(writer, sheet_name="departments", index=False)
        positions_df.to_excel(writer, sheet_name="positions", index=False)
        employees_df.to_excel(writer, sheet_name="employees", index=False)


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


def generate_next_id(df, prefix, width=3):
    if "id" not in df.columns or df.empty:
        return f"{prefix}-{1:0{width}d}"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith(f"{prefix}-"):
            try:
                number = int(text.replace(f"{prefix}-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"{prefix}-{max_number + 1:0{width}d}"


def generate_next_code(df, column, prefix, width=3):
    if column not in df.columns or df.empty:
        return f"{prefix}{1:0{width}d}"

    max_number = 0

    for value in df[column].dropna():
        text = str(value).strip()

        if text.startswith(prefix):
            try:
                number = int(text.replace(prefix, ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"{prefix}{max_number + 1:0{width}d}"


def normalize_checkbox(value):
    return "1" if str(value) in ["1", "on", "true", "True", "TRUE", "承認者", "有効"] else "0"


def add_department(
    department_name,
    department_code="",
    parent_department_id="",
    manager_employee_id="",
    description="",
    is_active="1",
):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    department_id = generate_next_id(departments_df, "DEPT")

    if not department_code:
        department_code = generate_next_code(departments_df, "department_code", "D")

    new_row = {
        "id": department_id,
        "department_code": department_code,
        "department_name": department_name,
        "parent_department_id": parent_department_id,
        "manager_employee_id": manager_employee_id,
        "description": description,
        "is_active": normalize_checkbox(is_active),
        "created_at": now_text,
        "updated_at": now_text,
    }

    departments_df = pd.concat(
        [departments_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_organization_excel(departments_df, positions_df, employees_df)

    return department_id


def update_department(
    department_id,
    department_code,
    department_name,
    parent_department_id,
    manager_employee_id,
    description,
    is_active,
):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    target_index = departments_df.index[
        departments_df["id"].astype(str) == str(department_id)
    ].tolist()

    if not target_index:
        return False

    index = target_index[0]
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    departments_df.loc[index, "department_code"] = department_code
    departments_df.loc[index, "department_name"] = department_name
    departments_df.loc[index, "parent_department_id"] = parent_department_id
    departments_df.loc[index, "manager_employee_id"] = manager_employee_id
    departments_df.loc[index, "description"] = description
    departments_df.loc[index, "is_active"] = normalize_checkbox(is_active)
    departments_df.loc[index, "updated_at"] = now_text

    write_organization_excel(departments_df, positions_df, employees_df)
    return True


def delete_department(department_id):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    departments_df = departments_df[
        departments_df["id"].astype(str) != str(department_id)
    ].copy()

    write_organization_excel(departments_df, positions_df, employees_df)
    return True


def add_position(
    position_name,
    position_code="",
    rank="",
    approval_limit_amount="",
    description="",
    is_active="1",
):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    position_id = generate_next_id(positions_df, "POS")

    if not position_code:
        position_code = generate_next_code(positions_df, "position_code", "P")

    new_row = {
        "id": position_id,
        "position_code": position_code,
        "position_name": position_name,
        "rank": rank,
        "approval_limit_amount": approval_limit_amount,
        "description": description,
        "is_active": normalize_checkbox(is_active),
        "created_at": now_text,
        "updated_at": now_text,
    }

    positions_df = pd.concat(
        [positions_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_organization_excel(departments_df, positions_df, employees_df)

    return position_id


def update_position(
    position_id,
    position_code,
    position_name,
    rank,
    approval_limit_amount,
    description,
    is_active,
):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    target_index = positions_df.index[
        positions_df["id"].astype(str) == str(position_id)
    ].tolist()

    if not target_index:
        return False

    index = target_index[0]
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    positions_df.loc[index, "position_code"] = position_code
    positions_df.loc[index, "position_name"] = position_name
    positions_df.loc[index, "rank"] = rank
    positions_df.loc[index, "approval_limit_amount"] = approval_limit_amount
    positions_df.loc[index, "description"] = description
    positions_df.loc[index, "is_active"] = normalize_checkbox(is_active)
    positions_df.loc[index, "updated_at"] = now_text

    write_organization_excel(departments_df, positions_df, employees_df)
    return True


def delete_position(position_id):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    positions_df = positions_df[
        positions_df["id"].astype(str) != str(position_id)
    ].copy()

    write_organization_excel(departments_df, positions_df, employees_df)
    return True


def add_employee(
    employee_name,
    employee_code="",
    department_id="",
    position_id="",
    email="",
    role="",
    supervisor_employee_id="",
    is_approver="0",
    is_active="1",
):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    employee_id = generate_next_id(employees_df, "EMP")

    if not employee_code:
        employee_code = generate_next_code(employees_df, "employee_code", "E")

    new_row = {
        "id": employee_id,
        "employee_code": employee_code,
        "employee_name": employee_name,
        "department_id": department_id,
        "position_id": position_id,
        "email": email,
        "role": role,
        "supervisor_employee_id": supervisor_employee_id,
        "is_approver": normalize_checkbox(is_approver),
        "is_active": normalize_checkbox(is_active),
        "created_at": now_text,
        "updated_at": now_text,
    }

    employees_df = pd.concat(
        [employees_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_organization_excel(departments_df, positions_df, employees_df)

    return employee_id


def update_employee(
    employee_id,
    employee_code,
    employee_name,
    department_id,
    position_id,
    email,
    role,
    supervisor_employee_id,
    is_approver,
    is_active,
):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    target_index = employees_df.index[
        employees_df["id"].astype(str) == str(employee_id)
    ].tolist()

    if not target_index:
        return False

    index = target_index[0]
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    employees_df.loc[index, "employee_code"] = employee_code
    employees_df.loc[index, "employee_name"] = employee_name
    employees_df.loc[index, "department_id"] = department_id
    employees_df.loc[index, "position_id"] = position_id
    employees_df.loc[index, "email"] = email
    employees_df.loc[index, "role"] = role
    employees_df.loc[index, "supervisor_employee_id"] = supervisor_employee_id
    employees_df.loc[index, "is_approver"] = normalize_checkbox(is_approver)
    employees_df.loc[index, "is_active"] = normalize_checkbox(is_active)
    employees_df.loc[index, "updated_at"] = now_text

    write_organization_excel(departments_df, positions_df, employees_df)
    return True


def delete_employee(employee_id):
    departments_df = read_sheet("departments", DEPARTMENT_COLUMNS)
    positions_df = read_sheet("positions", POSITION_COLUMNS)
    employees_df = read_sheet("employees", EMPLOYEE_COLUMNS)

    employees_df = employees_df[
        employees_df["id"].astype(str) != str(employee_id)
    ].copy()

    write_organization_excel(departments_df, positions_df, employees_df)
    return True