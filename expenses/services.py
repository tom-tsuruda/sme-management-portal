from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings

from core.formatters import format_amount, format_japanese_date, format_japanese_datetime


EXPENSE_COLUMNS = [
    "id",
    "expense_type",
    "title",
    "applicant",
    "department",
    "expense_date",
    "category",
    "amount",
    "payment_method",
    "status",
    "approver",
    "description",
    "created_at",
    "updated_at",
]

ATTACHMENT_COLUMNS = [
    "id",
    "expense_id",
    "original_filename",
    "stored_filename",
    "file_path",
    "file_url",
    "uploaded_by",
    "uploaded_at",
]

HISTORY_COLUMNS = [
    "id",
    "expense_id",
    "action",
    "actor",
    "comment",
    "created_at",
]


def get_expense_excel_path():
    return Path(settings.BASE_DIR) / "data" / "expense_data.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_expense_excel():
    excel_path = get_expense_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"expense_data_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_expense_excel():
    excel_path = get_expense_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    expenses_df = pd.DataFrame(columns=EXPENSE_COLUMNS)
    attachments_df = pd.DataFrame(columns=ATTACHMENT_COLUMNS)
    histories_df = pd.DataFrame(columns=HISTORY_COLUMNS)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        expenses_df.to_excel(writer, sheet_name="expenses", index=False)
        attachments_df.to_excel(writer, sheet_name="attachments", index=False)
        histories_df.to_excel(writer, sheet_name="histories", index=False)


def read_sheet(sheet_name, columns):
    ensure_expense_excel()
    excel_path = get_expense_excel_path()

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


def write_expense_data(expenses_df, attachments_df, histories_df):
    excel_path = get_expense_excel_path()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        expenses_df.to_excel(writer, sheet_name="expenses", index=False)
        attachments_df.to_excel(writer, sheet_name="attachments", index=False)
        histories_df.to_excel(writer, sheet_name="histories", index=False)


def load_expenses():
    df = read_sheet("expenses", EXPENSE_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        row["amount"] = format_amount(row.get("amount", ""))
        row["expense_date"] = format_japanese_date(row.get("expense_date", ""))
        row["created_at"] = format_japanese_datetime(row.get("created_at", ""))
        row["updated_at"] = format_japanese_datetime(row.get("updated_at", ""))

    return records


def load_attachments():
    df = read_sheet("attachments", ATTACHMENT_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        row["uploaded_at"] = format_japanese_datetime(row.get("uploaded_at", ""))

    return records


def load_histories():
    df = read_sheet("histories", HISTORY_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        row["created_at"] = format_japanese_datetime(row.get("created_at", ""))

    return records


def find_expense_by_id(expense_id):
    expenses = load_expenses()

    for expense in expenses:
        if str(expense.get("id")) == str(expense_id):
            return expense

    return None


def find_attachments_by_expense_id(expense_id):
    attachments = load_attachments()

    return [
        attachment for attachment in attachments
        if str(attachment.get("expense_id")) == str(expense_id)
    ]


def find_histories_by_expense_id(expense_id):
    histories = load_histories()

    return [
        history for history in histories
        if str(history.get("expense_id")) == str(expense_id)
    ]


def generate_next_expense_id(df):
    if "id" not in df.columns or df.empty:
        return "EXP-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("EXP-"):
            try:
                number = int(text.replace("EXP-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"EXP-{max_number + 1:03d}"


def generate_next_attachment_id(df):
    if "id" not in df.columns or df.empty:
        return "EATT-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("EATT-"):
            try:
                number = int(text.replace("EATT-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"EATT-{max_number + 1:03d}"


def generate_next_history_id(df):
    if "id" not in df.columns or df.empty:
        return "EHIST-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("EHIST-"):
            try:
                number = int(text.replace("EHIST-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"EHIST-{max_number + 1:03d}"


def add_history(histories_df, expense_id, action, actor, comment):
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history_id = generate_next_history_id(histories_df)

    new_row = {
        "id": history_id,
        "expense_id": expense_id,
        "action": action,
        "actor": actor,
        "comment": comment,
        "created_at": now_text,
    }

    return pd.concat([histories_df, pd.DataFrame([new_row])], ignore_index=True)


def create_expense(
    expense_type,
    title,
    applicant,
    department,
    expense_date,
    category,
    amount,
    payment_method,
    approver,
    description,
):
    expenses_df = read_sheet("expenses", EXPENSE_COLUMNS)
    attachments_df = read_sheet("attachments", ATTACHMENT_COLUMNS)
    histories_df = read_sheet("histories", HISTORY_COLUMNS)

    backup_expense_excel()

    expense_id = generate_next_expense_id(expenses_df)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = {
        "id": expense_id,
        "expense_type": expense_type,
        "title": title,
        "applicant": applicant,
        "department": department,
        "expense_date": expense_date,
        "category": category,
        "amount": amount,
        "payment_method": payment_method,
        "status": "申請中",
        "approver": approver,
        "description": description,
        "created_at": now_text,
        "updated_at": now_text,
    }

    expenses_df = pd.concat([expenses_df, pd.DataFrame([new_row])], ignore_index=True)

    histories_df = add_history(
        histories_df=histories_df,
        expense_id=expense_id,
        action="申請",
        actor=applicant,
        comment="経費申請を作成しました。",
    )

    write_expense_data(expenses_df, attachments_df, histories_df)

    return expense_id


def update_expense_status(expense_id, new_status, actor, comment):
    expenses_df = read_sheet("expenses", EXPENSE_COLUMNS)
    attachments_df = read_sheet("attachments", ATTACHMENT_COLUMNS)
    histories_df = read_sheet("histories", HISTORY_COLUMNS)

    target_index = expenses_df.index[expenses_df["id"].astype(str) == str(expense_id)].tolist()

    if not target_index:
        return False

    backup_expense_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    expenses_df.loc[target_index[0], "status"] = new_status
    expenses_df.loc[target_index[0], "updated_at"] = now_text

    histories_df = add_history(
        histories_df=histories_df,
        expense_id=expense_id,
        action=new_status,
        actor=actor,
        comment=comment,
    )

    write_expense_data(expenses_df, attachments_df, histories_df)

    return True


def sanitize_filename(filename):
    unsafe_chars = ["\\", "/", ":", "*", "?", "\"", "<", ">", "|"]
    safe_name = str(filename)

    for char in unsafe_chars:
        safe_name = safe_name.replace(char, "_")

    return safe_name


def save_expense_attachment(expense_id, uploaded_file, uploaded_by):
    if not uploaded_file:
        return None

    expenses_df = read_sheet("expenses", EXPENSE_COLUMNS)
    attachments_df = read_sheet("attachments", ATTACHMENT_COLUMNS)
    histories_df = read_sheet("histories", HISTORY_COLUMNS)

    backup_expense_excel()

    attachment_id = generate_next_attachment_id(attachments_df)
    now = datetime.now()
    now_text = now.strftime("%Y-%m-%d %H:%M:%S")
    timestamp = now.strftime("%Y%m%d_%H%M%S")

    original_filename = sanitize_filename(uploaded_file.name)
    stored_filename = f"{timestamp}_{original_filename}"

    relative_dir = Path("expense_attachments") / str(expense_id)
    save_dir = Path(settings.MEDIA_ROOT) / relative_dir
    save_dir.mkdir(parents=True, exist_ok=True)

    save_path = save_dir / stored_filename

    with open(save_path, "wb+") as destination:
        for chunk in uploaded_file.chunks():
            destination.write(chunk)

    relative_file_path = relative_dir / stored_filename
    file_url = f"{settings.MEDIA_URL}{relative_file_path.as_posix()}"

    new_row = {
        "id": attachment_id,
        "expense_id": expense_id,
        "original_filename": original_filename,
        "stored_filename": stored_filename,
        "file_path": str(save_path),
        "file_url": file_url,
        "uploaded_by": uploaded_by,
        "uploaded_at": now_text,
    }

    attachments_df = pd.concat([attachments_df, pd.DataFrame([new_row])], ignore_index=True)

    histories_df = add_history(
        histories_df=histories_df,
        expense_id=expense_id,
        action="添付追加",
        actor=uploaded_by,
        comment=f"{original_filename} を添付しました。",
    )

    write_expense_data(expenses_df, attachments_df, histories_df)

    return attachment_id