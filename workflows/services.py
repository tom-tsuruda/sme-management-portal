from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings

from core.formatters import format_amount, format_japanese_datetime


REQUEST_COLUMNS = [
    "id",
    "request_type",
    "title",
    "applicant",
    "department",
    "amount",
    "status",
    "approver",
    "description",
    "created_at",
    "updated_at",
]

HISTORY_COLUMNS = [
    "id",
    "request_id",
    "action",
    "actor",
    "comment",
    "created_at",
]

ATTACHMENT_COLUMNS = [
    "id",
    "request_id",
    "original_filename",
    "stored_filename",
    "file_path",
    "file_url",
    "uploaded_by",
    "uploaded_at",
]


def get_workflow_excel_path():
    return Path(settings.BASE_DIR) / "data" / "workflow_data.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_workflow_excel():
    excel_path = get_workflow_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"workflow_data_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_workflow_excel():
    excel_path = get_workflow_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    requests_df = pd.DataFrame(columns=REQUEST_COLUMNS)
    histories_df = pd.DataFrame(columns=HISTORY_COLUMNS)
    attachments_df = pd.DataFrame(columns=ATTACHMENT_COLUMNS)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        requests_df.to_excel(writer, sheet_name="requests", index=False)
        histories_df.to_excel(writer, sheet_name="approval_histories", index=False)
        attachments_df.to_excel(writer, sheet_name="attachments", index=False)


def read_sheet(sheet_name, columns):
    ensure_workflow_excel()
    excel_path = get_workflow_excel_path()

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=columns)

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    return df[columns]


def write_workflow_data(requests_df, histories_df, attachments_df):
    excel_path = get_workflow_excel_path()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        requests_df.to_excel(writer, sheet_name="requests", index=False)
        histories_df.to_excel(writer, sheet_name="approval_histories", index=False)
        attachments_df.to_excel(writer, sheet_name="attachments", index=False)


def load_requests():
    df = read_sheet("requests", REQUEST_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        row["amount"] = format_amount(row.get("amount", ""))
        row["created_at"] = format_japanese_datetime(row.get("created_at", ""))
        row["updated_at"] = format_japanese_datetime(row.get("updated_at", ""))

    return records


def load_approval_histories():
    df = read_sheet("approval_histories", HISTORY_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        row["created_at"] = format_japanese_datetime(row.get("created_at", ""))

    return records


def load_attachments():
    df = read_sheet("attachments", ATTACHMENT_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        row["uploaded_at"] = format_japanese_datetime(row.get("uploaded_at", ""))

    return records


def find_request_by_id(request_id):
    requests = load_requests()

    for request in requests:
        if str(request.get("id")) == str(request_id):
            return request

    return None


def find_histories_by_request_id(request_id):
    histories = load_approval_histories()

    return [
        history for history in histories
        if str(history.get("request_id")) == str(request_id)
    ]


def find_attachments_by_request_id(request_id):
    attachments = load_attachments()

    return [
        attachment for attachment in attachments
        if str(attachment.get("request_id")) == str(request_id)
    ]


def generate_next_request_id(df):
    if "id" not in df.columns or df.empty:
        return "REQ-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("REQ-"):
            try:
                number = int(text.replace("REQ-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"REQ-{max_number + 1:03d}"


def generate_next_history_id(df):
    if "id" not in df.columns or df.empty:
        return "HIST-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("HIST-"):
            try:
                number = int(text.replace("HIST-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"HIST-{max_number + 1:03d}"


def generate_next_attachment_id(df):
    if "id" not in df.columns or df.empty:
        return "ATT-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("ATT-"):
            try:
                number = int(text.replace("ATT-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"ATT-{max_number + 1:03d}"


def add_history(histories_df, request_id, action, actor, comment):
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history_id = generate_next_history_id(histories_df)

    new_row = {
        "id": history_id,
        "request_id": request_id,
        "action": action,
        "actor": actor,
        "comment": comment,
        "created_at": now_text,
    }

    return pd.concat([histories_df, pd.DataFrame([new_row])], ignore_index=True)


def create_request(request_type, title, applicant, department, amount, approver, description):
    requests_df = read_sheet("requests", REQUEST_COLUMNS)
    histories_df = read_sheet("approval_histories", HISTORY_COLUMNS)
    attachments_df = read_sheet("attachments", ATTACHMENT_COLUMNS)

    backup_workflow_excel()

    request_id = generate_next_request_id(requests_df)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = {
        "id": request_id,
        "request_type": request_type,
        "title": title,
        "applicant": applicant,
        "department": department,
        "amount": amount,
        "status": "承認待ち",
        "approver": approver,
        "description": description,
        "created_at": now_text,
        "updated_at": now_text,
    }

    requests_df = pd.concat([requests_df, pd.DataFrame([new_row])], ignore_index=True)

    histories_df = add_history(
        histories_df=histories_df,
        request_id=request_id,
        action="申請",
        actor=applicant,
        comment="申請を作成しました。",
    )

    write_workflow_data(requests_df, histories_df, attachments_df)

    return request_id


def update_request_status(request_id, new_status, actor, comment):
    requests_df = read_sheet("requests", REQUEST_COLUMNS)
    histories_df = read_sheet("approval_histories", HISTORY_COLUMNS)
    attachments_df = read_sheet("attachments", ATTACHMENT_COLUMNS)

    if "id" not in requests_df.columns or "status" not in requests_df.columns:
        return False

    target_index = requests_df.index[requests_df["id"].astype(str) == str(request_id)].tolist()

    if not target_index:
        return False

    backup_workflow_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    requests_df.loc[target_index[0], "status"] = new_status
    requests_df.loc[target_index[0], "updated_at"] = now_text

    histories_df = add_history(
        histories_df=histories_df,
        request_id=request_id,
        action=new_status,
        actor=actor,
        comment=comment,
    )

    write_workflow_data(requests_df, histories_df, attachments_df)

    return True


def sanitize_filename(filename):
    """
    Windowsでも扱いやすいように、最低限ファイル名を安全化する。
    """
    unsafe_chars = ["\\", "/", ":", "*", "?", "\"", "<", ">", "|"]

    safe_name = str(filename)

    for char in unsafe_chars:
        safe_name = safe_name.replace(char, "_")

    return safe_name


def save_request_attachment(request_id, uploaded_file, uploaded_by):
    """
    申請に添付ファイルを保存する。
    ファイル本体は media/workflow_attachments/REQ-001/ に保存し、
    ExcelにはパスとURLを保存する。
    """
    if not uploaded_file:
        return None

    requests_df = read_sheet("requests", REQUEST_COLUMNS)
    histories_df = read_sheet("approval_histories", HISTORY_COLUMNS)
    attachments_df = read_sheet("attachments", ATTACHMENT_COLUMNS)

    backup_workflow_excel()

    attachment_id = generate_next_attachment_id(attachments_df)
    now = datetime.now()
    now_text = now.strftime("%Y-%m-%d %H:%M:%S")
    timestamp = now.strftime("%Y%m%d_%H%M%S")

    original_filename = sanitize_filename(uploaded_file.name)
    stored_filename = f"{timestamp}_{original_filename}"

    relative_dir = Path("workflow_attachments") / str(request_id)
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
        "request_id": request_id,
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
        request_id=request_id,
        action="添付追加",
        actor=uploaded_by,
        comment=f"{original_filename} を添付しました。",
    )

    write_workflow_data(requests_df, histories_df, attachments_df)

    return attachment_id