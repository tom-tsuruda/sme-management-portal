from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings


def get_document_excel_path():
    return Path(settings.BASE_DIR) / "data" / "document_master.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_document_excel():
    excel_path = get_document_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"document_master_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def format_japanese_date(value):
    if value == "" or pd.isna(value):
        return ""

    try:
        date_value = pd.to_datetime(value)
        return f"{date_value.year}年{date_value.month}月{date_value.day}日"
    except Exception:
        return value


def normalize_document_record(doc):
    for key in [
        "created_at",
        "updated_at",
        "established_date",
        "revised_date",
        "next_review_date",
    ]:
        if key in doc:
            doc[key] = format_japanese_date(doc.get(key, ""))

    return doc


def load_documents():
    excel_path = get_document_excel_path()

    if not excel_path.exists():
        return []

    df = pd.read_excel(excel_path, sheet_name="documents")
    df = df.fillna("")

    records = df.to_dict(orient="records")

    for doc in records:
        normalize_document_record(doc)

    return records


def find_document_by_id(document_id):
    documents = load_documents()

    for doc in documents:
        if str(doc.get("id")) == str(document_id):
            return doc

    return None


def update_document_status(document_id, new_status):
    excel_path = get_document_excel_path()

    if not excel_path.exists():
        return False

    df = pd.read_excel(excel_path, sheet_name="documents")
    df = df.fillna("")

    if "id" not in df.columns or "status" not in df.columns:
        return False

    target_index = df.index[df["id"].astype(str) == str(document_id)].tolist()

    if not target_index:
        return False

    backup_document_excel()

    df.loc[target_index[0], "status"] = new_status

    if "updated_at" in df.columns:
        df.loc[target_index[0], "updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="documents", index=False)

    return True