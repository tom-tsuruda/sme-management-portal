from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings
from core.backup import backup_excel_file

ALLOWED_COMPLETED_FILE_EXTENSIONS = {
    ".doc",
    ".docx",
    ".pdf",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
}

def resolve_template_file_path(document):
    """
    document_master.xlsx の template_file_path を、
    プロジェクト配下の安全な実ファイルパスに変換する。
    存在しない場合やプロジェクト外を指す場合は None を返す。
    """
    if not document:
        return None

    template_file_path = str(document.get("template_file_path", "") or "").strip()

    if not template_file_path:
        return None

    normalized_path = template_file_path.replace("\\", "/").lstrip("/")
    base_dir = Path(settings.BASE_DIR).resolve()
    candidate_path = (base_dir / normalized_path).resolve()

    try:
        candidate_path.relative_to(base_dir)
    except ValueError:
        return None

    if not candidate_path.is_file():
        return None

    return candidate_path

def get_completed_document_dir():
    completed_dir = Path(settings.MEDIA_ROOT) / "documents" / "completed"
    completed_dir.mkdir(parents=True, exist_ok=True)
    return completed_dir


def resolve_completed_file_path(document):
    """
    document_master.xlsx の completed_file_path を、
    プロジェクト配下の安全な実ファイルパスに変換する。
    """
    if not document:
        return None

    completed_file_path = str(document.get("completed_file_path", "") or "").strip()

    if not completed_file_path:
        return None

    normalized_path = completed_file_path.replace("\\", "/").lstrip("/")
    base_dir = Path(settings.BASE_DIR).resolve()
    candidate_path = (base_dir / normalized_path).resolve()

    try:
        candidate_path.relative_to(base_dir)
    except ValueError:
        return None

    if not candidate_path.is_file():
        return None

    return candidate_path


def document_has_completed_file(document):
    return resolve_completed_file_path(document) is not None


def build_completed_file_storage_path(document_id, original_filename):
    original_path = Path(original_filename)
    suffix = original_path.suffix.lower()

    if suffix not in ALLOWED_COMPLETED_FILE_EXTENSIONS:
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_document_id = str(document_id).replace("/", "_").replace("\\", "_")
    filename = f"{safe_document_id}_{timestamp}{suffix}"

    completed_dir = get_completed_document_dir()
    return completed_dir / filename


def save_completed_document_file(document_id, uploaded_file, completed_by=""):
    """
    完成文書ファイルを media/documents/completed/ に保存し、
    document_master.xlsx の completed_file_path などを更新する。
    保存できたら True、失敗したら False を返す。
    """
    excel_path = get_document_excel_path()

    if not excel_path.exists():
        return False

    save_path = build_completed_file_storage_path(
        document_id=document_id,
        original_filename=uploaded_file.name,
    )

    if save_path is None:
        return False

    df = pd.read_excel(
        excel_path,
        sheet_name="documents",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    if "id" not in df.columns:
        return False

    target_index = df.index[df["id"].astype(str) == str(document_id)].tolist()

    if not target_index:
        return False

    backup_document_excel()

    with open(save_path, "wb+") as destination:
        for chunk in uploaded_file.chunks():
            destination.write(chunk)

    relative_path = save_path.resolve().relative_to(Path(settings.BASE_DIR).resolve())

    required_columns = [
        "completed_file_path",
        "completed_file_name",
        "completed_at",
        "completed_by",
        "status",
        "updated_at",
    ]

    for column in required_columns:
        if column not in df.columns:
            df[column] = ""

    index = target_index[0]
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df.loc[index, "completed_file_path"] = str(relative_path).replace("\\", "/")
    df.loc[index, "completed_file_name"] = uploaded_file.name
    df.loc[index, "completed_at"] = now_text
    df.loc[index, "completed_by"] = completed_by
    df.loc[index, "status"] = "整備済"
    df.loc[index, "updated_at"] = now_text

    for column in df.columns:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="documents", index=False)

    return True

def document_has_template_file(document):
    return resolve_template_file_path(document) is not None

def get_document_excel_path():
    return Path(settings.BASE_DIR) / "data" / "document_master.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_document_excel():
    return backup_excel_file(
        excel_path=get_document_excel_path(),
        base_dir=settings.BASE_DIR,
        file_prefix="document_master",
        keep_count=3,
    )


def format_japanese_date(value):
    if value == "" or value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

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
        "completed_at",
    ]:
        if key in doc:
            doc[key] = format_japanese_date(doc.get(key, ""))

    doc["template_exists"] = document_has_template_file(doc)
    doc["completed_file_exists"] = document_has_completed_file(doc)

    return doc

def load_documents():
    excel_path = get_document_excel_path()

    if not excel_path.exists():
        return []

    df = pd.read_excel(
        excel_path,
        sheet_name="documents",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

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

    df = pd.read_excel(
        excel_path,
        sheet_name="documents",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    if "id" not in df.columns or "status" not in df.columns:
        return False

    target_index = df.index[df["id"].astype(str) == str(document_id)].tolist()

    if not target_index:
        return False

    backup_document_excel()

    df.loc[target_index[0], "status"] = str(new_status)

    if "updated_at" in df.columns:
        df.loc[target_index[0], "updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for column in df.columns:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="documents", index=False)

    return True