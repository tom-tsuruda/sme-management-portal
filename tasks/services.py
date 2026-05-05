from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings
from core.backup import backup_excel_file

def get_task_excel_path():
    """
    task_data.xlsx のパスを返す。
    """
    return Path(settings.BASE_DIR) / "data" / "task_data.xlsx"


def get_backup_dir():
    """
    バックアップフォルダのパスを返す。
    なければ作成する。
    """
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_task_excel():
    return backup_excel_file(
        excel_path=get_task_excel_path(),
        base_dir=settings.BASE_DIR,
        file_prefix="task_data",
        keep_count=3,
    )


def format_japanese_date(value):
    """
    Excelの日付を 2026年5月31日 の形式に変換する。
    空欄や変換できない値はそのまま返す。
    """
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


def load_tasks():
    """
    data/task_data.xlsx の tasks シートを読み込み、
    タスク一覧データとして返す。
    """
    excel_path = get_task_excel_path()

    if not excel_path.exists():
        return []

    df = pd.read_excel(
        excel_path,
        sheet_name="tasks",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    records = df.to_dict(orient="records")

    for task in records:
        task["due_date"] = format_japanese_date(task.get("due_date", ""))

    return records


def find_task_by_id(task_id):
    """
    task_id に一致するタスクを1件返す。
    見つからない場合は None を返す。
    """
    tasks = load_tasks()

    for task in tasks:
        if str(task.get("id")) == str(task_id):
            return task

    return None


def generate_next_task_id(df):
    """
    既存の TASK-001 形式から次のIDを作る。
    例：
    TASK-001, TASK-002, TASK-003 がある場合は TASK-004 を返す。
    """
    if "id" not in df.columns or df.empty:
        return "TASK-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("TASK-"):
            try:
                number = int(text.replace("TASK-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"TASK-{max_number + 1:03d}"


def update_task_status(task_id, new_status):
    """
    task_data.xlsx の指定タスクの status を更新する。
    更新できたら True、対象がなければ False を返す。
    """
    excel_path = get_task_excel_path()

    if not excel_path.exists():
        return False

    df = pd.read_excel(
        excel_path,
        sheet_name="tasks",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    if "id" not in df.columns or "status" not in df.columns:
        return False

    target_index = df.index[df["id"].astype(str) == str(task_id)].tolist()

    if not target_index:
        return False

    backup_task_excel()

    df.loc[target_index[0], "status"] = str(new_status)

    if "updated_at" in df.columns:
        df.loc[target_index[0], "updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for column in df.columns:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="tasks", index=False)

    return True


def add_task(task_name, category, owner, due_date, status, priority, related_document_id):
    """
    task_data.xlsx に新しいタスクを追加する。
    タスクIDは TASK-001 形式で自動採番する。
    作成できた場合は新しいタスクIDを返す。
    失敗した場合は None を返す。
    """
    excel_path = get_task_excel_path()

    if not excel_path.exists():
        return None

    df = pd.read_excel(
        excel_path,
        sheet_name="tasks",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    required_columns = [
        "id",
        "task_name",
        "category",
        "owner",
        "due_date",
        "status",
        "priority",
        "related_document_id",
        "action_detail",
        "attachment_note",
    ]

    for column in required_columns:
        if column not in df.columns:
            df[column] = ""

    backup_task_excel()

    new_id = generate_next_task_id(df)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = {
        "id": new_id,
        "task_name": task_name,
        "category": category,
        "owner": owner,
        "due_date": due_date,
        "status": status,
        "priority": priority,
        "related_document_id": related_document_id,
    }

    if "created_at" in df.columns:
        new_row["created_at"] = now_text

    if "updated_at" in df.columns:
        new_row["updated_at"] = now_text

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    for column in df.columns:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="tasks", index=False)

    return new_id

def update_task(
    task_id,
    task_name,
    category,
    owner,
    due_date,
    status,
    priority,
    action_detail,
    attachment_note,
):
    """
    task_data.xlsx の指定タスクを更新する。
    更新できたら True、対象がなければ False を返す。
    """
    excel_path = get_task_excel_path()

    if not excel_path.exists():
        return False

    df = pd.read_excel(
        excel_path,
        sheet_name="tasks",
        dtype=str,
    )
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    required_columns = [
        "id",
        "task_name",
        "category",
        "owner",
        "due_date",
        "status",
        "priority",
        "related_document_id",
        "action_detail",
        "attachment_note",
    ]

    for column in required_columns:
        if column not in df.columns:
            df[column] = ""

    target_index = df.index[df["id"].astype(str) == str(task_id)].tolist()

    if not target_index:
        return False

    backup_task_excel()

    index = target_index[0]

    df.loc[index, "task_name"] = str(task_name)
    df.loc[index, "category"] = str(category)
    df.loc[index, "owner"] = str(owner)
    df.loc[index, "due_date"] = str(due_date)
    df.loc[index, "status"] = str(status)
    df.loc[index, "priority"] = str(priority)
    df.loc[index, "action_detail"] = str(action_detail)
    df.loc[index, "attachment_note"] = str(attachment_note)

    if "updated_at" in df.columns:
        df.loc[index, "updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for column in df.columns:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="tasks", index=False)

    return True