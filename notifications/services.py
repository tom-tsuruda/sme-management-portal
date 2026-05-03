from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings


NOTIFICATION_COLUMNS = [
    "id",
    "title",
    "message",
    "target_user",
    "category",
    "priority",
    "related_type",
    "related_id",
    "is_read",
    "created_at",
    "read_at",
]


def get_notification_excel_path():
    return Path(settings.BASE_DIR) / "data" / "notification_data.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_notification_excel():
    excel_path = get_notification_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"notification_data_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_notification_excel():
    excel_path = get_notification_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = [
        {
            "id": "NOTIF-001",
            "title": "承認待ち申請があります",
            "message": "申請・承認画面で承認待ちの申請を確認してください。",
            "target_user": "承認者",
            "category": "申請承認",
            "priority": "高",
            "related_type": "workflows",
            "related_id": "",
            "is_read": "0",
            "created_at": now_text,
            "read_at": "",
        },
        {
            "id": "NOTIF-002",
            "title": "製造管理項目に未整備があります",
            "message": "製造管理画面で、品質・安全・設備などの未整備項目を確認してください。",
            "target_user": "製造責任者",
            "category": "製造管理",
            "priority": "高",
            "related_type": "manufacturing",
            "related_id": "",
            "is_read": "0",
            "created_at": now_text,
            "read_at": "",
        },
        {
            "id": "NOTIF-003",
            "title": "ガバナンス文書に要確認項目があります",
            "message": "ガバナンス文書台帳で、法定必須級・高リスク文書の状態を確認してください。",
            "target_user": "管理者",
            "category": "ガバナンス",
            "priority": "高",
            "related_type": "governance",
            "related_id": "",
            "is_read": "0",
            "created_at": now_text,
            "read_at": "",
        },
        {
            "id": "NOTIF-004",
            "title": "経費申請の確認が必要です",
            "message": "申請中または差戻しの経費申請があります。経費・旅費精算画面を確認してください。",
            "target_user": "経理責任者",
            "category": "経費精算",
            "priority": "中",
            "related_type": "expenses",
            "related_id": "",
            "is_read": "1",
            "created_at": now_text,
            "read_at": now_text,
        },
        {
            "id": "NOTIF-005",
            "title": "文書レビュー期限が近づいています",
            "message": "文書台帳の次回確認日を確認し、必要に応じて改定タスクを作成してください。",
            "target_user": "総務部",
            "category": "文書管理",
            "priority": "中",
            "related_type": "documents",
            "related_id": "",
            "is_read": "0",
            "created_at": now_text,
            "read_at": "",
        },
    ]

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(rows, columns=NOTIFICATION_COLUMNS).to_excel(
            writer,
            sheet_name="notifications",
            index=False,
        )


def format_japanese_datetime(value):
    if value == "" or value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    try:
        dt = pd.to_datetime(value)
        return f"{dt.year}年{dt.month}月{dt.day}日 {dt.hour:02d}:{dt.minute:02d}"
    except Exception:
        return value


def read_notifications_df():
    ensure_notification_excel()
    excel_path = get_notification_excel_path()

    try:
        df = pd.read_excel(
            excel_path,
            sheet_name="notifications",
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=NOTIFICATION_COLUMNS)

    for column in NOTIFICATION_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[NOTIFICATION_COLUMNS].copy()

    for column in NOTIFICATION_COLUMNS:
        df[column] = df[column].astype(str)

    return df


def write_notifications_df(df):
    excel_path = get_notification_excel_path()

    df = df.copy()

    for column in NOTIFICATION_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[NOTIFICATION_COLUMNS].copy()

    for column in NOTIFICATION_COLUMNS:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="notifications", index=False)


def normalize_notification(notification):
    notification["created_at"] = format_japanese_datetime(notification.get("created_at", ""))
    notification["read_at"] = format_japanese_datetime(notification.get("read_at", ""))

    return notification


def load_notifications():
    df = read_notifications_df()
    records = df.to_dict(orient="records")

    for notification in records:
        normalize_notification(notification)

    records = sorted(
        records,
        key=lambda x: str(x.get("created_at", "")),
        reverse=True,
    )

    return records


def find_notification_by_id(notification_id):
    notifications = load_notifications()

    for notification in notifications:
        if str(notification.get("id")) == str(notification_id):
            return notification

    return None


def generate_next_notification_id(df):
    if "id" not in df.columns or df.empty:
        return "NOTIF-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("NOTIF-"):
            try:
                number = int(text.replace("NOTIF-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"NOTIF-{max_number + 1:03d}"


def create_notification(
    title,
    message,
    target_user,
    category,
    priority,
    related_type="",
    related_id="",
):
    df = read_notifications_df()

    backup_notification_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    notification_id = generate_next_notification_id(df)

    new_row = {
        "id": notification_id,
        "title": title,
        "message": message,
        "target_user": target_user,
        "category": category,
        "priority": priority,
        "related_type": related_type,
        "related_id": related_id,
        "is_read": "0",
        "created_at": now_text,
        "read_at": "",
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    write_notifications_df(df)

    return notification_id


def update_notification_read_status(notification_id, is_read):
    df = read_notifications_df()

    target_index = df.index[df["id"].astype(str) == str(notification_id)].tolist()

    if not target_index:
        return False

    backup_notification_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df.loc[target_index[0], "is_read"] = "1" if is_read else "0"
    df.loc[target_index[0], "read_at"] = now_text if is_read else ""

    write_notifications_df(df)

    return True


def mark_notification_as_read(notification_id):
    return update_notification_read_status(notification_id, True)


def mark_notification_as_unread(notification_id):
    return update_notification_read_status(notification_id, False)


def load_unread_notifications():
    notifications = load_notifications()

    return [
        notification for notification in notifications
        if str(notification.get("is_read")) not in ["1", "True", "true", "既読"]
    ]