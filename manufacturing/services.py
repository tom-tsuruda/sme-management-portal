from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings

from tasks.services import add_task


MANAGEMENT_ITEM_COLUMNS = [
    "id",
    "area",
    "category",
    "item_name",
    "description",
    "owner_department",
    "owner",
    "check_frequency",
    "required_document_id",
    "related_law_or_standard",
    "risk_level",
    "status",
    "last_checked_at",
    "next_check_date",
    "created_at",
    "updated_at",
]


def get_manufacturing_excel_path():
    return Path(settings.BASE_DIR) / "data" / "manufacturing_master.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_manufacturing_excel():
    excel_path = get_manufacturing_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"manufacturing_master_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_manufacturing_excel():
    excel_path = get_manufacturing_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = [
        {
            "id": "MFG-001",
            "area": "品質管理",
            "category": "検査",
            "item_name": "検査基準書の整備",
            "description": "受入検査・工程内検査・出荷検査の基準が文書化されているか確認する。",
            "owner_department": "品質管理部",
            "owner": "品質 次郎",
            "check_frequency": "四半期",
            "required_document_id": "",
            "related_law_or_standard": "ISO9001相当",
            "risk_level": "高",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-002",
            "area": "品質管理",
            "category": "不適合",
            "item_name": "不適合品管理台帳",
            "description": "不適合品の発生、隔離、処置、再発防止が記録されているか確認する。",
            "owner_department": "品質管理部",
            "owner": "品質 次郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "ISO9001相当",
            "risk_level": "高",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-003",
            "area": "品質管理",
            "category": "クレーム",
            "item_name": "品質クレーム対応記録",
            "description": "顧客クレームの受付、原因分析、是正処置、効果確認を管理する。",
            "owner_department": "品質管理部",
            "owner": "品質 次郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "高",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-004",
            "area": "品質管理",
            "category": "計測器",
            "item_name": "計測器管理台帳",
            "description": "計測器の校正期限、管理番号、使用部署を管理する。",
            "owner_department": "品質管理部",
            "owner": "品質 次郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "計測器管理",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-005",
            "area": "安全衛生",
            "category": "リスクアセスメント",
            "item_name": "リスクアセスメント実施状況",
            "description": "危険源の洗い出し、評価、対策の実施状況を確認する。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "半期",
            "required_document_id": "",
            "related_law_or_standard": "労働安全衛生法",
            "risk_level": "高",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-006",
            "area": "安全衛生",
            "category": "ヒヤリハット",
            "item_name": "ヒヤリハット記録",
            "description": "ヒヤリハット情報を収集し、再発防止・事故予防につなげる。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "労働安全衛生",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-007",
            "area": "安全衛生",
            "category": "教育",
            "item_name": "安全教育記録",
            "description": "新入社員教育、作業変更時教育、定期安全教育の記録を管理する。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "四半期",
            "required_document_id": "",
            "related_law_or_standard": "労働安全衛生法",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-008",
            "area": "設備保全",
            "category": "設備台帳",
            "item_name": "設備台帳",
            "description": "主要設備の取得日、型式、保全担当、点検周期を管理する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "四半期",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "中",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-009",
            "area": "設備保全",
            "category": "日常点検",
            "item_name": "日常点検記録",
            "description": "設備の日常点検が実施され、記録が保存されているか確認する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-010",
            "area": "設備保全",
            "category": "故障管理",
            "item_name": "設備故障履歴",
            "description": "設備故障の発生日、原因、修理内容、再発防止を記録する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-011",
            "area": "化学物質",
            "category": "SDS",
            "item_name": "SDS管理",
            "description": "使用化学物質のSDSを最新版で管理し、必要部署が閲覧できる状態にする。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "四半期",
            "required_document_id": "",
            "related_law_or_standard": "化学物質管理",
            "risk_level": "高",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-012",
            "area": "化学物質",
            "category": "リスクアセスメント",
            "item_name": "化学物質リスクアセスメント",
            "description": "化学物質の使用、保管、廃棄に関するリスク評価を管理する。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "半期",
            "required_document_id": "",
            "related_law_or_standard": "労働安全衛生法",
            "risk_level": "高",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-013",
            "area": "環境管理",
            "category": "産業廃棄物",
            "item_name": "産業廃棄物管理",
            "description": "産業廃棄物の委託先、マニフェスト、保管状況を管理する。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "廃棄物処理法",
            "risk_level": "高",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-014",
            "area": "環境管理",
            "category": "排水・排気",
            "item_name": "排水・排気管理",
            "description": "排水、排気、騒音等の測定記録と基準値超過の有無を管理する。",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "check_frequency": "四半期",
            "required_document_id": "",
            "related_law_or_standard": "環境関連法令",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-015",
            "area": "研究開発",
            "category": "テーマ管理",
            "item_name": "研究テーマ管理",
            "description": "研究テーマ、責任者、進捗、課題、成果を管理する。",
            "owner_department": "研究開発部",
            "owner": "研究 三郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-016",
            "area": "研究開発",
            "category": "試作",
            "item_name": "試作評価記録",
            "description": "試作品の評価条件、評価結果、改善事項、量産移管可否を管理する。",
            "owner_department": "研究開発部",
            "owner": "研究 三郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "中",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-017",
            "area": "製造管理",
            "category": "標準化",
            "item_name": "作業標準書",
            "description": "主要工程の作業手順、注意点、品質ポイントを文書化する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "四半期",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "高",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-018",
            "area": "製造管理",
            "category": "工程",
            "item_name": "工程フロー図",
            "description": "製造工程、検査工程、外注工程の流れを可視化する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "半期",
            "required_document_id": "",
            "related_law_or_standard": "",
            "risk_level": "中",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-019",
            "area": "IT/OT",
            "category": "システム管理",
            "item_name": "工場システム一覧",
            "description": "工場で利用しているPC、制御システム、ネットワーク機器を一覧化する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "半期",
            "required_document_id": "",
            "related_law_or_standard": "情報セキュリティ",
            "risk_level": "中",
            "status": "未整備",
            "last_checked_at": "",
            "next_check_date": "2026-06-30",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFG-020",
            "area": "IT/OT",
            "category": "バックアップ",
            "item_name": "工場システムバックアップ確認",
            "description": "生産管理・検査・制御系データのバックアップ状況を確認する。",
            "owner_department": "製造部",
            "owner": "製造 一郎",
            "check_frequency": "月次",
            "required_document_id": "",
            "related_law_or_standard": "情報セキュリティ",
            "risk_level": "高",
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "2026-05-31",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(rows, columns=MANAGEMENT_ITEM_COLUMNS).to_excel(
            writer, sheet_name="management_items", index=False
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


def read_management_items_df():
    ensure_manufacturing_excel()
    excel_path = get_manufacturing_excel_path()

    try:
        df = pd.read_excel(excel_path, sheet_name="management_items")
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=MANAGEMENT_ITEM_COLUMNS)

    for column in MANAGEMENT_ITEM_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    return df[MANAGEMENT_ITEM_COLUMNS]


def write_management_items_df(df):
    excel_path = get_manufacturing_excel_path()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="management_items", index=False)


def normalize_management_item(item):
    for key in ["last_checked_at", "next_check_date", "created_at", "updated_at"]:
        if key in item:
            item[key] = format_japanese_date(item.get(key, ""))

    return item


def load_management_items():
    df = read_management_items_df()
    records = df.to_dict(orient="records")

    for item in records:
        normalize_management_item(item)

    return records


def find_management_item_by_id(item_id):
    items = load_management_items()

    for item in items:
        if str(item.get("id")) == str(item_id):
            return item

    return None


def update_management_item_status(item_id, new_status):
    df = read_management_items_df()

    target_index = df.index[df["id"].astype(str) == str(item_id)].tolist()

    if not target_index:
        return False

    backup_manufacturing_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df.loc[target_index[0], "status"] = new_status
    df.loc[target_index[0], "last_checked_at"] = now_text
    df.loc[target_index[0], "updated_at"] = now_text

    write_management_items_df(df)

    return True


def create_task_from_management_item(item_id, task_name, owner, due_date, priority):
    item = find_management_item_by_id(item_id)

    if not item:
        return None

    if not task_name:
        task_name = f"{item.get('item_name', '')}を整備・確認する"

    if not owner:
        owner = item.get("owner", "")

    if not priority:
        priority = "中"

    return add_task(
        task_name=task_name,
        category=f"製造管理：{item.get('area', '')}",
        owner=owner,
        due_date=due_date,
        status="未着手",
        priority=priority,
        related_document_id=item.get("required_document_id", ""),
    )