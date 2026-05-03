from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings

from tasks.services import add_task


GOVERNANCE_ITEM_COLUMNS = [
    "id",
    "category",
    "item_name",
    "required_level",
    "owner_department",
    "owner",
    "status",
    "latest_version",
    "document_id",
    "file_path",
    "review_frequency",
    "last_review_date",
    "next_review_date",
    "risk_level",
    "remarks",
    "created_at",
    "updated_at",
]


def get_governance_excel_path():
    return Path(settings.BASE_DIR) / "data" / "governance_master.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_governance_excel():
    excel_path = get_governance_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"governance_master_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_governance_excel():
    excel_path = get_governance_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = [
        {
            "id": "GOV-001",
            "category": "会社基本",
            "item_name": "定款",
            "required_level": "法定必須",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "整備済",
            "latest_version": "1.0",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "2026-04-01",
            "next_review_date": "2027-04-01",
            "risk_level": "高",
            "remarks": "会社の根本規則。最新版管理が必要。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-002",
            "category": "会社基本",
            "item_name": "登記簿謄本・履歴事項全部証明書",
            "required_level": "法定必須級",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "要確認",
            "latest_version": "1.0",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "中",
            "remarks": "役員変更、本店移転等の反映状況を確認。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-003",
            "category": "機関運営",
            "item_name": "株主名簿",
            "required_level": "法定必須",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "要確認",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "高",
            "remarks": "株主構成、持株数、移動履歴の確認が必要。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-004",
            "category": "機関運営",
            "item_name": "株主総会議事録",
            "required_level": "法定必須",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "高",
            "remarks": "定時株主総会議事録の保存状況を確認。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-005",
            "category": "機関運営",
            "item_name": "取締役会議事録",
            "required_level": "条件付き必須",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "四半期",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "高",
            "remarks": "取締役会設置会社の場合は重要。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-006",
            "category": "決裁・稟議",
            "item_name": "稟議規程",
            "required_level": "重要",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "要確認",
            "latest_version": "1.0",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-07-31",
            "risk_level": "高",
            "remarks": "決裁権限、金額基準、承認ルートと連動させる。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-007",
            "category": "決裁・稟議",
            "item_name": "職務権限規程",
            "required_level": "重要",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "要確認",
            "latest_version": "1.0",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-07-31",
            "risk_level": "高",
            "remarks": "役職・部署マスタと整合させる。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-008",
            "category": "経理・税務",
            "item_name": "決算書",
            "required_level": "法定必須級",
            "owner_department": "経理部",
            "owner": "経理 五郎",
            "status": "要確認",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-08-31",
            "risk_level": "高",
            "remarks": "貸借対照表、損益計算書などの保存状況。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-009",
            "category": "経理・税務",
            "item_name": "税務申告書",
            "required_level": "法定必須級",
            "owner_department": "経理部",
            "owner": "経理 五郎",
            "status": "要確認",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-08-31",
            "risk_level": "高",
            "remarks": "法人税、消費税、地方税申告関連。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-010",
            "category": "経理・税務",
            "item_name": "証憑保存台帳",
            "required_level": "重要",
            "owner_department": "経理部",
            "owner": "経理 五郎",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "月次",
            "last_review_date": "",
            "next_review_date": "2026-05-31",
            "risk_level": "高",
            "remarks": "請求書、領収書、契約書、納品書等の保存管理。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-011",
            "category": "人事労務",
            "item_name": "就業規則",
            "required_level": "条件付き必須",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "要確認",
            "latest_version": "1.0",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-07-31",
            "risk_level": "高",
            "remarks": "従業員10名以上の場合は届出管理が重要。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-012",
            "category": "人事労務",
            "item_name": "賃金規程",
            "required_level": "重要",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-07-31",
            "risk_level": "中",
            "remarks": "給与、手当、賞与、控除のルール。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-013",
            "category": "契約・取引",
            "item_name": "契約書管理台帳",
            "required_level": "重要",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "月次",
            "last_review_date": "",
            "next_review_date": "2026-05-31",
            "risk_level": "高",
            "remarks": "契約期限、更新日、自動更新条項の管理。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-014",
            "category": "契約・取引",
            "item_name": "取引先管理台帳",
            "required_level": "重要",
            "owner_department": "経営企画部",
            "owner": "経営 太郎",
            "status": "要確認",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "四半期",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "中",
            "remarks": "主要取引先、与信、契約、反社チェックの管理。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-015",
            "category": "個人情報",
            "item_name": "個人情報保護方針",
            "required_level": "重要",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "要確認",
            "latest_version": "1.0",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-07-31",
            "risk_level": "高",
            "remarks": "個人情報の取得、利用、保存、廃棄の基本方針。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-016",
            "category": "情報セキュリティ",
            "item_name": "情報セキュリティ基本規程",
            "required_level": "重要",
            "owner_department": "総務部",
            "owner": "総務 花子",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-07-31",
            "risk_level": "高",
            "remarks": "アカウント、端末、クラウド、情報資産管理と連動。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-017",
            "category": "内部統制",
            "item_name": "内部監査記録",
            "required_level": "推奨",
            "owner_department": "経営企画部",
            "owner": "経営 太郎",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-12-31",
            "risk_level": "中",
            "remarks": "管理体制の成熟度向上に有効。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-018",
            "category": "安全衛生",
            "item_name": "安全衛生管理体制図",
            "required_level": "重要",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "status": "要確認",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "高",
            "remarks": "製造部門管理・安全衛生管理と連動。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-019",
            "category": "環境",
            "item_name": "産業廃棄物委託契約書",
            "required_level": "重要",
            "owner_department": "安全環境部",
            "owner": "安全 四郎",
            "status": "要確認",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-06-30",
            "risk_level": "高",
            "remarks": "廃棄物処理法対応。委託契約とマニフェスト管理。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "GOV-020",
            "category": "BCP",
            "item_name": "事業継続計画 BCP",
            "required_level": "推奨",
            "owner_department": "経営企画部",
            "owner": "経営 太郎",
            "status": "未整備",
            "latest_version": "",
            "document_id": "",
            "file_path": "",
            "review_frequency": "年次",
            "last_review_date": "",
            "next_review_date": "2026-12-31",
            "risk_level": "中",
            "remarks": "災害、感染症、サプライチェーン寸断への備え。",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(rows, columns=GOVERNANCE_ITEM_COLUMNS).to_excel(
            writer,
            sheet_name="governance_items",
            index=False,
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


def read_governance_items_df():
    ensure_governance_excel()
    excel_path = get_governance_excel_path()

    try:
        df = pd.read_excel(
            excel_path,
            sheet_name="governance_items",
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=GOVERNANCE_ITEM_COLUMNS)

    for column in GOVERNANCE_ITEM_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[GOVERNANCE_ITEM_COLUMNS].copy()

    for column in GOVERNANCE_ITEM_COLUMNS:
        df[column] = df[column].astype(str)

    return df


def write_governance_items_df(df):
    excel_path = get_governance_excel_path()

    df = df.copy()

    for column in GOVERNANCE_ITEM_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[GOVERNANCE_ITEM_COLUMNS].copy()

    for column in GOVERNANCE_ITEM_COLUMNS:
        df[column] = df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="governance_items", index=False)


def normalize_governance_item(item):
    for key in ["last_review_date", "next_review_date", "created_at", "updated_at"]:
        if key in item:
            item[key] = format_japanese_date(item.get(key, ""))

    return item


def load_governance_items():
    df = read_governance_items_df()
    records = df.to_dict(orient="records")

    for item in records:
        normalize_governance_item(item)

    return records


def find_governance_item_by_id(item_id):
    items = load_governance_items()

    for item in items:
        if str(item.get("id")) == str(item_id):
            return item

    return None


def update_governance_item_status(item_id, new_status):
    df = read_governance_items_df()

    target_index = df.index[df["id"].astype(str) == str(item_id)].tolist()

    if not target_index:
        return False

    backup_governance_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df.loc[target_index[0], "status"] = new_status
    df.loc[target_index[0], "last_review_date"] = now_text
    df.loc[target_index[0], "updated_at"] = now_text

    write_governance_items_df(df)

    return True


def create_task_from_governance_item(item_id, task_name, owner, due_date, priority):
    item = find_governance_item_by_id(item_id)

    if not item:
        return None

    if not task_name:
        task_name = f"{item.get('item_name', '')}を整備・確認する"

    if not owner:
        owner = item.get("owner", "")

    if not priority:
        priority = item.get("risk_level", "中")

    if priority not in ["高", "中", "低"]:
        priority = "中"

    return add_task(
        task_name=task_name,
        category=f"ガバナンス：{item.get('category', '')}",
        owner=owner,
        due_date=due_date,
        status="未着手",
        priority=priority,
        related_document_id=item.get("document_id", ""),
    )