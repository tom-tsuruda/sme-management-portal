from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings

from tasks.services import add_task


MANAGEMENT_ITEM_COLUMNS = [
    "id",
    "template_id",
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

MONITORING_RECORD_COLUMNS = [
    "id",
    "management_item_id",
    "checked_at",
    "checked_by",
    "result",
    "score",
    "comment",
    "evidence_file",
    "next_action",
    "created_at",
    "updated_at",
]

INCIDENT_COLUMNS = [
    "id",
    "area",
    "incident_date",
    "title",
    "description",
    "severity",
    "detected_by",
    "owner",
    "status",
    "corrective_action",
    "preventive_action",
    "related_management_item_id",
    "related_task_id",
    "created_at",
    "updated_at",
]

MANAGEMENT_TEMPLATE_COLUMNS = [
    "id",
    "area",
    "category",
    "template_name",
    "description",
    "recommended_frequency",
    "recommended_owner_department",
    "related_law_or_standard",
    "risk_level",
    "is_default",
    "is_active",
]


DEFAULT_MANAGEMENT_TEMPLATE_ROWS = [
    {
        "id": "MTPL-001",
        "area": "品質管理",
        "category": "品質",
        "template_name": "不良率の月次確認",
        "description": "月次で不良率、不良件数、主な不良内容を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "品質管理部",
        "related_law_or_standard": "ISO9001相当",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-002",
        "area": "品質管理",
        "category": "検査",
        "template_name": "出荷検査結果の確認",
        "description": "出荷検査の実施状況、不合格件数、再検査状況を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "品質管理部",
        "related_law_or_standard": "ISO9001相当",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-003",
        "area": "品質管理",
        "category": "クレーム",
        "template_name": "品質クレーム発生状況の確認",
        "description": "顧客クレームの件数、原因、是正処置、再発防止策を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "品質管理部",
        "related_law_or_standard": "",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-004",
        "area": "設備保全",
        "category": "設備",
        "template_name": "主要設備の点検実施状況確認",
        "description": "主要設備の日常点検、定期点検、未実施項目を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "製造部",
        "related_law_or_standard": "",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-005",
        "area": "設備保全",
        "category": "故障",
        "template_name": "設備故障履歴の確認",
        "description": "設備故障の発生件数、停止時間、原因、再発防止策を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "製造部",
        "related_law_or_standard": "",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-006",
        "area": "安全衛生",
        "category": "安全",
        "template_name": "ヒヤリハット件数確認",
        "description": "ヒヤリハットの件数、内容、対策状況を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "安全環境部",
        "related_law_or_standard": "労働安全衛生",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-007",
        "area": "安全衛生",
        "category": "教育",
        "template_name": "安全教育実施状況確認",
        "description": "安全教育、作業変更時教育、定期教育の実施状況を確認する。",
        "recommended_frequency": "四半期",
        "recommended_owner_department": "安全環境部",
        "related_law_or_standard": "労働安全衛生法",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-008",
        "area": "化学物質",
        "category": "SDS",
        "template_name": "SDS最新版確認",
        "description": "使用化学物質のSDSが最新版で管理されているか確認する。",
        "recommended_frequency": "四半期",
        "recommended_owner_department": "安全環境部",
        "related_law_or_standard": "化学物質管理",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-009",
        "area": "化学物質",
        "category": "リスクアセスメント",
        "template_name": "化学物質リスクアセスメント確認",
        "description": "化学物質の使用、保管、廃棄に関するリスク評価を確認する。",
        "recommended_frequency": "半期",
        "recommended_owner_department": "安全環境部",
        "related_law_or_standard": "労働安全衛生法",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-010",
        "area": "環境管理",
        "category": "廃棄物",
        "template_name": "産業廃棄物管理状況確認",
        "description": "委託先、マニフェスト、保管状況、処理記録を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "安全環境部",
        "related_law_or_standard": "廃棄物処理法",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-011",
        "area": "製造管理",
        "category": "生産",
        "template_name": "生産計画と実績の確認",
        "description": "生産計画、実績、差異、遅延理由を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "製造部",
        "related_law_or_standard": "",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-012",
        "area": "製造管理",
        "category": "納期",
        "template_name": "納期遅延状況確認",
        "description": "納期遵守率、遅延案件、原因、対策を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "製造部",
        "related_law_or_standard": "",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-013",
        "area": "研究開発",
        "category": "テーマ管理",
        "template_name": "研究開発テーマ進捗確認",
        "description": "研究テーマごとの進捗、課題、成果、次アクションを確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "研究開発部",
        "related_law_or_standard": "",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-014",
        "area": "IT/OT",
        "category": "バックアップ",
        "template_name": "工場システムバックアップ確認",
        "description": "生産管理、検査、制御系データのバックアップ状況を確認する。",
        "recommended_frequency": "月次",
        "recommended_owner_department": "製造部",
        "related_law_or_standard": "情報セキュリティ",
        "risk_level": "高",
        "is_default": "1",
        "is_active": "1",
    },
    {
        "id": "MTPL-015",
        "area": "IT/OT",
        "category": "セキュリティ",
        "template_name": "工場PC・制御機器の管理状況確認",
        "description": "工場PC、制御機器、ネットワーク機器の管理状況を確認する。",
        "recommended_frequency": "四半期",
        "recommended_owner_department": "製造部",
        "related_law_or_standard": "情報セキュリティ",
        "risk_level": "中",
        "is_default": "1",
        "is_active": "1",
    },
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

        pd.DataFrame(columns=MONITORING_RECORD_COLUMNS).to_excel(
            writer, sheet_name="monitoring_records", index=False
        )

        pd.DataFrame(columns=INCIDENT_COLUMNS).to_excel(
            writer, sheet_name="incidents", index=False
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
        df = pd.read_excel(
            excel_path,
            sheet_name="management_items",
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=MANAGEMENT_ITEM_COLUMNS)

    for column in MANAGEMENT_ITEM_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[MANAGEMENT_ITEM_COLUMNS].copy()

    for column in MANAGEMENT_ITEM_COLUMNS:
        df[column] = df[column].astype(str)

    return df

def read_monitoring_records_df():
    ensure_manufacturing_excel()
    excel_path = get_manufacturing_excel_path()

    try:
        df = pd.read_excel(
            excel_path,
            sheet_name="monitoring_records",
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=MONITORING_RECORD_COLUMNS)

    for column in MONITORING_RECORD_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[MONITORING_RECORD_COLUMNS].copy()

    for column in MONITORING_RECORD_COLUMNS:
        df[column] = df[column].astype(str)

    return df


def write_monitoring_records_df(monitoring_df):
    excel_path = get_manufacturing_excel_path()

    management_df = read_management_items_df()
    incident_df = read_incidents_df()

    monitoring_df = monitoring_df.copy()

    for column in MONITORING_RECORD_COLUMNS:
        if column not in monitoring_df.columns:
            monitoring_df[column] = ""

    monitoring_df = monitoring_df[MONITORING_RECORD_COLUMNS].copy()

    for column in MONITORING_RECORD_COLUMNS:
        monitoring_df[column] = monitoring_df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        management_df.to_excel(writer, sheet_name="management_items", index=False)
        monitoring_df.to_excel(writer, sheet_name="monitoring_records", index=False)
        incident_df.to_excel(writer, sheet_name="incidents", index=False)


def generate_next_monitoring_record_id(df):
    if "id" not in df.columns or df.empty:
        return "REC-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("REC-"):
            try:
                number = int(text.replace("REC-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"REC-{max_number + 1:03d}"


def normalize_monitoring_record(record):
    for key in ["checked_at", "created_at", "updated_at"]:
        if key in record:
            record[key] = format_japanese_date(record.get(key, ""))

    return record


def load_monitoring_records():
    df = read_monitoring_records_df()
    records = df.to_dict(orient="records")

    for record in records:
        normalize_monitoring_record(record)

    return records


def find_monitoring_records_by_item_id(item_id):
    records = load_monitoring_records()

    return [
        record for record in records
        if str(record.get("management_item_id")) == str(item_id)
    ]


def create_monitoring_record(
    management_item_id,
    checked_at,
    checked_by,
    result,
    score,
    comment,
    evidence_file,
    next_action,
):
    monitoring_df = read_monitoring_records_df()
    management_df = read_management_items_df()

    backup_manufacturing_excel()

    record_id = generate_next_monitoring_record_id(monitoring_df)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = {
        "id": record_id,
        "management_item_id": management_item_id,
        "checked_at": checked_at,
        "checked_by": checked_by,
        "result": result,
        "score": score,
        "comment": comment,
        "evidence_file": evidence_file,
        "next_action": next_action,
        "created_at": now_text,
        "updated_at": now_text,
    }

    monitoring_df = pd.concat(
        [monitoring_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    target_index = management_df.index[
        management_df["id"].astype(str) == str(management_item_id)
    ].tolist()

    if target_index:
        management_df.loc[target_index[0], "last_checked_at"] = checked_at or now_text
        management_df.loc[target_index[0], "updated_at"] = now_text

        if result in ["要確認", "要改善", "NG", "異常"]:
            management_df.loc[target_index[0], "status"] = "要改善"
        elif result in ["OK", "良好", "問題なし"]:
            management_df.loc[target_index[0], "status"] = "整備済"

    with pd.ExcelWriter(get_manufacturing_excel_path(), engine="openpyxl", mode="w") as writer:
        management_df.to_excel(writer, sheet_name="management_items", index=False)
        monitoring_df.to_excel(writer, sheet_name="monitoring_records", index=False)

    return record_id

def read_incidents_df():
    ensure_manufacturing_excel()
    excel_path = get_manufacturing_excel_path()

    try:
        df = pd.read_excel(
            excel_path,
            sheet_name="incidents",
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=INCIDENT_COLUMNS)

    for column in INCIDENT_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[INCIDENT_COLUMNS].copy()

    for column in INCIDENT_COLUMNS:
        df[column] = df[column].astype(str)

    return df


def write_incidents_df(incident_df):
    excel_path = get_manufacturing_excel_path()

    management_df = read_management_items_df()
    monitoring_df = read_monitoring_records_df()

    incident_df = incident_df.copy()

    for column in INCIDENT_COLUMNS:
        if column not in incident_df.columns:
            incident_df[column] = ""

    incident_df = incident_df[INCIDENT_COLUMNS].copy()

    for column in INCIDENT_COLUMNS:
        incident_df[column] = incident_df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        management_df.to_excel(writer, sheet_name="management_items", index=False)
        monitoring_df.to_excel(writer, sheet_name="monitoring_records", index=False)
        incident_df.to_excel(writer, sheet_name="incidents", index=False)


def generate_next_incident_id(df):
    if "id" not in df.columns or df.empty:
        return "INC-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("INC-"):
            try:
                number = int(text.replace("INC-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"INC-{max_number + 1:03d}"


def normalize_incident(incident):
    for key in ["incident_date", "created_at", "updated_at"]:
        if key in incident:
            incident[key] = format_japanese_date(incident.get(key, ""))

    return incident


def load_incidents():
    df = read_incidents_df()
    records = df.to_dict(orient="records")

    for incident in records:
        normalize_incident(incident)

    return records


def find_incident_by_id(incident_id):
    incidents = load_incidents()

    for incident in incidents:
        if str(incident.get("id")) == str(incident_id):
            return incident

    return None


def create_incident(
    area,
    incident_date,
    title,
    description,
    severity,
    detected_by,
    owner,
    status,
    corrective_action,
    preventive_action,
    related_management_item_id,
    create_task=False,
    due_date="",
):
    incident_df = read_incidents_df()

    backup_manufacturing_excel()

    incident_id = generate_next_incident_id(incident_df)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    related_task_id = ""

    if create_task and title:
        task_priority = "高" if severity in ["重大", "高"] else "中"

        related_task_id = add_task(
            task_name=f"インシデント対応：{title}",
            category=f"製造インシデント：{area}",
            owner=owner,
            due_date=due_date,
            status="未着手",
            priority=task_priority,
            related_document_id="",
        ) or ""

    new_row = {
        "id": incident_id,
        "area": area,
        "incident_date": incident_date,
        "title": title,
        "description": description,
        "severity": severity,
        "detected_by": detected_by,
        "owner": owner,
        "status": status or "未対応",
        "corrective_action": corrective_action,
        "preventive_action": preventive_action,
        "related_management_item_id": related_management_item_id,
        "related_task_id": related_task_id,
        "created_at": now_text,
        "updated_at": now_text,
    }

    incident_df = pd.concat(
        [incident_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_incidents_df(incident_df)

    return incident_id, related_task_id


def write_management_items_df(df):
    excel_path = get_manufacturing_excel_path()

    management_df = df.copy()

    for column in MANAGEMENT_ITEM_COLUMNS:
        if column not in management_df.columns:
            management_df[column] = ""

    management_df = management_df[MANAGEMENT_ITEM_COLUMNS].copy()

    for column in MANAGEMENT_ITEM_COLUMNS:
        management_df[column] = management_df[column].astype(str)

    monitoring_df = read_monitoring_records_df()
    incident_df = read_incidents_df()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        management_df.to_excel(writer, sheet_name="management_items", index=False)
        monitoring_df.to_excel(writer, sheet_name="monitoring_records", index=False)
        incident_df.to_excel(writer, sheet_name="incidents", index=False)


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

def read_management_templates_df():
    ensure_manufacturing_excel()
    excel_path = get_manufacturing_excel_path()

    try:
        df = pd.read_excel(
            excel_path,
            sheet_name="management_templates",
            dtype=str,
        )
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(DEFAULT_MANAGEMENT_TEMPLATE_ROWS)

        for column in MANAGEMENT_TEMPLATE_COLUMNS:
            if column not in df.columns:
                df[column] = ""

        df = df[MANAGEMENT_TEMPLATE_COLUMNS].copy()

        management_df = read_management_items_df()
        monitoring_df = read_monitoring_records_df()
        incident_df = read_incidents_df()

        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
            management_df.to_excel(writer, sheet_name="management_items", index=False)
            monitoring_df.to_excel(writer, sheet_name="monitoring_records", index=False)
            incident_df.to_excel(writer, sheet_name="incidents", index=False)
            df.to_excel(writer, sheet_name="management_templates", index=False)

    for column in MANAGEMENT_TEMPLATE_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[MANAGEMENT_TEMPLATE_COLUMNS].copy()

    for column in MANAGEMENT_TEMPLATE_COLUMNS:
        df[column] = df[column].astype(str)

    return df


def write_all_manufacturing_sheets(
    management_df=None,
    monitoring_df=None,
    incident_df=None,
    template_df=None,
):
    excel_path = get_manufacturing_excel_path()

    if management_df is None:
        management_df = read_management_items_df()

    if monitoring_df is None:
        monitoring_df = read_monitoring_records_df()

    if incident_df is None:
        incident_df = read_incidents_df()

    if template_df is None:
        template_df = read_management_templates_df()

    for column in MANAGEMENT_ITEM_COLUMNS:
        if column not in management_df.columns:
            management_df[column] = ""

    for column in MONITORING_RECORD_COLUMNS:
        if column not in monitoring_df.columns:
            monitoring_df[column] = ""

    for column in INCIDENT_COLUMNS:
        if column not in incident_df.columns:
            incident_df[column] = ""

    for column in MANAGEMENT_TEMPLATE_COLUMNS:
        if column not in template_df.columns:
            template_df[column] = ""

    management_df = management_df[MANAGEMENT_ITEM_COLUMNS].copy()
    monitoring_df = monitoring_df[MONITORING_RECORD_COLUMNS].copy()
    incident_df = incident_df[INCIDENT_COLUMNS].copy()
    template_df = template_df[MANAGEMENT_TEMPLATE_COLUMNS].copy()

    for column in MANAGEMENT_ITEM_COLUMNS:
        management_df[column] = management_df[column].astype(str)

    for column in MONITORING_RECORD_COLUMNS:
        monitoring_df[column] = monitoring_df[column].astype(str)

    for column in INCIDENT_COLUMNS:
        incident_df[column] = incident_df[column].astype(str)

    for column in MANAGEMENT_TEMPLATE_COLUMNS:
        template_df[column] = template_df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        management_df.to_excel(writer, sheet_name="management_items", index=False)
        monitoring_df.to_excel(writer, sheet_name="monitoring_records", index=False)
        incident_df.to_excel(writer, sheet_name="incidents", index=False)
        template_df.to_excel(writer, sheet_name="management_templates", index=False)


def load_management_templates():
    template_df = read_management_templates_df()
    management_df = read_management_items_df()

    selected_template_ids = set(
        str(value).strip()
        for value in management_df.get("template_id", pd.Series(dtype=str)).fillna("")
        if str(value).strip()
    )

    records = template_df.to_dict(orient="records")

    for record in records:
        record["is_selected"] = str(record.get("id", "")).strip() in selected_template_ids

    return records


def generate_next_management_item_id_from_df(df):
    if "id" not in df.columns or df.empty:
        return "MFG-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("MFG-"):
            try:
                number = int(text.replace("MFG-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"MFG-{max_number + 1:03d}"


def create_management_items_from_templates(template_ids):
    if not template_ids:
        return 0

    template_ids = [
        str(template_id).strip()
        for template_id in template_ids
        if str(template_id).strip()
    ]

    if not template_ids:
        return 0

    template_df = read_management_templates_df()
    management_df = read_management_items_df()

    existing_template_ids = set(
        str(value).strip()
        for value in management_df.get("template_id", pd.Series(dtype=str)).fillna("")
        if str(value).strip()
    )

    backup_manufacturing_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []

    for template_id in template_ids:
        if template_id in existing_template_ids:
            continue

        matched_templates = template_df[
            template_df["id"].astype(str) == str(template_id)
        ]

        if matched_templates.empty:
            continue

        template = matched_templates.iloc[0].to_dict()

        current_df = pd.concat(
            [management_df, pd.DataFrame(new_rows)],
            ignore_index=True,
        )

        item_id = generate_next_management_item_id_from_df(current_df)

        new_rows.append({
            "id": item_id,
            "template_id": template.get("id", ""),
            "area": template.get("area", ""),
            "category": template.get("category", ""),
            "item_name": template.get("template_name", ""),
            "description": template.get("description", ""),
            "owner_department": template.get("recommended_owner_department", ""),
            "owner": "",
            "check_frequency": template.get("recommended_frequency", ""),
            "required_document_id": "",
            "related_law_or_standard": template.get("related_law_or_standard", ""),
            "risk_level": template.get("risk_level", "中"),
            "status": "要確認",
            "last_checked_at": "",
            "next_check_date": "",
            "created_at": now_text,
            "updated_at": now_text,
        })

    if not new_rows:
        return 0

    management_df = pd.concat(
        [management_df, pd.DataFrame(new_rows)],
        ignore_index=True,
    )

    write_all_manufacturing_sheets(
        management_df=management_df,
        template_df=template_df,
    )

    return len(new_rows)