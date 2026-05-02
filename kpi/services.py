from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings


MONTHLY_KPI_COLUMNS = [
    "id",
    "year_month",
    "sales_amount",
    "gross_profit",
    "operating_profit",
    "cash_balance",
    "accounts_receivable",
    "accounts_payable",
    "expense_total",
    "new_orders",
    "order_backlog",
    "comment",
    "created_at",
    "updated_at",
]

MANUFACTURING_KPI_COLUMNS = [
    "id",
    "year_month",
    "department",
    "production_volume",
    "defect_count",
    "defect_rate",
    "yield_rate",
    "on_time_delivery_rate",
    "equipment_availability",
    "downtime_hours",
    "accident_count",
    "near_miss_count",
    "quality_claim_count",
    "energy_usage",
    "comment",
    "created_at",
    "updated_at",
]


def get_kpi_excel_path():
    return Path(settings.BASE_DIR) / "data" / "kpi_data.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_kpi_excel():
    excel_path = get_kpi_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"kpi_data_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_kpi_excel():
    excel_path = get_kpi_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    monthly_rows = [
        {
            "id": "KPI-202601",
            "year_month": "2026-01",
            "sales_amount": 12500000,
            "gross_profit": 3900000,
            "operating_profit": 850000,
            "cash_balance": 8200000,
            "accounts_receivable": 4300000,
            "accounts_payable": 2600000,
            "expense_total": 3050000,
            "new_orders": 18,
            "order_backlog": 9200000,
            "comment": "年初の受注は安定。経費はやや高め。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "KPI-202602",
            "year_month": "2026-02",
            "sales_amount": 13200000,
            "gross_profit": 4200000,
            "operating_profit": 980000,
            "cash_balance": 8700000,
            "accounts_receivable": 4600000,
            "accounts_payable": 2500000,
            "expense_total": 3220000,
            "new_orders": 21,
            "order_backlog": 9800000,
            "comment": "売上・粗利ともに改善。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "KPI-202603",
            "year_month": "2026-03",
            "sales_amount": 14100000,
            "gross_profit": 4550000,
            "operating_profit": 1120000,
            "cash_balance": 9100000,
            "accounts_receivable": 5100000,
            "accounts_payable": 2800000,
            "expense_total": 3430000,
            "new_orders": 23,
            "order_backlog": 10400000,
            "comment": "年度末需要で売上増。売掛金も増加。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "KPI-202604",
            "year_month": "2026-04",
            "sales_amount": 13600000,
            "gross_profit": 4380000,
            "operating_profit": 1040000,
            "cash_balance": 8900000,
            "accounts_receivable": 4800000,
            "accounts_payable": 2750000,
            "expense_total": 3340000,
            "new_orders": 20,
            "order_backlog": 9900000,
            "comment": "前月比ではやや減少。利益率は維持。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "KPI-202605",
            "year_month": "2026-05",
            "sales_amount": 14800000,
            "gross_profit": 4900000,
            "operating_profit": 1280000,
            "cash_balance": 9500000,
            "accounts_receivable": 5200000,
            "accounts_payable": 2900000,
            "expense_total": 3620000,
            "new_orders": 25,
            "order_backlog": 11100000,
            "comment": "受注好調。粗利率も改善。",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    manufacturing_rows = [
        {
            "id": "MFGKPI-202601",
            "year_month": "2026-01",
            "department": "製造部",
            "production_volume": 9800,
            "defect_count": 145,
            "defect_rate": 1.48,
            "yield_rate": 96.8,
            "on_time_delivery_rate": 94.5,
            "equipment_availability": 91.2,
            "downtime_hours": 18.5,
            "accident_count": 0,
            "near_miss_count": 4,
            "quality_claim_count": 2,
            "energy_usage": 12500,
            "comment": "設備停止がやや多い。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFGKPI-202602",
            "year_month": "2026-02",
            "department": "製造部",
            "production_volume": 10400,
            "defect_count": 132,
            "defect_rate": 1.27,
            "yield_rate": 97.1,
            "on_time_delivery_rate": 95.8,
            "equipment_availability": 92.5,
            "downtime_hours": 15.0,
            "accident_count": 0,
            "near_miss_count": 3,
            "quality_claim_count": 1,
            "energy_usage": 12800,
            "comment": "不良率が改善。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFGKPI-202603",
            "year_month": "2026-03",
            "department": "製造部",
            "production_volume": 11200,
            "defect_count": 150,
            "defect_rate": 1.34,
            "yield_rate": 96.9,
            "on_time_delivery_rate": 96.2,
            "equipment_availability": 93.0,
            "downtime_hours": 14.0,
            "accident_count": 0,
            "near_miss_count": 5,
            "quality_claim_count": 2,
            "energy_usage": 13400,
            "comment": "生産量増。ヒヤリハット報告も増加。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFGKPI-202604",
            "year_month": "2026-04",
            "department": "製造部",
            "production_volume": 10900,
            "defect_count": 128,
            "defect_rate": 1.17,
            "yield_rate": 97.5,
            "on_time_delivery_rate": 96.8,
            "equipment_availability": 94.1,
            "downtime_hours": 11.5,
            "accident_count": 0,
            "near_miss_count": 2,
            "quality_claim_count": 1,
            "energy_usage": 13100,
            "comment": "設備稼働率が改善。",
            "created_at": now_text,
            "updated_at": now_text,
        },
        {
            "id": "MFGKPI-202605",
            "year_month": "2026-05",
            "department": "製造部",
            "production_volume": 11800,
            "defect_count": 119,
            "defect_rate": 1.01,
            "yield_rate": 98.0,
            "on_time_delivery_rate": 97.3,
            "equipment_availability": 95.0,
            "downtime_hours": 9.5,
            "accident_count": 0,
            "near_miss_count": 3,
            "quality_claim_count": 1,
            "energy_usage": 13900,
            "comment": "不良率・稼働率ともに改善。",
            "created_at": now_text,
            "updated_at": now_text,
        },
    ]

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(monthly_rows, columns=MONTHLY_KPI_COLUMNS).to_excel(
            writer, sheet_name="monthly_kpis", index=False
        )
        pd.DataFrame(manufacturing_rows, columns=MANUFACTURING_KPI_COLUMNS).to_excel(
            writer, sheet_name="manufacturing_kpis", index=False
        )


def read_sheet(sheet_name, columns):
    ensure_kpi_excel()
    excel_path = get_kpi_excel_path()

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        df = df.fillna("")
    except Exception:
        df = pd.DataFrame(columns=columns)

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    return df[columns]


def to_number(value):
    if value == "" or value is None:
        return 0

    try:
        return float(str(value).replace(",", ""))
    except Exception:
        return 0


def format_amount(value):
    try:
        number = int(float(value))
        return f"{number:,}"
    except Exception:
        return value


def format_percent(value):
    if value == "" or value is None:
        return ""

    try:
        number = float(value)
        return f"{number:.1f}%"
    except Exception:
        return value


def enrich_monthly_kpi(row):
    sales_amount = to_number(row.get("sales_amount"))
    gross_profit = to_number(row.get("gross_profit"))
    operating_profit = to_number(row.get("operating_profit"))

    if sales_amount > 0:
        row["gross_profit_rate"] = round(gross_profit / sales_amount * 100, 1)
        row["operating_profit_rate"] = round(operating_profit / sales_amount * 100, 1)
    else:
        row["gross_profit_rate"] = 0
        row["operating_profit_rate"] = 0

    amount_fields = [
        "sales_amount",
        "gross_profit",
        "operating_profit",
        "cash_balance",
        "accounts_receivable",
        "accounts_payable",
        "expense_total",
        "order_backlog",
    ]

    for field in amount_fields:
        row[f"{field}_display"] = format_amount(row.get(field, ""))

    row["gross_profit_rate_display"] = format_percent(row.get("gross_profit_rate"))
    row["operating_profit_rate_display"] = format_percent(row.get("operating_profit_rate"))

    return row


def enrich_manufacturing_kpi(row):
    percent_fields = [
        "defect_rate",
        "yield_rate",
        "on_time_delivery_rate",
        "equipment_availability",
    ]

    for field in percent_fields:
        row[f"{field}_display"] = format_percent(row.get(field, ""))

    return row


def load_monthly_kpis():
    df = read_sheet("monthly_kpis", MONTHLY_KPI_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        enrich_monthly_kpi(row)

    records = sorted(
        records,
        key=lambda x: str(x.get("year_month", "")),
        reverse=True,
    )

    return records


def load_manufacturing_kpis():
    df = read_sheet("manufacturing_kpis", MANUFACTURING_KPI_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        enrich_manufacturing_kpi(row)

    records = sorted(
        records,
        key=lambda x: str(x.get("year_month", "")),
        reverse=True,
    )

    return records


def find_monthly_kpi_by_id(kpi_id):
    for row in load_monthly_kpis():
        if str(row.get("id")) == str(kpi_id):
            return row
    return None


def find_manufacturing_kpi_by_id(kpi_id):
    for row in load_manufacturing_kpis():
        if str(row.get("id")) == str(kpi_id):
            return row
    return None


def get_latest_monthly_kpi():
    records = load_monthly_kpis()
    if records:
        return records[0]
    return None


def get_latest_manufacturing_kpi():
    records = load_manufacturing_kpis()
    if records:
        return records[0]
    return None