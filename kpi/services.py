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
            writer,
            sheet_name="monthly_kpis",
            index=False,
        )
        pd.DataFrame(manufacturing_rows, columns=MANUFACTURING_KPI_COLUMNS).to_excel(
            writer,
            sheet_name="manufacturing_kpis",
            index=False,
        )


def read_sheet(sheet_name, columns):
    ensure_kpi_excel()
    excel_path = get_kpi_excel_path()

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


def write_kpi_excel(monthly_df=None, manufacturing_df=None):
    ensure_kpi_excel()
    backup_kpi_excel()

    if monthly_df is None:
        monthly_df = read_sheet("monthly_kpis", MONTHLY_KPI_COLUMNS)

    if manufacturing_df is None:
        manufacturing_df = read_sheet("manufacturing_kpis", MANUFACTURING_KPI_COLUMNS)

    monthly_df = monthly_df.copy()
    manufacturing_df = manufacturing_df.copy()

    for column in MONTHLY_KPI_COLUMNS:
        if column not in monthly_df.columns:
            monthly_df[column] = ""

    for column in MANUFACTURING_KPI_COLUMNS:
        if column not in manufacturing_df.columns:
            manufacturing_df[column] = ""

    monthly_df = monthly_df[MONTHLY_KPI_COLUMNS].copy()
    manufacturing_df = manufacturing_df[MANUFACTURING_KPI_COLUMNS].copy()

    for column in MONTHLY_KPI_COLUMNS:
        monthly_df[column] = monthly_df[column].astype(str)

    for column in MANUFACTURING_KPI_COLUMNS:
        manufacturing_df[column] = manufacturing_df[column].astype(str)

    excel_path = get_kpi_excel_path()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        monthly_df.to_excel(writer, sheet_name="monthly_kpis", index=False)
        manufacturing_df.to_excel(writer, sheet_name="manufacturing_kpis", index=False)


def normalize_number_text(value):
    if value is None:
        return ""

    return str(value).replace(",", "").replace("%", "").strip()


def to_number(value):
    text = normalize_number_text(value)

    if text == "":
        return 0

    try:
        return float(text)
    except Exception:
        return 0


def format_amount(value):
    text = normalize_number_text(value)

    if text == "":
        return ""

    try:
        number = float(text)

        if number.is_integer():
            return f"{int(number):,}"

        return f"{number:,.1f}".rstrip("0").rstrip(".")
    except Exception:
        return value


def format_percent(value):
    text = normalize_number_text(value)

    if text == "":
        return ""

    try:
        number = float(text)
        return f"{number:,.1f}%"
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
        "new_orders",
        "order_backlog",
    ]

    for field in amount_fields:
        row[f"{field}_display"] = format_amount(row.get(field, ""))

    row["gross_profit_rate_display"] = format_percent(row.get("gross_profit_rate"))
    row["operating_profit_rate_display"] = format_percent(row.get("operating_profit_rate"))

    return row


def enrich_manufacturing_kpi(row):
    amount_fields = [
        "production_volume",
        "defect_count",
        "downtime_hours",
        "accident_count",
        "near_miss_count",
        "quality_claim_count",
        "energy_usage",
    ]

    for field in amount_fields:
        row[f"{field}_display"] = format_amount(row.get(field, ""))

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


def generate_next_monthly_kpi_id(year_month, df):
    cleaned_year_month = str(year_month).replace("-", "").replace("/", "").strip()

    if cleaned_year_month:
        base_id = f"KPI-{cleaned_year_month}"
    else:
        base_id = "KPI-NEW"

    existing_ids = set(df["id"].astype(str).tolist()) if "id" in df.columns else set()

    if base_id not in existing_ids:
        return base_id

    number = 2

    while f"{base_id}-{number}" in existing_ids:
        number += 1

    return f"{base_id}-{number}"


def generate_next_manufacturing_kpi_id(year_month, department, df):
    cleaned_year_month = str(year_month).replace("-", "").replace("/", "").strip()
    department_text = str(department).strip()

    if cleaned_year_month:
        base_id = f"MFGKPI-{cleaned_year_month}"
    else:
        base_id = "MFGKPI-NEW"

    if department_text:
        department_code = department_text.replace(" ", "").replace("　", "")
        base_id = f"{base_id}-{department_code}"

    existing_ids = set(df["id"].astype(str).tolist()) if "id" in df.columns else set()

    if base_id not in existing_ids:
        return base_id

    number = 2

    while f"{base_id}-{number}" in existing_ids:
        number += 1

    return f"{base_id}-{number}"


def create_monthly_kpi(
    year_month,
    sales_amount,
    gross_profit,
    operating_profit,
    cash_balance,
    accounts_receivable,
    accounts_payable,
    expense_total,
    new_orders,
    order_backlog,
    comment,
):
    monthly_df = read_sheet("monthly_kpis", MONTHLY_KPI_COLUMNS)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    kpi_id = generate_next_monthly_kpi_id(year_month, monthly_df)

    new_row = {
        "id": kpi_id,
        "year_month": year_month,
        "sales_amount": normalize_number_text(sales_amount),
        "gross_profit": normalize_number_text(gross_profit),
        "operating_profit": normalize_number_text(operating_profit),
        "cash_balance": normalize_number_text(cash_balance),
        "accounts_receivable": normalize_number_text(accounts_receivable),
        "accounts_payable": normalize_number_text(accounts_payable),
        "expense_total": normalize_number_text(expense_total),
        "new_orders": normalize_number_text(new_orders),
        "order_backlog": normalize_number_text(order_backlog),
        "comment": comment,
        "created_at": now_text,
        "updated_at": now_text,
    }

    monthly_df = pd.concat(
        [monthly_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_kpi_excel(monthly_df=monthly_df)

    return kpi_id


def create_manufacturing_kpi(
    year_month,
    department,
    production_volume,
    defect_count,
    defect_rate,
    yield_rate,
    on_time_delivery_rate,
    equipment_availability,
    downtime_hours,
    accident_count,
    near_miss_count,
    quality_claim_count,
    energy_usage,
    comment,
):
    manufacturing_df = read_sheet("manufacturing_kpis", MANUFACTURING_KPI_COLUMNS)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    kpi_id = generate_next_manufacturing_kpi_id(year_month, department, manufacturing_df)

    new_row = {
        "id": kpi_id,
        "year_month": year_month,
        "department": department,
        "production_volume": normalize_number_text(production_volume),
        "defect_count": normalize_number_text(defect_count),
        "defect_rate": normalize_number_text(defect_rate),
        "yield_rate": normalize_number_text(yield_rate),
        "on_time_delivery_rate": normalize_number_text(on_time_delivery_rate),
        "equipment_availability": normalize_number_text(equipment_availability),
        "downtime_hours": normalize_number_text(downtime_hours),
        "accident_count": normalize_number_text(accident_count),
        "near_miss_count": normalize_number_text(near_miss_count),
        "quality_claim_count": normalize_number_text(quality_claim_count),
        "energy_usage": normalize_number_text(energy_usage),
        "comment": comment,
        "created_at": now_text,
        "updated_at": now_text,
    }

    manufacturing_df = pd.concat(
        [manufacturing_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_kpi_excel(manufacturing_df=manufacturing_df)

    return kpi_id


def update_monthly_kpi(
    kpi_id,
    year_month,
    sales_amount,
    gross_profit,
    operating_profit,
    cash_balance,
    accounts_receivable,
    accounts_payable,
    expense_total,
    new_orders,
    order_backlog,
    comment,
):
    monthly_df = read_sheet("monthly_kpis", MONTHLY_KPI_COLUMNS)
    matched = monthly_df["id"].astype(str) == str(kpi_id)

    if not matched.any():
        return None

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    monthly_df.loc[matched, "year_month"] = year_month
    monthly_df.loc[matched, "sales_amount"] = normalize_number_text(sales_amount)
    monthly_df.loc[matched, "gross_profit"] = normalize_number_text(gross_profit)
    monthly_df.loc[matched, "operating_profit"] = normalize_number_text(operating_profit)
    monthly_df.loc[matched, "cash_balance"] = normalize_number_text(cash_balance)
    monthly_df.loc[matched, "accounts_receivable"] = normalize_number_text(accounts_receivable)
    monthly_df.loc[matched, "accounts_payable"] = normalize_number_text(accounts_payable)
    monthly_df.loc[matched, "expense_total"] = normalize_number_text(expense_total)
    monthly_df.loc[matched, "new_orders"] = normalize_number_text(new_orders)
    monthly_df.loc[matched, "order_backlog"] = normalize_number_text(order_backlog)
    monthly_df.loc[matched, "comment"] = comment
    monthly_df.loc[matched, "updated_at"] = now_text

    write_kpi_excel(monthly_df=monthly_df)

    return kpi_id


def update_manufacturing_kpi(
    kpi_id,
    year_month,
    department,
    production_volume,
    defect_count,
    defect_rate,
    yield_rate,
    on_time_delivery_rate,
    equipment_availability,
    downtime_hours,
    accident_count,
    near_miss_count,
    quality_claim_count,
    energy_usage,
    comment,
):
    manufacturing_df = read_sheet("manufacturing_kpis", MANUFACTURING_KPI_COLUMNS)
    matched = manufacturing_df["id"].astype(str) == str(kpi_id)

    if not matched.any():
        return None

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    manufacturing_df.loc[matched, "year_month"] = year_month
    manufacturing_df.loc[matched, "department"] = department
    manufacturing_df.loc[matched, "production_volume"] = normalize_number_text(production_volume)
    manufacturing_df.loc[matched, "defect_count"] = normalize_number_text(defect_count)
    manufacturing_df.loc[matched, "defect_rate"] = normalize_number_text(defect_rate)
    manufacturing_df.loc[matched, "yield_rate"] = normalize_number_text(yield_rate)
    manufacturing_df.loc[matched, "on_time_delivery_rate"] = normalize_number_text(on_time_delivery_rate)
    manufacturing_df.loc[matched, "equipment_availability"] = normalize_number_text(equipment_availability)
    manufacturing_df.loc[matched, "downtime_hours"] = normalize_number_text(downtime_hours)
    manufacturing_df.loc[matched, "accident_count"] = normalize_number_text(accident_count)
    manufacturing_df.loc[matched, "near_miss_count"] = normalize_number_text(near_miss_count)
    manufacturing_df.loc[matched, "quality_claim_count"] = normalize_number_text(quality_claim_count)
    manufacturing_df.loc[matched, "energy_usage"] = normalize_number_text(energy_usage)
    manufacturing_df.loc[matched, "comment"] = comment
    manufacturing_df.loc[matched, "updated_at"] = now_text

    write_kpi_excel(manufacturing_df=manufacturing_df)

    return kpi_id


def format_signed_number(value, suffix=""):
    try:
        number = float(value)
    except Exception:
        return ""

    if number > 0:
        sign = "+"
    elif number < 0:
        sign = ""
    else:
        sign = "±"

    if abs(number) >= 1000:
        text = f"{number:,.0f}"
    else:
        text = f"{number:.1f}".rstrip("0").rstrip(".")

    return f"{sign}{text}{suffix}"


def format_signed_amount(value):
    try:
        number = float(value)
    except Exception:
        return ""

    if number > 0:
        sign = "+"
    elif number < 0:
        sign = ""
    else:
        sign = "±"

    return f"{sign}{number:,.0f}"


def format_signed_percent_point(value):
    return format_signed_number(value, "pt")


def calculate_diff(current_value, previous_value):
    return to_number(current_value) - to_number(previous_value)


def calculate_change_rate(current_value, previous_value):
    current_number = to_number(current_value)
    previous_number = to_number(previous_value)

    if previous_number == 0:
        return ""

    return round((current_number - previous_number) / previous_number * 100, 1)


def get_trend_badge_for_higher_better(diff):
    if diff > 0:
        return "良好"

    if diff < 0:
        return "注意"

    return "横ばい"


def get_trend_badge_for_lower_better(diff):
    if diff < 0:
        return "良好"

    if diff > 0:
        return "要確認"

    return "横ばい"


def get_trend_badge_for_amount_cost(diff):
    if diff < 0:
        return "良好"

    if diff > 0:
        return "注意"

    return "横ばい"


def find_previous_record(records, current_record, same_department=False):
    current_id = str(current_record.get("id", ""))

    if same_department:
        current_department = str(current_record.get("department", ""))
        candidates = [
            row for row in records
            if str(row.get("id", "")) != current_id
            and str(row.get("department", "")) == current_department
            and str(row.get("year_month", "")) < str(current_record.get("year_month", ""))
        ]
    else:
        candidates = [
            row for row in records
            if str(row.get("id", "")) != current_id
            and str(row.get("year_month", "")) < str(current_record.get("year_month", ""))
        ]

    candidates = sorted(
        candidates,
        key=lambda x: str(x.get("year_month", "")),
        reverse=True,
    )

    if candidates:
        return candidates[0]

    return None


def enrich_monthly_kpi_comparison(row, previous_row):
    if not previous_row:
        row["previous_year_month"] = ""
        row["comparison_comment"] = "比較できる前月データがありません。"
        return row

    row["previous_year_month"] = previous_row.get("year_month", "")

    sales_diff = calculate_diff(row.get("sales_amount"), previous_row.get("sales_amount"))
    sales_rate = calculate_change_rate(row.get("sales_amount"), previous_row.get("sales_amount"))

    gross_profit_rate_diff = calculate_diff(
        row.get("gross_profit_rate"),
        previous_row.get("gross_profit_rate"),
    )

    operating_profit_rate_diff = calculate_diff(
        row.get("operating_profit_rate"),
        previous_row.get("operating_profit_rate"),
    )

    cash_diff = calculate_diff(row.get("cash_balance"), previous_row.get("cash_balance"))
    expense_diff = calculate_diff(row.get("expense_total"), previous_row.get("expense_total"))
    order_backlog_diff = calculate_diff(row.get("order_backlog"), previous_row.get("order_backlog"))

    row["sales_amount_diff_display"] = format_signed_amount(sales_diff)
    row["sales_amount_change_rate_display"] = (
        format_signed_number(sales_rate, "%") if sales_rate != "" else ""
    )
    row["sales_amount_trend"] = get_trend_badge_for_higher_better(sales_diff)

    row["gross_profit_rate_diff_display"] = format_signed_percent_point(gross_profit_rate_diff)
    row["gross_profit_rate_trend"] = get_trend_badge_for_higher_better(gross_profit_rate_diff)

    row["operating_profit_rate_diff_display"] = format_signed_percent_point(operating_profit_rate_diff)
    row["operating_profit_rate_trend"] = get_trend_badge_for_higher_better(operating_profit_rate_diff)

    row["cash_balance_diff_display"] = format_signed_amount(cash_diff)
    row["cash_balance_trend"] = get_trend_badge_for_higher_better(cash_diff)

    row["expense_total_diff_display"] = format_signed_amount(expense_diff)
    row["expense_total_trend"] = get_trend_badge_for_amount_cost(expense_diff)

    row["order_backlog_diff_display"] = format_signed_amount(order_backlog_diff)
    row["order_backlog_trend"] = get_trend_badge_for_higher_better(order_backlog_diff)

    row["comparison_comment"] = (
        f"前月（{row['previous_year_month']}）比："
        f"売上 {row['sales_amount_diff_display']}、"
        f"営業利益率 {row['operating_profit_rate_diff_display']}、"
        f"現預金 {row['cash_balance_diff_display']}。"
    )

    return row


def enrich_manufacturing_kpi_comparison(row, previous_row):
    if not previous_row:
        row["previous_year_month"] = ""
        row["comparison_comment"] = "比較できる前月データがありません。"
        return row

    row["previous_year_month"] = previous_row.get("year_month", "")

    production_diff = calculate_diff(
        row.get("production_volume"),
        previous_row.get("production_volume"),
    )

    defect_rate_diff = calculate_diff(
        row.get("defect_rate"),
        previous_row.get("defect_rate"),
    )

    yield_rate_diff = calculate_diff(
        row.get("yield_rate"),
        previous_row.get("yield_rate"),
    )

    on_time_delivery_rate_diff = calculate_diff(
        row.get("on_time_delivery_rate"),
        previous_row.get("on_time_delivery_rate"),
    )

    equipment_availability_diff = calculate_diff(
        row.get("equipment_availability"),
        previous_row.get("equipment_availability"),
    )

    downtime_diff = calculate_diff(
        row.get("downtime_hours"),
        previous_row.get("downtime_hours"),
    )

    accident_diff = calculate_diff(
        row.get("accident_count"),
        previous_row.get("accident_count"),
    )

    near_miss_diff = calculate_diff(
        row.get("near_miss_count"),
        previous_row.get("near_miss_count"),
    )

    quality_claim_diff = calculate_diff(
        row.get("quality_claim_count"),
        previous_row.get("quality_claim_count"),
    )

    row["production_volume_diff_display"] = format_signed_number(production_diff)
    row["production_volume_trend"] = get_trend_badge_for_higher_better(production_diff)

    row["defect_rate_diff_display"] = format_signed_percent_point(defect_rate_diff)
    row["defect_rate_trend"] = get_trend_badge_for_lower_better(defect_rate_diff)

    row["yield_rate_diff_display"] = format_signed_percent_point(yield_rate_diff)
    row["yield_rate_trend"] = get_trend_badge_for_higher_better(yield_rate_diff)

    row["on_time_delivery_rate_diff_display"] = format_signed_percent_point(on_time_delivery_rate_diff)
    row["on_time_delivery_rate_trend"] = get_trend_badge_for_higher_better(on_time_delivery_rate_diff)

    row["equipment_availability_diff_display"] = format_signed_percent_point(equipment_availability_diff)
    row["equipment_availability_trend"] = get_trend_badge_for_higher_better(equipment_availability_diff)

    row["downtime_hours_diff_display"] = format_signed_number(downtime_diff, "h")
    row["downtime_hours_trend"] = get_trend_badge_for_lower_better(downtime_diff)

    row["accident_count_diff_display"] = format_signed_number(accident_diff, "件")
    row["accident_count_trend"] = get_trend_badge_for_lower_better(accident_diff)

    row["near_miss_count_diff_display"] = format_signed_number(near_miss_diff, "件")
    row["near_miss_count_trend"] = get_trend_badge_for_lower_better(near_miss_diff)

    row["quality_claim_count_diff_display"] = format_signed_number(quality_claim_diff, "件")
    row["quality_claim_count_trend"] = get_trend_badge_for_lower_better(quality_claim_diff)

    row["comparison_comment"] = (
        f"前月（{row['previous_year_month']}）比："
        f"不良率 {row['defect_rate_diff_display']}、"
        f"歩留まり {row['yield_rate_diff_display']}、"
        f"設備稼働率 {row['equipment_availability_diff_display']}。"
    )

    return row


def attach_monthly_kpi_comparisons(records):
    enriched_records = []

    for row in records:
        previous_row = find_previous_record(records, row, same_department=False)
        enriched_records.append(enrich_monthly_kpi_comparison(row, previous_row))

    return enriched_records


def attach_manufacturing_kpi_comparisons(records):
    enriched_records = []

    for row in records:
        previous_row = find_previous_record(records, row, same_department=True)
        enriched_records.append(enrich_manufacturing_kpi_comparison(row, previous_row))

    return enriched_records


def get_latest_monthly_kpi_with_comparison():
    records = attach_monthly_kpi_comparisons(load_monthly_kpis())

    if records:
        return records[0]

    return None


def get_latest_manufacturing_kpi_with_comparison():
    records = attach_manufacturing_kpi_comparisons(load_manufacturing_kpis())

    if records:
        return records[0]

    return None


def find_monthly_kpi_by_id_with_comparison(kpi_id):
    records = attach_monthly_kpi_comparisons(load_monthly_kpis())

    for row in records:
        if str(row.get("id")) == str(kpi_id):
            return row

    return None


def find_manufacturing_kpi_by_id_with_comparison(kpi_id):
    records = attach_manufacturing_kpi_comparisons(load_manufacturing_kpis())

    for row in records:
        if str(row.get("id")) == str(kpi_id):
            return row

    return None