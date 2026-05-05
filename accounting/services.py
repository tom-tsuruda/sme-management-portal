from pathlib import Path
from datetime import datetime
import re
import shutil

import pandas as pd
from django.conf import settings

from core.formatters import (
    format_amount,
    format_japanese_date,
    normalize_amount,
)

try:
    from expenses.services import load_expenses
except Exception:
    load_expenses = None


SALES_COLUMNS = [
    "id",
    "sale_date",
    "customer",
    "title",
    "amount",
    "tax_amount",
    "total_amount",
    "due_date",
    "status",
    "memo",
    "created_at",
    "updated_at",
]

RECEIVABLE_COLUMNS = [
    "id",
    "source_type",
    "source_id",
    "invoice_date",
    "customer",
    "title",
    "amount",
    "due_date",
    "status",
    "collected_date",
    "memo",
    "created_at",
    "updated_at",
]

RECEIVABLE_COLUMNS = [
    "id",
    "source_type",
    "source_id",
    "invoice_date",
    "customer",
    "title",
    "amount",
    "due_date",
    "status",
    "collected_date",
    "memo",
    "created_at",
    "updated_at",
]

PAYABLE_COLUMNS = [
    "id",
    "purchase_date",
    "vendor",
    "title",
    "amount",
    "due_date",
    "status",
    "paid_date",
    "memo",
    "created_at",
    "updated_at",
]

BALANCE_SHEET_COLUMNS = [
    "id",
    "target_month",
    "cash_balance",
    "inventory_amount",
    "fixed_assets",
    "loan_balance",
    "memo",
    "created_at",
    "updated_at",
]


def get_accounting_excel_path():
    return Path(settings.BASE_DIR) / "data" / "accounting_data.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def cleanup_old_backups(file_prefix, keep_count=3):
    backup_dir = get_backup_dir()
    backup_files = sorted(
        backup_dir.glob(f"{file_prefix}_*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

    for old_file in backup_files[keep_count:]:
        try:
            old_file.unlink()
        except Exception:
            pass


def backup_accounting_excel():
    excel_path = get_accounting_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"accounting_data_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)
    cleanup_old_backups("accounting_data", keep_count=3)

    return backup_path


def ensure_accounting_excel():
    excel_path = get_accounting_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(columns=SALES_COLUMNS).to_excel(
            writer,
            sheet_name="sales",
            index=False,
        )
        pd.DataFrame(columns=RECEIVABLE_COLUMNS).to_excel(
            writer,
            sheet_name="receivables",
            index=False,
        )
        pd.DataFrame(columns=PAYABLE_COLUMNS).to_excel(
            writer,
            sheet_name="payables",
            index=False,
        )
        pd.DataFrame(columns=BALANCE_SHEET_COLUMNS).to_excel(
            writer,
            sheet_name="balance_sheet",
            index=False,
        )


def read_sheet(sheet_name, columns):
    ensure_accounting_excel()
    excel_path = get_accounting_excel_path()

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


def write_accounting_excel(
    sales_df=None,
    receivables_df=None,
    payables_df=None,
    balance_sheet_df=None,
):
    ensure_accounting_excel()
    backup_accounting_excel()

    if sales_df is None:
        sales_df = read_sheet("sales", SALES_COLUMNS)

    if receivables_df is None:
        receivables_df = read_sheet("receivables", RECEIVABLE_COLUMNS)

    if payables_df is None:
        payables_df = read_sheet("payables", PAYABLE_COLUMNS)

    if balance_sheet_df is None:
        balance_sheet_df = read_sheet("balance_sheet", BALANCE_SHEET_COLUMNS)

    sales_df = normalize_dataframe_columns(sales_df, SALES_COLUMNS)
    receivables_df = normalize_dataframe_columns(receivables_df, RECEIVABLE_COLUMNS)
    payables_df = normalize_dataframe_columns(payables_df, PAYABLE_COLUMNS)
    balance_sheet_df = normalize_dataframe_columns(balance_sheet_df, BALANCE_SHEET_COLUMNS)

    excel_path = get_accounting_excel_path()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        sales_df.to_excel(writer, sheet_name="sales", index=False)
        receivables_df.to_excel(writer, sheet_name="receivables", index=False)
        payables_df.to_excel(writer, sheet_name="payables", index=False)
        balance_sheet_df.to_excel(writer, sheet_name="balance_sheet", index=False)


def normalize_dataframe_columns(df, columns):
    df = df.copy()

    for column in columns:
        if column not in df.columns:
            df[column] = ""

    df = df[columns].copy()

    for column in columns:
        df[column] = df[column].astype(str)

    return df


def to_number(value):
    text = normalize_amount(value)

    if text == "":
        return 0

    try:
        return int(float(text))
    except Exception:
        return 0


def generate_next_id(df, prefix):
    if "id" not in df.columns or df.empty:
        return f"{prefix}-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith(f"{prefix}-"):
            try:
                number = int(text.replace(f"{prefix}-", ""))
                max_number = max(max_number, number)
            except Exception:
                continue

    return f"{prefix}-{max_number + 1:03d}"


def normalize_record_display(row, amount_fields=None, date_fields=None):
    amount_fields = amount_fields or []
    date_fields = date_fields or []

    for field in amount_fields:
        row[f"{field}_display"] = format_amount(row.get(field, ""))

    for field in date_fields:
        row[f"{field}_display"] = format_japanese_date(row.get(field, ""))

    return row


def extract_year_month(date_text):
    text = str(date_text or "").strip()

    if not text:
        return ""

    text = text.replace("年", "-")
    text = text.replace("月", "-")
    text = text.replace("日", "")
    text = text.replace("/", "-")
    text = text.strip()

    match = re.search(r"(\d{4})-(\d{1,2})", text)
    if match:
        year = match.group(1)
        month = int(match.group(2))
        return f"{year}-{month:02d}"

    return ""


def filter_by_month(records, date_field, target_month):
    if not target_month:
        return records

    return [
        row for row in records
        if extract_year_month(row.get(date_field, "")) == target_month
    ]


def load_sales():
    df = read_sheet("sales", SALES_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        normalize_record_display(
            row,
            amount_fields=["amount", "tax_amount", "total_amount"],
            date_fields=["sale_date", "due_date"],
        )

    return sorted(records, key=lambda x: str(x.get("sale_date", "")), reverse=True)


def load_receivables():
    df = read_sheet("receivables", RECEIVABLE_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        normalize_record_display(
            row,
            amount_fields=["amount"],
            date_fields=["invoice_date", "due_date", "collected_date"],
        )

    return sorted(records, key=lambda x: str(x.get("invoice_date", "")), reverse=True)


def load_payables():
    df = read_sheet("payables", PAYABLE_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        normalize_record_display(
            row,
            amount_fields=["amount"],
            date_fields=["purchase_date", "due_date", "paid_date"],
        )

    return sorted(records, key=lambda x: str(x.get("purchase_date", "")), reverse=True)


def load_balance_sheets():
    df = read_sheet("balance_sheet", BALANCE_SHEET_COLUMNS)
    records = df.to_dict(orient="records")

    for row in records:
        normalize_record_display(
            row,
            amount_fields=[
                "cash_balance",
                "inventory_amount",
                "fixed_assets",
                "loan_balance",
            ],
            date_fields=[],
        )

    return sorted(records, key=lambda x: str(x.get("target_month", "")), reverse=True)


def find_balance_sheet_by_month(target_month):
    for row in load_balance_sheets():
        if str(row.get("target_month", "")) == str(target_month):
            return row

    return None

def receivable_status_from_sale_status(sale_status):
    """
    売上ステータスから売掛金ステータスを決める。
    """
    if sale_status == "入金済":
        return "回収済"

    if sale_status in ["未請求", "請求済"]:
        return "未回収"

    return "未回収"


def sync_receivable_from_sale(sale_id):
    """
    売上データをもとに、対応する売掛金を自動作成・自動更新する。
    売上の反対側として売掛金を必ずセットで持つ。
    """
    sales_df = read_sheet("sales", SALES_COLUMNS)
    receivables_df = read_sheet("receivables", RECEIVABLE_COLUMNS)

    sale_indexes = sales_df.index[
        sales_df["id"].astype(str) == str(sale_id)
    ].tolist()

    if not sale_indexes:
        return None

    sale = sales_df.loc[sale_indexes[0]].to_dict()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    receivable_indexes = receivables_df.index[
        (receivables_df["source_type"].astype(str) == "sale")
        & (receivables_df["source_id"].astype(str) == str(sale_id))
    ].tolist()

    receivable_status = receivable_status_from_sale_status(
        str(sale.get("status", ""))
    )

    collected_date = ""
    if receivable_status == "回収済":
        collected_date = sale.get("sale_date", "")

    if receivable_indexes:
        index = receivable_indexes[0]
        receivable_id = receivables_df.loc[index, "id"]
        created_at = receivables_df.loc[index, "created_at"]
    else:
        index = None
        receivable_id = generate_next_id(receivables_df, "AR")
        created_at = now_text

    receivable_row = {
        "id": receivable_id,
        "source_type": "sale",
        "source_id": sale_id,
        "invoice_date": sale.get("sale_date", ""),
        "customer": sale.get("customer", ""),
        "title": sale.get("title", ""),
        "amount": sale.get("total_amount", "0"),
        "due_date": sale.get("due_date", ""),
        "status": receivable_status,
        "collected_date": collected_date,
        "memo": f"売上 {sale_id} から自動作成",
        "created_at": created_at,
        "updated_at": now_text,
    }

    if index is None:
        receivables_df = pd.concat(
            [receivables_df, pd.DataFrame([receivable_row])],
            ignore_index=True,
        )
    else:
        for key, value in receivable_row.items():
            receivables_df.loc[index, key] = value

    write_accounting_excel(
        sales_df=sales_df,
        receivables_df=receivables_df,
    )

    return receivable_id

def create_sale(
    sale_date,
    customer,
    title,
    amount,
    tax_amount,
    due_date,
    status,
    memo,
):
    sales_df = read_sheet("sales", SALES_COLUMNS)
    sale_id = generate_next_id(sales_df, "SALE")
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    amount_value = to_number(amount)
    tax_value = to_number(tax_amount)
    total_value = amount_value + tax_value

    new_row = {
        "id": sale_id,
        "sale_date": sale_date,
        "customer": customer,
        "title": title,
        "amount": str(amount_value),
        "tax_amount": str(tax_value),
        "total_amount": str(total_value),
        "due_date": due_date,
        "status": status,
        "memo": memo,
        "created_at": now_text,
        "updated_at": now_text,
    }

    sales_df = pd.concat(
        [sales_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )

    write_accounting_excel(sales_df=sales_df)

    # 売上の反対側として売掛金を自動作成
    sync_receivable_from_sale(sale_id)

    return sale_id

def update_sale(
    sale_id,
    sale_date,
    customer,
    title,
    amount,
    tax_amount,
    due_date,
    status,
    memo,
):
    sales_df = read_sheet("sales", SALES_COLUMNS)

    target_index = sales_df.index[
        sales_df["id"].astype(str) == str(sale_id)
    ].tolist()

    if not target_index:
        return False

    index = target_index[0]
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    amount_value = to_number(amount)
    tax_value = to_number(tax_amount)
    total_value = amount_value + tax_value

    sales_df.loc[index, "sale_date"] = sale_date
    sales_df.loc[index, "customer"] = customer
    sales_df.loc[index, "title"] = title
    sales_df.loc[index, "amount"] = str(amount_value)
    sales_df.loc[index, "tax_amount"] = str(tax_value)
    sales_df.loc[index, "total_amount"] = str(total_value)
    sales_df.loc[index, "due_date"] = due_date
    sales_df.loc[index, "status"] = status
    sales_df.loc[index, "memo"] = memo
    sales_df.loc[index, "updated_at"] = now_text

    write_accounting_excel(sales_df=sales_df)

    # 売上を編集したら、対応する売掛金も自動更新
    sync_receivable_from_sale(sale_id)

    return True

def create_receivable(
    invoice_date,
    customer,
    title,
    amount,
    due_date,
    status,
    collected_date,
    memo,
):
    receivables_df = read_sheet("receivables", RECEIVABLE_COLUMNS)
    receivable_id = generate_next_id(receivables_df, "AR")
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = {
        "id": receivable_id,
        "source_type": "manual",
        "source_id": "",
        "invoice_date": invoice_date,
        "customer": customer,
        "title": title,
        "amount": str(to_number(amount)),
        "due_date": due_date,
        "status": status,
        "collected_date": collected_date,
        "memo": memo or "手入力売掛金",
        "created_at": now_text,
        "updated_at": now_text,
    }

    receivables_df = pd.concat(
        [receivables_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )
    write_accounting_excel(receivables_df=receivables_df)

    return receivable_id


def create_payable(
    purchase_date,
    vendor,
    title,
    amount,
    due_date,
    status,
    paid_date,
    memo,
):
    payables_df = read_sheet("payables", PAYABLE_COLUMNS)
    payable_id = generate_next_id(payables_df, "AP")
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = {
        "id": payable_id,
        "purchase_date": purchase_date,
        "vendor": vendor,
        "title": title,
        "amount": str(to_number(amount)),
        "due_date": due_date,
        "status": status,
        "paid_date": paid_date,
        "memo": memo,
        "created_at": now_text,
        "updated_at": now_text,
    }

    payables_df = pd.concat(
        [payables_df, pd.DataFrame([new_row])],
        ignore_index=True,
    )
    write_accounting_excel(payables_df=payables_df)

    return payable_id


def save_balance_sheet(
    target_month,
    cash_balance,
    inventory_amount,
    fixed_assets,
    loan_balance,
    memo,
):
    balance_sheet_df = read_sheet("balance_sheet", BALANCE_SHEET_COLUMNS)
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    target_index = balance_sheet_df.index[
        balance_sheet_df["target_month"].astype(str) == str(target_month)
    ].tolist()

    if target_index:
        index = target_index[0]
        balance_sheet_id = balance_sheet_df.loc[index, "id"]
        created_at = balance_sheet_df.loc[index, "created_at"]
    else:
        index = None
        balance_sheet_id = generate_next_id(balance_sheet_df, "BS")
        created_at = now_text

    row_data = {
        "id": balance_sheet_id,
        "target_month": target_month,
        "cash_balance": str(to_number(cash_balance)),
        "inventory_amount": str(to_number(inventory_amount)),
        "fixed_assets": str(to_number(fixed_assets)),
        "loan_balance": str(to_number(loan_balance)),
        "memo": memo,
        "created_at": created_at,
        "updated_at": now_text,
    }

    if index is None:
        balance_sheet_df = pd.concat(
            [balance_sheet_df, pd.DataFrame([row_data])],
            ignore_index=True,
        )
    else:
        for key, value in row_data.items():
            balance_sheet_df.loc[index, key] = value

    write_accounting_excel(balance_sheet_df=balance_sheet_df)

    return balance_sheet_id


def get_expense_records_for_accounting():
    if load_expenses is None:
        return []

    try:
        return load_expenses()
    except Exception:
        return []


def calculate_accounting_summary(target_month=""):
    sales = load_sales()
    receivables = load_receivables()
    payables = load_payables()
    expenses = get_expense_records_for_accounting()

    available_months = sorted(
        {
            extract_year_month(row.get("sale_date", ""))
            for row in sales
        }
        | {
            extract_year_month(row.get("invoice_date", ""))
            for row in receivables
        }
        | {
            extract_year_month(row.get("purchase_date", ""))
            for row in payables
        }
        | {
            extract_year_month(row.get("expense_date", ""))
            for row in expenses
        }
        | {
            str(row.get("target_month", ""))
            for row in load_balance_sheets()
        },
        reverse=True,
    )

    available_months = [month for month in available_months if month]

    target_month = str(target_month or "").strip()

    if not target_month and available_months:
        target_month = available_months[0]

    monthly_sales = filter_by_month(sales, "sale_date", target_month)
    monthly_receivables = filter_by_month(receivables, "invoice_date", target_month)
    monthly_payables = filter_by_month(payables, "purchase_date", target_month)
    monthly_expenses = filter_by_month(expenses, "expense_date", target_month)

    sales_total = sum(to_number(row.get("total_amount")) for row in monthly_sales)

    approved_expense_statuses = ["承認済", "精算済"]
    expense_total = sum(
        to_number(row.get("amount"))
        for row in monthly_expenses
        if str(row.get("status", "")) in approved_expense_statuses
    )

    receivable_total = sum(
        to_number(row.get("amount"))
        for row in monthly_receivables
    )

    receivable_uncollected = sum(
        to_number(row.get("amount"))
        for row in receivables
        if str(row.get("status", "")) != "回収済"
    )

    payable_total = sum(
        to_number(row.get("amount"))
        for row in monthly_payables
    )

    payable_unpaid = sum(
        to_number(row.get("amount"))
        for row in payables
        if str(row.get("status", "")) != "支払済"
    )

    simple_profit = sales_total - expense_total - payable_total

    balance_sheet = find_balance_sheet_by_month(target_month)

    cash_balance = to_number(balance_sheet.get("cash_balance")) if balance_sheet else 0
    inventory_amount = to_number(balance_sheet.get("inventory_amount")) if balance_sheet else 0
    fixed_assets = to_number(balance_sheet.get("fixed_assets")) if balance_sheet else 0
    loan_balance = to_number(balance_sheet.get("loan_balance")) if balance_sheet else 0

    asset_total = (
        cash_balance
        + receivable_uncollected
        + inventory_amount
        + fixed_assets
    )
    liability_total = payable_unpaid + loan_balance
    net_assets = asset_total - liability_total

    return {
        "target_month": target_month,
        "available_months": available_months,

        "sales_total": sales_total,
        "expense_total": expense_total,
        "receivable_total": receivable_total,
        "receivable_uncollected": receivable_uncollected,
        "payable_total": payable_total,
        "payable_unpaid": payable_unpaid,
        "simple_profit": simple_profit,

        "sales_total_display": format_amount(sales_total),
        "expense_total_display": format_amount(expense_total),
        "receivable_total_display": format_amount(receivable_total),
        "receivable_uncollected_display": format_amount(receivable_uncollected),
        "payable_total_display": format_amount(payable_total),
        "payable_unpaid_display": format_amount(payable_unpaid),
        "simple_profit_display": format_amount(simple_profit),

        "balance_sheet": balance_sheet,
        "cash_balance": cash_balance,
        "inventory_amount": inventory_amount,
        "fixed_assets": fixed_assets,
        "loan_balance": loan_balance,
        "asset_total": asset_total,
        "liability_total": liability_total,
        "net_assets": net_assets,

        "cash_balance_display": format_amount(cash_balance),
        "inventory_amount_display": format_amount(inventory_amount),
        "fixed_assets_display": format_amount(fixed_assets),
        "loan_balance_display": format_amount(loan_balance),
        "asset_total_display": format_amount(asset_total),
        "liability_total_display": format_amount(liability_total),
        "net_assets_display": format_amount(net_assets),

        "monthly_sales": monthly_sales,
        "monthly_receivables": monthly_receivables,
        "monthly_payables": monthly_payables,
        "monthly_expenses": monthly_expenses,
    }

def calculate_cashflow_schedule(target_month=""):
    """
    売掛金・買掛金から簡易資金繰り表を作成する。
    入金予定 = 売掛金の due_date
    支払予定 = 買掛金の due_date
    開始現預金 = 簡易BSの現預金
    """
    receivables = load_receivables()
    payables = load_payables()

    available_months = sorted(
        {
            extract_year_month(row.get("due_date", ""))
            for row in receivables
        }
        | {
            extract_year_month(row.get("due_date", ""))
            for row in payables
        }
        | {
            str(row.get("target_month", ""))
            for row in load_balance_sheets()
        },
        reverse=True,
    )

    available_months = [month for month in available_months if month]

    target_month = str(target_month or "").strip()

    if not target_month and available_months:
        target_month = available_months[0]

    target_receivables = filter_by_month(receivables, "due_date", target_month)
    target_payables = filter_by_month(payables, "due_date", target_month)

    expected_receipts = sum(
        to_number(row.get("amount"))
        for row in target_receivables
        if str(row.get("status", "")) != "回収済"
    )

    expected_payments = sum(
        to_number(row.get("amount"))
        for row in target_payables
        if str(row.get("status", "")) != "支払済"
    )

    balance_sheet = find_balance_sheet_by_month(target_month)
    opening_cash = to_number(balance_sheet.get("cash_balance")) if balance_sheet else 0

    ending_cash = opening_cash + expected_receipts - expected_payments

    daily_map = {}

    for row in target_receivables:
        if str(row.get("status", "")) == "回収済":
            continue

        due_date = str(row.get("due_date", "")).strip() or "日付未設定"

        if due_date not in daily_map:
            daily_map[due_date] = {
                "date": due_date,
                "receipts": 0,
                "payments": 0,
                "balance": 0,
                "items": [],
            }

        amount = to_number(row.get("amount"))
        daily_map[due_date]["receipts"] += amount
        daily_map[due_date]["items"].append({
            "type": "入金",
            "partner": row.get("customer", ""),
            "title": row.get("title", ""),
            "amount": amount,
            "amount_display": format_amount(amount),
        })

    for row in target_payables:
        if str(row.get("status", "")) == "支払済":
            continue

        due_date = str(row.get("due_date", "")).strip() or "日付未設定"

        if due_date not in daily_map:
            daily_map[due_date] = {
                "date": due_date,
                "receipts": 0,
                "payments": 0,
                "balance": 0,
                "items": [],
            }

        amount = to_number(row.get("amount"))
        daily_map[due_date]["payments"] += amount
        daily_map[due_date]["items"].append({
            "type": "支払",
            "partner": row.get("vendor", ""),
            "title": row.get("title", ""),
            "amount": amount,
            "amount_display": format_amount(amount),
        })

    cashflow_rows = sorted(
        daily_map.values(),
        key=lambda row: str(row.get("date", "")),
    )

    running_balance = opening_cash

    for row in cashflow_rows:
        running_balance = running_balance + row["receipts"] - row["payments"]
        row["balance"] = running_balance
        row["receipts_display"] = format_amount(row["receipts"])
        row["payments_display"] = format_amount(row["payments"])
        row["balance_display"] = format_amount(row["balance"])
        row["date_display"] = format_japanese_date(row["date"])

    return {
        "target_month": target_month,
        "available_months": available_months,
        "opening_cash": opening_cash,
        "expected_receipts": expected_receipts,
        "expected_payments": expected_payments,
        "ending_cash": ending_cash,
        "opening_cash_display": format_amount(opening_cash),
        "expected_receipts_display": format_amount(expected_receipts),
        "expected_payments_display": format_amount(expected_payments),
        "ending_cash_display": format_amount(ending_cash),
        "cashflow_rows": cashflow_rows,
    }