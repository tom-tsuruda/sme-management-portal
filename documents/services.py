from pathlib import Path

import pandas as pd
from django.conf import settings


def format_japanese_date(value):
    """
    Excelの日付を 2026年5月31日 の形式に変換する。
    空欄や変換できない値はそのまま返す。
    """
    if value == "" or pd.isna(value):
        return ""

    try:
        date_value = pd.to_datetime(value)
        return f"{date_value.year}年{date_value.month}月{date_value.day}日"
    except Exception:
        return value


def load_documents():
    """
    data/document_master.xlsx の documents シートを読み込み、
    文書台帳データとして返す。
    """
    excel_path = Path(settings.BASE_DIR) / "data" / "document_master.xlsx"

    if not excel_path.exists():
        return []

    df = pd.read_excel(excel_path, sheet_name="documents")
    df = df.fillna("")

    records = df.to_dict(orient="records")

    for doc in records:
        doc["next_review_date"] = format_japanese_date(doc.get("next_review_date", ""))

    return records