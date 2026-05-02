import pandas as pd


def normalize_amount(value):
    """
    画面入力された金額を保存用に正規化する。
    例：
    50,000 → 50000
    ５０，０００ → 50000
    空欄 → 空欄
    """
    if value is None:
        return ""

    text = str(value).strip()

    if text == "":
        return ""

    text = text.replace(",", "")
    text = text.replace("，", "")
    text = text.replace("円", "")
    text = text.replace("￥", "")
    text = text.replace("\\", "")
    text = text.strip()

    return text


def format_amount(value):
    """
    金額を3桁区切りで表示する。
    例：
    50000 → 50,000
    50000.0 → 50,000
    空欄や数値化できない値はそのまま返す。
    """
    if value == "" or value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    try:
        text = normalize_amount(value)

        if text == "":
            return ""

        number = int(float(text))
        return f"{number:,}"
    except Exception:
        return value


def format_japanese_datetime(value):
    """
    日時を日本語表示に整形する。
    例：
    2026-05-03 10:30:00 → 2026年5月3日 10:30
    """
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


def format_japanese_date(value):
    """
    日付を日本語表示に整形する。
    例：
    2026-05-03 → 2026年5月3日
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