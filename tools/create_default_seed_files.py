from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SEED_DIR = DATA_DIR / "seeds"
BACKUP_DIR = DATA_DIR / "backups"


DEFAULT_TARGETS = [
    {
        "source": "document_master.xlsx",
        "seed": "document_master_default.xlsx",
        "label": "文書マスタ",
        "mode": "keep_master",
    },
    {
        "source": "task_data.xlsx",
        "seed": "task_data_default.xlsx",
        "label": "タスクデータ",
        "mode": "empty_all",
    },
    {
        "source": "questionnaire_data.xlsx",
        "seed": "questionnaire_data_default.xlsx",
        "label": "質問票データ",
        "mode": "questionnaire",
    },
    {
        "source": "manufacturing_master.xlsx",
        "seed": "manufacturing_master_default.xlsx",
        "label": "製造管理データ",
        "mode": "manufacturing",
    },
    {
        "source": "kpi_data.xlsx",
        "seed": "kpi_data_default.xlsx",
        "label": "経営KPIデータ",
        "mode": "empty_all",
    },
    {
        "source": "workflow_data.xlsx",
        "seed": "workflow_data_default.xlsx",
        "label": "申請承認データ",
        "mode": "empty_all",
    },
    {
        "source": "expense_data.xlsx",
        "seed": "expense_data_default.xlsx",
        "label": "経費精算データ",
        "mode": "empty_all",
    },
    {
        "source": "organization_master.xlsx",
        "seed": "organization_master_default.xlsx",
        "label": "組織マスタ",
        "mode": "empty_all",
    },
    {
        "source": "notification_data.xlsx",
        "seed": "notification_data_default.xlsx",
        "label": "操作履歴データ",
        "mode": "empty_all",
    },
]


def ensure_directories():
    SEED_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)


def timestamp_text():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def backup_existing_seed(seed_path):
    if not seed_path.exists():
        return None

    backup_path = BACKUP_DIR / f"{seed_path.stem}_{timestamp_text()}{seed_path.suffix}"
    shutil.copy2(seed_path, backup_path)
    return backup_path


def read_excel_sheets(excel_path):
    return pd.read_excel(
        excel_path,
        sheet_name=None,
        dtype=str,
    )


def normalize_df(df):
    df = df.fillna("")

    for column in df.columns:
        df[column] = df[column].astype(str)

    return df


def empty_df_keep_columns(df):
    df = normalize_df(df)
    return df.iloc[0:0].copy()


def build_keep_master_seed(sheets):
    cleaned = {}

    for sheet_name, df in sheets.items():
        cleaned[sheet_name] = normalize_df(df)

    return cleaned


def build_empty_all_seed(sheets):
    cleaned = {}

    for sheet_name, df in sheets.items():
        cleaned[sheet_name] = empty_df_keep_columns(df)

    return cleaned


def build_questionnaire_seed(sheets):
    cleaned = {}

    for sheet_name, df in sheets.items():
        if sheet_name == "questions":
            cleaned[sheet_name] = normalize_df(df)
        else:
            cleaned[sheet_name] = empty_df_keep_columns(df)

    return cleaned


def build_manufacturing_seed(sheets):
    cleaned = {}

    for sheet_name, df in sheets.items():
        if sheet_name == "management_templates":
            cleaned[sheet_name] = normalize_df(df)
        else:
            cleaned[sheet_name] = empty_df_keep_columns(df)

    return cleaned


def build_seed_sheets(sheets, mode):
    if mode == "keep_master":
        return build_keep_master_seed(sheets)

    if mode == "empty_all":
        return build_empty_all_seed(sheets)

    if mode == "questionnaire":
        return build_questionnaire_seed(sheets)

    if mode == "manufacturing":
        return build_manufacturing_seed(sheets)

    raise ValueError(f"Unknown mode: {mode}")


def write_excel_sheets(seed_path, sheets):
    with pd.ExcelWriter(seed_path, engine="openpyxl", mode="w") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def create_seed_file(source_name, seed_name, label, mode, force=False):
    source_path = DATA_DIR / source_name
    seed_path = SEED_DIR / seed_name

    if not source_path.exists():
        print(f"[SKIP] {label}: 運用ファイルが見つかりません: {source_path}")
        return False

    if seed_path.exists() and not force:
        print(f"[SKIP] {label}: 既に初期原本があります: {seed_path}")
        return True

    if seed_path.exists() and force:
        backup_path = backup_existing_seed(seed_path)
        print(f"[BACKUP] {label}: 既存の初期原本をバックアップしました: {backup_path}")

    sheets = read_excel_sheets(source_path)
    cleaned_sheets = build_seed_sheets(sheets, mode)

    write_excel_sheets(seed_path, cleaned_sheets)

    print(f"[OK] {label}: 初期原本を作成しました: {seed_path}")
    return True


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="既存の data/*.xlsx から、配布用の data/seeds/*_default.xlsx を作成します。"
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="既存の default.xlsx をバックアップしたうえで上書きします。",
    )

    args = parser.parse_args()

    ensure_directories()

    print("=== default seed file 作成開始 ===")
    print(f"BASE_DIR: {BASE_DIR}")
    print(f"DATA_DIR: {DATA_DIR}")
    print(f"SEED_DIR: {SEED_DIR}")

    success_count = 0

    for item in DEFAULT_TARGETS:
        result = create_seed_file(
            source_name=item["source"],
            seed_name=item["seed"],
            label=item["label"],
            mode=item["mode"],
            force=args.force,
        )

        if result:
            success_count += 1

    print("=== default seed file 作成完了 ===")
    print(f"{success_count}/{len(DEFAULT_TARGETS)} 件を処理しました。")


if __name__ == "__main__":
    main()