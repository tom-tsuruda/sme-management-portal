from pathlib import Path
from datetime import datetime
import argparse
import os
import shutil
import sys


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SEED_DIR = DATA_DIR / "seeds"
BACKUP_DIR = DATA_DIR / "backups"
MEDIA_DIR = BASE_DIR / "media"


sys.path.append(str(BASE_DIR))
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")


def setup_django():
    import django
    django.setup()


DEFAULT_FILES = [
    {
        "seed": "document_master_default.xlsx",
        "target": "document_master.xlsx",
        "label": "文書マスタ",
    },
    {
        "seed": "task_data_default.xlsx",
        "target": "task_data.xlsx",
        "label": "タスクデータ",
    },
    {
        "seed": "questionnaire_data_default.xlsx",
        "target": "questionnaire_data.xlsx",
        "label": "質問票データ",
    },
    {
        "seed": "manufacturing_master_default.xlsx",
        "target": "manufacturing_master.xlsx",
        "label": "製造管理データ",
    },
    {
        "seed": "kpi_data_default.xlsx",
        "target": "kpi_data.xlsx",
        "label": "経営KPIデータ",
    },
    {
        "seed": "workflow_data_default.xlsx",
        "target": "workflow_data.xlsx",
        "label": "申請承認データ",
    },
    {
        "seed": "expense_data_default.xlsx",
        "target": "expense_data.xlsx",
        "label": "経費精算データ",
    },
    {
        "seed": "organization_master_default.xlsx",
        "target": "organization_master.xlsx",
        "label": "組織マスタ",
    },
    {
        "seed": "notification_data_default.xlsx",
        "target": "notification_data.xlsx",
        "label": "操作履歴データ",
    },
    {
        "seed": "accounting_data_default.xlsx",
        "target": "accounting_data.xlsx",
        "label": "経理・決算データ",
        "allow_generate": True,
    },
]


def ensure_directories():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SEED_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    completed_document_dir = MEDIA_DIR / "documents" / "completed"
    completed_document_dir.mkdir(parents=True, exist_ok=True)


def timestamp_text():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def cleanup_old_backups(file_prefix, keep_count=3):
    backup_files = sorted(
        BACKUP_DIR.glob(f"{file_prefix}_*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

    for old_file in backup_files[keep_count:]:
        try:
            old_file.unlink()
            print(f"[CLEANUP] 古いバックアップを削除しました: {old_file}")
        except Exception as exc:
            print(f"[WARN] 古いバックアップを削除できませんでした: {old_file} / {exc}")


def backup_existing_file(target_path):
    if not target_path.exists():
        return None

    backup_path = BACKUP_DIR / f"{target_path.stem}_{timestamp_text()}{target_path.suffix}"
    shutil.copy2(target_path, backup_path)
    cleanup_old_backups(target_path.stem, keep_count=3)

    return backup_path


def initialize_file(seed_name, target_name, label, force=False, allow_generate=False):
    seed_path = SEED_DIR / seed_name
    target_path = DATA_DIR / target_name

    if target_path.exists() and not force:
        print(f"[SKIP] {label}: 既に運用データがあります: {target_path}")
        return True

    if target_path.exists() and force:
        backup_path = backup_existing_file(target_path)
        print(f"[BACKUP] {label}: 既存ファイルをバックアップしました: {backup_path}")

    if seed_path.exists():
        shutil.copy2(seed_path, target_path)
        print(f"[OK] {label}: 初期原本から運用データを作成しました: {target_path}")
        return True

    if allow_generate and target_name == "accounting_data.xlsx":
        print(f"[INFO] {label}: 初期原本がないため、空の経理・決算Excelを生成します。")

        if target_path.exists() and force:
            target_path.unlink()

        from accounting.services import ensure_accounting_excel
        ensure_accounting_excel()

        if target_path.exists():
            print(f"[OK] {label}: accounting.services から運用データを生成しました: {target_path}")
            return True

        print(f"[NG] {label}: accounting_data.xlsx の生成に失敗しました: {target_path}")
        return False

    print(f"[NG] {label}: 初期原本が見つかりません: {seed_path}")
    return False


def main():
    parser = argparse.ArgumentParser(
        description="初期原本 data/seeds から運用データ data/*.xlsx を作成します。"
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="既存の運用データをバックアップしたうえで、初期原本から上書きします。",
    )

    args = parser.parse_args()

    ensure_directories()
    setup_django()

    print("=== 初期データセットアップ開始 ===")
    print(f"BASE_DIR: {BASE_DIR}")
    print(f"DATA_DIR: {DATA_DIR}")
    print(f"SEED_DIR: {SEED_DIR}")
    print(f"BACKUP_DIR: {BACKUP_DIR}")
    print(f"FORCE: {args.force}")

    success_count = 0

    for item in DEFAULT_FILES:
        result = initialize_file(
            seed_name=item["seed"],
            target_name=item["target"],
            label=item["label"],
            force=args.force,
            allow_generate=item.get("allow_generate", False),
        )

        if result:
            success_count += 1

    print("=== 初期データセットアップ完了 ===")
    print(f"{success_count}/{len(DEFAULT_FILES)} 件を処理しました。")

    if success_count != len(DEFAULT_FILES):
        sys.exit(1)


if __name__ == "__main__":
    main()