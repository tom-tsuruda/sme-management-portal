from pathlib import Path
from datetime import datetime
import argparse
import shutil
import sys


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SEED_DIR = DATA_DIR / "seeds"
BACKUP_DIR = DATA_DIR / "backups"
MEDIA_DIR = BASE_DIR / "media"

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
]


def ensure_directories():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SEED_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    completed_document_dir = MEDIA_DIR / "documents" / "completed"
    completed_document_dir.mkdir(parents=True, exist_ok=True)


def timestamp_text():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def backup_existing_file(target_path):
    if not target_path.exists():
        return None

    backup_path = BACKUP_DIR / f"{target_path.stem}_{timestamp_text()}{target_path.suffix}"
    shutil.copy2(target_path, backup_path)
    return backup_path


def initialize_file(seed_name, target_name, label, force=False):
    seed_path = SEED_DIR / seed_name
    target_path = DATA_DIR / target_name

    if not seed_path.exists():
        print(f"[NG] {label}: 初期原本が見つかりません: {seed_path}")
        return False

    if target_path.exists() and not force:
        print(f"[SKIP] {label}: 既に運用データがあります: {target_path}")
        return True

    if target_path.exists() and force:
        backup_path = backup_existing_file(target_path)
        print(f"[BACKUP] {label}: 既存ファイルをバックアップしました: {backup_path}")

    shutil.copy2(seed_path, target_path)
    print(f"[OK] {label}: 初期原本から運用データを作成しました: {target_path}")
    return True


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

    print("=== 初期データセットアップ開始 ===")
    print(f"BASE_DIR: {BASE_DIR}")
    print(f"DATA_DIR: {DATA_DIR}")
    print(f"SEED_DIR: {SEED_DIR}")

    success_count = 0

    for item in DEFAULT_FILES:
        result = initialize_file(
            seed_name=item["seed"],
            target_name=item["target"],
            label=item["label"],
            force=args.force,
        )

        if result:
            success_count += 1

    print("=== 初期データセットアップ完了 ===")
    print(f"{success_count}/{len(DEFAULT_FILES)} 件を処理しました。")

    if success_count != len(DEFAULT_FILES):
        sys.exit(1)


if __name__ == "__main__":
    main()