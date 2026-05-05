from pathlib import Path
from datetime import datetime
import shutil


def get_backup_dir(base_dir):
    backup_dir = Path(base_dir) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def cleanup_old_backups(backup_dir, file_prefix, keep_count=3):
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


def backup_excel_file(excel_path, base_dir, file_prefix, keep_count=3):
    excel_path = Path(excel_path)

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir(base_dir)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"{file_prefix}_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    cleanup_old_backups(
        backup_dir=backup_dir,
        file_prefix=file_prefix,
        keep_count=keep_count,
    )

    return backup_path