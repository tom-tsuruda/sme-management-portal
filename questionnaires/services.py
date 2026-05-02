from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings


def get_questionnaire_excel_path():
    return Path(settings.BASE_DIR) / "data" / "questionnaire_master.xlsx"


def get_answer_history_excel_path():
    return Path(settings.BASE_DIR) / "data" / "questionnaire_answer_history.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_answer_history_excel():
    excel_path = get_answer_history_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"questionnaire_answer_history_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def load_questions():
    """
    data/questionnaire_master.xlsx の questions シートを読み込み、
    質問票データとして返す。
    """
    excel_path = get_questionnaire_excel_path()

    if not excel_path.exists():
        return []

    df = pd.read_excel(excel_path, sheet_name="questions")
    df = df.fillna("")

    return df.to_dict(orient="records")


def generate_diagnosis_id():
    """
    診断IDを自動生成する。
    例：DIAG-20260502-093015
    """
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"DIAG-{timestamp}"


def save_answer_history(diagnosis_id, answers):
    """
    質問票の回答履歴を questionnaire_answer_history.xlsx に保存する。
    """
    excel_path = get_answer_history_excel_path()

    backup_answer_history_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = []

    for answer in answers:
        rows.append({
            "diagnosis_id": diagnosis_id,
            "answered_at": now_text,
            "question_id": answer.get("id", ""),
            "category": answer.get("category", ""),
            "question_text": answer.get("question_text", ""),
            "answer": answer.get("answer", ""),
            "related_task_id": answer.get("related_task_id", ""),
            "related_document_id": answer.get("related_document_id", ""),
            "generated_task_id": answer.get("generated_task_id", ""),
        })

    new_df = pd.DataFrame(rows)

    if excel_path.exists():
        old_df = pd.read_excel(excel_path, sheet_name="answers")
        old_df = old_df.fillna("")
        df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        df = new_df

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="answers", index=False)

    return True

def load_answer_history():
    """
    questionnaire_answer_history.xlsx の answers シートを読み込み、
    回答履歴データとして返す。
    """
    excel_path = get_answer_history_excel_path()

    if not excel_path.exists():
        return []

    df = pd.read_excel(excel_path, sheet_name="answers")
    df = df.fillna("")

    return df.to_dict(orient="records")


def load_diagnosis_summaries():
    """
    診断IDごとのサマリーを返す。
    """
    records = load_answer_history()

    summaries = {}

    for row in records:
        diagnosis_id = row.get("diagnosis_id", "")

        if not diagnosis_id:
            continue

        if diagnosis_id not in summaries:
            summaries[diagnosis_id] = {
                "diagnosis_id": diagnosis_id,
                "answered_at": row.get("answered_at", ""),
                "answer_count": 0,
                "generated_task_count": 0,
                "problem_count": 0,
            }

        summaries[diagnosis_id]["answer_count"] += 1

        if row.get("generated_task_id"):
            summaries[diagnosis_id]["generated_task_count"] += 1

        if row.get("answer") in ["未整備", "不明", "未見直し"]:
            summaries[diagnosis_id]["problem_count"] += 1

    return list(summaries.values())


def load_answers_by_diagnosis_id(diagnosis_id):
    """
    指定した診断IDの回答履歴を返す。
    """
    records = load_answer_history()

    return [
        row for row in records
        if str(row.get("diagnosis_id")) == str(diagnosis_id)
    ]