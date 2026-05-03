from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd
from django.conf import settings

from tasks.services import add_task


QUESTION_COLUMNS = [
    "id",
    "category",
    "question_text",
    "answer_type",
    "option_a",
    "option_b",
    "option_c",
    "related_task_id",
    "related_document_id",
]

ANSWER_COLUMNS = [
    "id",
    "diagnosis_id",
    "question_id",
    "category",
    "question_text",
    "answer",
    "related_task_id",
    "related_document_id",
    "generated_task_id",
    "comment",
    "created_at",
]


def get_questionnaire_excel_path():
    return Path(settings.BASE_DIR) / "data" / "questionnaire_data.xlsx"


def get_backup_dir():
    backup_dir = Path(settings.BASE_DIR) / "data" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def backup_questionnaire_excel():
    excel_path = get_questionnaire_excel_path()

    if not excel_path.exists():
        return None

    backup_dir = get_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"questionnaire_data_{timestamp}.xlsx"

    shutil.copy2(excel_path, backup_path)

    return backup_path


def ensure_questionnaire_excel():
    excel_path = get_questionnaire_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if excel_path.exists():
        return

    questions = [
        {
            "id": "Q-001",
            "category": "文書管理",
            "question_text": "主要な社内規程や業務手順書は整備されていますか？",
            "risk_level": "高",
            "recommended_task_name": "主要な社内規程・業務手順書を整備する",
            "recommended_task_category": "文書整備",
            "recommended_owner": "総務部",
            "recommended_due_date": "",
            "related_document_id": "",
        },
        {
            "id": "Q-002",
            "category": "承認管理",
            "question_text": "稟議・申請・承認のルールは明文化されていますか？",
            "risk_level": "高",
            "recommended_task_name": "稟議・申請・承認ルールを明文化する",
            "recommended_task_category": "申請承認",
            "recommended_owner": "管理部",
            "recommended_due_date": "",
            "related_document_id": "",
        },
        {
            "id": "Q-003",
            "category": "経費管理",
            "question_text": "経費精算の証憑や承認履歴は管理されていますか？",
            "risk_level": "中",
            "recommended_task_name": "経費精算ルールと証憑管理を確認する",
            "recommended_task_category": "経費精算",
            "recommended_owner": "経理部",
            "recommended_due_date": "",
            "related_document_id": "",
        },
        {
            "id": "Q-004",
            "category": "製造管理",
            "question_text": "品質・安全・設備に関する管理項目は定期的に確認されていますか？",
            "risk_level": "高",
            "recommended_task_name": "製造管理項目の確認体制を整備する",
            "recommended_task_category": "製造管理",
            "recommended_owner": "製造部",
            "recommended_due_date": "",
            "related_document_id": "",
        },
        {
            "id": "Q-005",
            "category": "ガバナンス",
            "question_text": "定款・議事録・契約書・決算書などの重要文書は管理されていますか？",
            "risk_level": "高",
            "recommended_task_name": "ガバナンス文書台帳を確認する",
            "recommended_task_category": "ガバナンス",
            "recommended_owner": "総務部",
            "recommended_due_date": "",
            "related_document_id": "",
        },
    ]

    answers_df = pd.DataFrame(columns=ANSWER_COLUMNS)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        pd.DataFrame(questions, columns=QUESTION_COLUMNS).to_excel(
            writer,
            sheet_name="questions",
            index=False,
        )
        answers_df.to_excel(
            writer,
            sheet_name="answers",
            index=False,
        )


def read_sheet(sheet_name, columns):
    ensure_questionnaire_excel()
    excel_path = get_questionnaire_excel_path()

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


def write_questionnaire_data(questions_df, answers_df):
    excel_path = get_questionnaire_excel_path()

    questions_df = questions_df.copy()
    answers_df = answers_df.copy()

    for column in QUESTION_COLUMNS:
        if column not in questions_df.columns:
            questions_df[column] = ""

    for column in ANSWER_COLUMNS:
        if column not in answers_df.columns:
            answers_df[column] = ""

    questions_df = questions_df[QUESTION_COLUMNS].copy()
    answers_df = answers_df[ANSWER_COLUMNS].copy()

    for column in QUESTION_COLUMNS:
        questions_df[column] = questions_df[column].astype(str)

    for column in ANSWER_COLUMNS:
        answers_df[column] = answers_df[column].astype(str)

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        questions_df.to_excel(writer, sheet_name="questions", index=False)
        answers_df.to_excel(writer, sheet_name="answers", index=False)


def load_questions():
    df = read_sheet("questions", QUESTION_COLUMNS)
    return df.to_dict(orient="records")


def load_answers():
    df = read_sheet("answers", ANSWER_COLUMNS)
    return df.to_dict(orient="records")


def find_question_by_id(question_id):
    questions = load_questions()

    for question in questions:
        if str(question.get("id")) == str(question_id):
            return question

    return None


def generate_next_answer_id(df):
    if "id" not in df.columns or df.empty:
        return "ANS-001"

    max_number = 0

    for value in df["id"].dropna():
        text = str(value).strip()

        if text.startswith("ANS-"):
            try:
                number = int(text.replace("ANS-", ""))
                max_number = max(max_number, number)
            except ValueError:
                continue

    return f"ANS-{max_number + 1:03d}"


def generate_diagnosis_id():
    return "DIAG-" + datetime.now().strftime("%Y%m%d%H%M%S")


def save_answers(answer_rows):
    questions_df = read_sheet("questions", QUESTION_COLUMNS)
    answers_df = read_sheet("answers", ANSWER_COLUMNS)

    backup_questionnaire_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    diagnosis_id = generate_diagnosis_id()

    new_rows = []

    for row in answer_rows:
        answer_id = generate_next_answer_id(
            pd.concat([answers_df, pd.DataFrame(new_rows)], ignore_index=True)
        )

        new_rows.append({
            "id": answer_id,
            "diagnosis_id": diagnosis_id,
            "question_id": row.get("question_id", ""),
            "answer": row.get("answer", ""),
            "comment": row.get("comment", ""),
            "created_at": now_text,
        })

    if new_rows:
        answers_df = pd.concat([answers_df, pd.DataFrame(new_rows)], ignore_index=True)

    write_questionnaire_data(questions_df, answers_df)

    return diagnosis_id


def load_diagnosis_summaries():
    answers = load_answers()

    diagnosis_map = {}

    for answer in answers:
        diagnosis_id = answer.get("diagnosis_id", "")

        if not diagnosis_id:
            continue

        if diagnosis_id not in diagnosis_map:
            diagnosis_map[diagnosis_id] = {
                "diagnosis_id": diagnosis_id,
                "answered_at": answer.get("created_at", ""),
                "answer_count": 0,
                "problem_count": 0,
                "generated_task_count": 0,
            }

        diagnosis_map[diagnosis_id]["answer_count"] += 1

        if answer.get("answer") in ["いいえ", "未対応", "不明"]:
            diagnosis_map[diagnosis_id]["problem_count"] += 1

    summaries = list(diagnosis_map.values())

    summaries = sorted(
        summaries,
        key=lambda x: str(x.get("answered_at", "")),
        reverse=True,
    )

    return summaries


def load_answers_by_diagnosis_id(diagnosis_id):
    answers = load_answers()

    return [
        answer for answer in answers
        if str(answer.get("diagnosis_id")) == str(diagnosis_id)
    ]


def generate_tasks_from_diagnosis(diagnosis_id):
    answers = load_answers_by_diagnosis_id(diagnosis_id)
    questions = load_questions()

    question_map = {
        str(question.get("id")): question
        for question in questions
    }

    generated_count = 0

    for answer in answers:
        if answer.get("answer") not in ["いいえ", "未対応", "不明"]:
            continue

        question = question_map.get(str(answer.get("question_id")))

        if not question:
            continue

        task_id = add_task(
            task_name=question.get("recommended_task_name", ""),
            category=question.get("recommended_task_category", ""),
            owner=question.get("recommended_owner", ""),
            due_date=question.get("recommended_due_date", ""),
            status="未着手",
            priority=question.get("risk_level", "中"),
            related_document_id=question.get("related_document_id", ""),
        )

        if task_id:
            generated_count += 1

    return generated_count

def save_answer_history(diagnosis_id, answers):
    questions_df = read_sheet("questions", QUESTION_COLUMNS)
    answers_df = read_sheet("answers", ANSWER_COLUMNS)

    backup_questionnaire_excel()

    now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []

    for answer in answers:
        current_df = pd.concat([answers_df, pd.DataFrame(new_rows)], ignore_index=True)
        answer_id = generate_next_answer_id(current_df)

        new_rows.append({
            "id": answer_id,
            "diagnosis_id": diagnosis_id,
            "question_id": answer.get("id", ""),
            "category": answer.get("category", ""),
            "question_text": answer.get("question_text", ""),
            "answer": answer.get("answer", ""),
            "related_task_id": answer.get("related_task_id", ""),
            "related_document_id": answer.get("related_document_id", ""),
            "generated_task_id": answer.get("generated_task_id", ""),
            "comment": answer.get("comment", ""),
            "created_at": now_text,
        })

    if new_rows:
        answers_df = pd.concat([answers_df, pd.DataFrame(new_rows)], ignore_index=True)

    for column in answers_df.columns:
        answers_df[column] = answers_df[column].astype(str)

    write_questionnaire_data(questions_df, answers_df)

    return True