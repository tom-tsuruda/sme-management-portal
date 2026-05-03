from django.shortcuts import render

from notifications.services import create_notification

from .services import (
    PROBLEM_ANSWERS,
    generate_diagnosis_id,
    load_answers_by_diagnosis_id,
    load_diagnosis_summaries,
    load_questions,
    save_answer_history,
)

from tasks.services import add_task, find_task_by_id, load_tasks

OPEN_TASK_STATUSES = [
    "未着手",
    "進行中",
    "要対応",
    "確認中",
    "差戻し",
]


def normalize_text(value):
    return str(value or "").strip()


def find_existing_open_task(task_name, category, related_document_id=""):
    """
    診断から同じ改善タスクを何度も作らないため、
    同じタスク名・カテゴリ・関連文書IDの未完了タスクを探す。
    """
    task_name = normalize_text(task_name)
    category = normalize_text(category)
    related_document_id = normalize_text(related_document_id)

    for task in load_tasks():
        existing_task_name = normalize_text(task.get("task_name", ""))
        existing_category = normalize_text(task.get("category", ""))
        existing_related_document_id = normalize_text(task.get("related_document_id", ""))
        existing_status = normalize_text(task.get("status", ""))

        if existing_status not in OPEN_TASK_STATUSES:
            continue

        if existing_task_name != task_name:
            continue

        if existing_category != category:
            continue

        if existing_related_document_id != related_document_id:
            continue

        return task

    return None

def question_list(request):
    all_questions = load_questions()
    questions = all_questions

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered_questions = []

        for q in questions:
            target_text = " ".join([
                str(q.get("id", "")),
                str(q.get("category", "")),
                str(q.get("question_text", "")),
                str(q.get("answer_type", "")),
                str(q.get("option_a", "")),
                str(q.get("option_b", "")),
                str(q.get("option_c", "")),
                str(q.get("option_d", "")),
                str(q.get("option_e", "")),
                str(q.get("related_task_id", "")),
                str(q.get("related_document_id", "")),
                str(q.get("risk_level", "")),
                str(q.get("recommended_task_name", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_questions.append(q)

        questions = filtered_questions

    grouped_questions = {}

    for q in questions:
        category = q.get("category", "") or "未分類"

        if category not in grouped_questions:
            grouped_questions[category] = []

        grouped_questions[category].append(q)

    return render(request, "questionnaires/question_list.html", {
        "questions": questions,
        "grouped_questions": grouped_questions,
        "keyword": keyword,
        "total_count": len(all_questions),
        "display_count": len(questions),
    })


def answer_questions(request):
    questions = load_questions()
    answers = []
    action_tasks = []
    diagnosis_id = generate_diagnosis_id()

    problem_count = 0
    generated_task_count = 0

    if request.method == "POST":
        for q in questions:
            q_id = q.get("id")
            answer = request.POST.get(q_id, "")

            related_task_id = q.get("related_task_id", "")
            related_document_id = q.get("related_document_id", "")

            answer_row = {
                "id": q_id,
                "category": q.get("category", ""),
                "question_text": q.get("question_text", ""),
                "answer": answer,
                "related_task_id": related_task_id,
                "related_document_id": related_document_id,
                "generated_task_id": "",
            }

            if answer in PROBLEM_ANSWERS:
                problem_count += 1

                task = None

                if related_task_id:
                    task = find_task_by_id(related_task_id)

                if task:
                    action_tasks.append(task)
                else:
                    new_task_name = (
                        q.get("recommended_task_name", "").strip()
                        or f"{q.get('category', '')}：{q.get('question_text', '')}"
                    )

                    new_task_category = (
                        q.get("recommended_task_category", "").strip()
                        or q.get("category", "")
                    )

                    new_task_owner = (
                        q.get("recommended_owner", "").strip()
                        or "未設定"
                    )

                    new_task_priority = (
                        q.get("priority", "").strip()
                        or q.get("risk_level", "").strip()
                        or "中"
                    )

                    if new_task_priority not in ["高", "中", "低"]:
                        new_task_priority = "中"

                    existing_task = find_existing_open_task(
                        task_name=new_task_name,
                        category=new_task_category,
                        related_document_id=related_document_id,
                    )

                    if existing_task:
                        existing_task_id = existing_task.get("id", "")
                        answer_row["generated_task_id"] = existing_task_id
                        action_tasks.append(existing_task)

                        create_notification(
                            title="診断結果から既存タスクに紐づけました",
                            message=(
                                f"診断ID：{diagnosis_id}。"
                                f"質問「{q.get('question_text', '')}」への回答が"
                                f"「{answer}」だったため、既存の未完了タスクに紐づけました。"
                                f"タスクID：{existing_task_id}。"
                            ),
                            target_user=new_task_owner or "管理者",
                            category="診断質問票",
                            priority=new_task_priority,
                            related_type="tasks",
                            related_id=existing_task_id,
                        )

                    else:
                        new_task_id = add_task(
                            task_name=new_task_name,
                            category=new_task_category,
                            owner=new_task_owner,
                            due_date="",
                            status="未着手",
                            priority=new_task_priority,
                            related_document_id=related_document_id,
                        )

                        if new_task_id:
                            generated_task_count += 1
                            answer_row["generated_task_id"] = new_task_id

                            new_task = find_task_by_id(new_task_id)

                            if new_task:
                                action_tasks.append(new_task)

                            create_notification(
                                title="診断結果からタスクが作成されました",
                                message=(
                                    f"診断ID：{diagnosis_id}。"
                                    f"質問「{q.get('question_text', '')}」への回答が"
                                    f"「{answer}」だったため、改善タスクを作成しました。"
                                    f"タスクID：{new_task_id}。"
                                ),
                                target_user=new_task_owner or "管理者",
                                category="診断質問票",
                                priority=new_task_priority,
                                related_type="tasks",
                                related_id=new_task_id,
                            )

            answers.append(answer_row)

        save_answer_history(diagnosis_id, answers)

        if problem_count > 0:
            create_notification(
                title="診断結果に要対応項目があります",
                message=(
                    f"診断ID：{diagnosis_id} の回答で、"
                    f"要対応項目が {problem_count} 件見つかりました。"
                    f"自動生成タスク数：{generated_task_count} 件。"
                    f"診断履歴画面で内容を確認してください。"
                ),
                target_user="管理者",
                category="診断質問票",
                priority="高",
                related_type="questionnaires",
                related_id=diagnosis_id,
            )
        else:
            create_notification(
                title="診断回答が完了しました",
                message=(
                    f"診断ID：{diagnosis_id} の回答が完了しました。"
                    f"要対応項目はありませんでした。"
                ),
                target_user="管理者",
                category="診断質問票",
                priority="低",
                related_type="questionnaires",
                related_id=diagnosis_id,
            )

    grouped_answers = {}

    answered_count = 0
    not_applicable_count = 0
    unanswered_count = 0
    summary_problem_count = 0
    summary_generated_task_count = 0

    for answer_row in answers:
        category = answer_row.get("category", "") or "未分類"

        if category not in grouped_answers:
            grouped_answers[category] = []

        grouped_answers[category].append(answer_row)

        answer_value = answer_row.get("answer", "")

        if answer_value:
            answered_count += 1

        if answer_value in PROBLEM_ANSWERS:
            summary_problem_count += 1

        if answer_value == "対象外":
            not_applicable_count += 1

        if not answer_value:
            unanswered_count += 1

        if answer_row.get("generated_task_id"):
            summary_generated_task_count += 1

    return render(request, "questionnaires/answer_result.html", {
        "diagnosis_id": diagnosis_id,
        "answers": answers,
        "grouped_answers": grouped_answers,
        "action_tasks": action_tasks,
        "answered_count": answered_count,
        "problem_count": summary_problem_count,
        "generated_task_count": summary_generated_task_count,
        "not_applicable_count": not_applicable_count,
        "unanswered_count": unanswered_count,
    })

def diagnosis_history(request):
    summaries = load_diagnosis_summaries()

    summaries = sorted(
        summaries,
        key=lambda x: x.get("answered_at", ""),
        reverse=True,
    )

    return render(request, "questionnaires/diagnosis_history.html", {
        "summaries": summaries,
    })


def diagnosis_detail(request, diagnosis_id):
    answers = load_answers_by_diagnosis_id(diagnosis_id)

    action_tasks = []

    for answer in answers:
        task_id = answer.get("generated_task_id") or answer.get("related_task_id")

        if task_id:
            task = find_task_by_id(task_id)

            if task:
                action_tasks.append(task)

    grouped_answers = {}

    for answer in answers:
        category = answer.get("category", "") or "未分類"

        if category not in grouped_answers:
            grouped_answers[category] = []

        grouped_answers[category].append(answer)

    return render(request, "questionnaires/diagnosis_detail.html", {
        "diagnosis_id": diagnosis_id,
        "answers": answers,
        "grouped_answers": grouped_answers,
        "action_tasks": action_tasks,
    })