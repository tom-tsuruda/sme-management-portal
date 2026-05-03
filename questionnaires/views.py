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

from tasks.services import add_task, find_task_by_id


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

    if request.method == "POST":
        problem_count = 0
        generated_task_count = 0

        for q in questions:
            q_id = q.get("id")
            answer = request.POST.get(q_id, "")

            related_task_id = q.get("related_task_id", "")
            related_document_id = q.get("related_document_id", "")
            generated_task_id = ""

            answer_row = {
                "id": q_id,
                "category": q.get("category", ""),
                "question_text": q.get("question_text", ""),
                "answer": answer,
                "related_task_id": related_task_id,
                "related_document_id": related_document_id,
                "generated_task_id": generated_task_id,
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
                        generated_task_id = new_task_id
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

    return render(request, "questionnaires/answer_result.html", {
        "diagnosis_id": diagnosis_id,
        "answers": answers,
        "action_tasks": action_tasks,
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