from django.shortcuts import render

from notifications.services import create_notification

from .services import (
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
                str(q.get("related_task_id", "")),
                str(q.get("related_document_id", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_questions.append(q)

        questions = filtered_questions

    return render(request, "questionnaires/question_list.html", {
        "questions": questions,
        "keyword": keyword,
        "total_count": len(all_questions),
        "display_count": len(questions),
    })


def answer_questions(request):
    questions = load_questions()
    answers = []
    action_tasks = []
    diagnosis_id = generate_diagnosis_id()

    problem_answers = [
        "未整備",
        "不明",
        "未見直し",
    ]

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

            if answer in problem_answers:
                problem_count += 1
                task = None

                if related_task_id:
                    task = find_task_by_id(related_task_id)

                if task:
                    action_tasks.append(task)
                else:
                    new_task_name = f"{q.get('category', '')}：{q.get('question_text', '')}"
                    new_task_id = add_task(
                        task_name=new_task_name,
                        category=q.get("category", ""),
                        owner="未設定",
                        due_date="",
                        status="未着手",
                        priority="高",
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
                            target_user="管理者",
                            category="診断質問票",
                            priority="高",
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

    return render(request, "questionnaires/diagnosis_detail.html", {
        "diagnosis_id": diagnosis_id,
        "answers": answers,
        "action_tasks": action_tasks,
    })