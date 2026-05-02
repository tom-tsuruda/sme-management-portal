from django.shortcuts import render

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
                        generated_task_id = new_task_id
                        answer_row["generated_task_id"] = new_task_id

                        new_task = find_task_by_id(new_task_id)

                        if new_task:
                            action_tasks.append(new_task)

            answers.append(answer_row)

        save_answer_history(diagnosis_id, answers)

    return render(request, "questionnaires/answer_result.html", {
        "diagnosis_id": diagnosis_id,
        "answers": answers,
        "action_tasks": action_tasks,
    })

def diagnosis_history(request):
    """
    診断履歴一覧を表示する。
    """
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
    """
    診断IDごとの回答結果を表示する。
    """
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