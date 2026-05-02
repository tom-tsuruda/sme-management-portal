from django.urls import path

from . import views

app_name = "questionnaires"

urlpatterns = [
    path("", views.question_list, name="question_list"),
    path("answer/", views.answer_questions, name="answer_questions"),
]