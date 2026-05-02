from django.urls import path

from . import views

app_name = "questionnaires"

urlpatterns = [
    path("", views.question_list, name="question_list"),
    path("answer/", views.answer_questions, name="answer_questions"),
    path("history/", views.diagnosis_history, name="diagnosis_history"),
    path("history/<str:diagnosis_id>/", views.diagnosis_detail, name="diagnosis_detail"),
]