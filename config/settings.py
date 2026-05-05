"""
Django settings for config project.

現段階では、社長が一人で会社の状況を確認するための
Excelベース軽量ポータルとして使う設定にしています。
本格運用・外部公開が必要になった段階で、環境変数やDB設定を強化します。
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


# -----------------------------------------------------------------------------
# 基本設定
# -----------------------------------------------------------------------------
SECRET_KEY = os.environ.get("SECRET_KEY", "django-insecure-dev-only-change-me")

DEBUG = os.environ.get("DEBUG", "True").lower() in ["1", "true", "yes", "on"]

ALLOWED_HOSTS = [
    host.strip()
    for host in os.environ.get("ALLOWED_HOSTS", "localhost,127.0.0.1").split(",")
    if host.strip()
]


# -----------------------------------------------------------------------------
# アプリケーション
# -----------------------------------------------------------------------------
INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "django.contrib.humanize",

    "core",
    "dashboard",
    "documents",
    "tasks",
    "questionnaires",
    "workflows",
    "expenses",
    "organizations",
    "manufacturing",
    "kpi",
    "governance",
    "notifications",
    "accounting",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "config.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "config.wsgi.application"


# -----------------------------------------------------------------------------
# データベース
# -----------------------------------------------------------------------------
# 今は Excel + SQLite の軽量運用で十分です。
# 複数人で本格運用する段階になったら PostgreSQL 化を検討します。
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}


# -----------------------------------------------------------------------------
# パスワード検証
# -----------------------------------------------------------------------------
AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]


# -----------------------------------------------------------------------------
# 日本語・日本時間
# -----------------------------------------------------------------------------
LANGUAGE_CODE = "ja"
TIME_ZONE = "Asia/Tokyo"

USE_I18N = True
USE_TZ = True


# -----------------------------------------------------------------------------
# 静的ファイル・アップロードファイル
# -----------------------------------------------------------------------------
STATIC_URL = "static/"

MEDIA_URL = "/media/"
MEDIA_ROOT = BASE_DIR / "media"

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"