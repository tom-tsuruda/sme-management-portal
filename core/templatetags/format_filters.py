from django import template

from core.formatters import format_amount

register = template.Library()


@register.filter
def amount(value):
    """
    既存の core.formatters.format_amount をテンプレートで使うためのフィルター。
    例：
    {{ value|amount }}
    """
    return format_amount(value)