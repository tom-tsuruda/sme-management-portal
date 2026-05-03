from django.shortcuts import render

# Create your views here.
from django.shortcuts import redirect, render

from .services import (
    find_notification_by_id,
    load_notifications,
    load_unread_notifications,
    mark_notification_as_read,
    mark_notification_as_unread,
)


def notification_list(request):
    all_notifications = load_notifications()
    notifications = all_notifications

    keyword = request.GET.get("q", "").strip()
    category = request.GET.get("category", "").strip()
    priority = request.GET.get("priority", "").strip()
    read_status = request.GET.get("read_status", "").strip()

    if keyword:
        filtered = []

        for notification in notifications:
            target_text = " ".join([
                str(notification.get("id", "")),
                str(notification.get("title", "")),
                str(notification.get("message", "")),
                str(notification.get("target_user", "")),
                str(notification.get("category", "")),
                str(notification.get("priority", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered.append(notification)

        notifications = filtered

    if category:
        notifications = [
            notification for notification in notifications
            if str(notification.get("category", "")) == category
        ]

    if priority:
        notifications = [
            notification for notification in notifications
            if str(notification.get("priority", "")) == priority
        ]

    if read_status == "unread":
        notifications = [
            notification for notification in notifications
            if str(notification.get("is_read")) not in ["1", "True", "true", "既読"]
        ]

    if read_status == "read":
        notifications = [
            notification for notification in notifications
            if str(notification.get("is_read")) in ["1", "True", "true", "既読"]
        ]

    categories = sorted(set([
        notification.get("category", "")
        for notification in all_notifications
        if notification.get("category", "")
    ]))

    priorities = ["高", "中", "低"]

    unread_count = len(load_unread_notifications())
    high_priority_unread_count = len([
        notification for notification in load_unread_notifications()
        if notification.get("priority") == "高"
    ])

    context = {
        "notifications": notifications,
        "keyword": keyword,
        "category": category,
        "priority": priority,
        "read_status": read_status,
        "categories": categories,
        "priorities": priorities,
        "total_count": len(all_notifications),
        "display_count": len(notifications),
        "unread_count": unread_count,
        "high_priority_unread_count": high_priority_unread_count,
    }

    return render(request, "notifications/notification_list.html", context)


def notification_detail(request, notification_id):
    notification = find_notification_by_id(notification_id)

    if notification:
        mark_notification_as_read(notification_id)
        notification = find_notification_by_id(notification_id)

    return render(request, "notifications/notification_detail.html", {
        "notification": notification,
    })


def mark_read(request, notification_id):
    mark_notification_as_read(notification_id)
    return redirect("notifications:notification_list")


def mark_unread(request, notification_id):
    mark_notification_as_unread(notification_id)
    return redirect("notifications:notification_list")