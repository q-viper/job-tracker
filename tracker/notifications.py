from __future__ import annotations

import sys
from contextlib import nullcontext
from datetime import datetime, time, timedelta

from django.utils import timezone
from loguru import logger

from .models import PortalJobSuggestion, User
from .services import get_records, get_user_table, refresh_portal_jobs
from .smtp_handler import SMTPHandler


logger.remove()
logger.add(sys.stderr, format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}")


def should_notify_now(user: User, now: datetime, force: bool) -> bool:
    if force:
        return True
    if not user.notify_enabled:
        return False
    notify_time = user.notify_time or time(hour=8, minute=0)
    if now.time() < notify_time:
        return False
    if user.last_notification_sent_at and timezone.localtime(user.last_notification_sent_at).date() == now.date():
        return False
    return True


def upcoming_deadlines_for_user(user: User):
    try:
        table = get_user_table(user, "to-apply")
    except Exception:
        return []
    records = get_records(table, user)
    today = timezone.localdate()
    end_date = today + timedelta(days=user.deadline_notification_window_days or 7)
    matched = []
    for record in records:
        deadline_value = record.values.get("deadline")
        if not deadline_value:
            continue
        try:
            deadline_date = datetime.fromisoformat(deadline_value).date()
        except ValueError:
            continue
        if today <= deadline_date <= end_date:
            matched.append((deadline_date, record))
    matched.sort(key=lambda item: item[0])
    return matched


def refresh_portals_for_user(user: User):
    if not user.auto_fetch_portals:
        return []
    try:
        portal_table = get_user_table(user, "portals")
    except Exception:
        return []

    discoveries = []
    for record in get_records(portal_table, user):
        previous_urls = set(
            PortalJobSuggestion.objects.filter(user=user, portal_row_key=record.row_key).values_list("url", flat=True)
        )
        try:
            jobs = refresh_portal_jobs(user, portal_table, record.row_key)
        except Exception as exc:
            logger.warning(f"Portal refresh failed for {record.label}: {exc}")
            continue
        new_jobs = [job for job in jobs if job.url not in previous_urls]
        if new_jobs:
            discoveries.append((record, new_jobs))
    return discoveries


def build_email_body(user: User, deadlines, portal_discoveries) -> str:
    lines = [
        f"Hello {user.display_name or user.username},",
        "",
        "Here is your job-search update.",
        "",
    ]

    if deadlines:
        lines.append("Upcoming deadlines:")
        for deadline_date, record in deadlines:
            lines.append(f"- {deadline_date.isoformat()} | {record.label}")
            url = record.values.get("url", "")
            if url:
                lines.append(f"  {url}")
        lines.append("")
    else:
        lines.append("No upcoming deadlines were found in your current notification window.")
        lines.append("")

    if portal_discoveries:
        lines.append("New portal matches:")
        for portal_record, jobs in portal_discoveries:
            lines.append(f"- {portal_record.label}: {len(jobs)} new matches")
            for job in jobs[:5]:
                lines.append(f"  * {job.title}")
                lines.append(f"    {job.url}")
        lines.append("")
    else:
        lines.append("No new portal matches were found during auto-fetch.")
        lines.append("")

    lines.extend(
        [
            "You can change notification timing, email destination, and auto-fetch behavior from the Settings page in the app.",
            "",
            "Job Tracker",
        ]
    )
    return "\n".join(lines)


def run_notifications(*, force: bool = False, dry_run: bool = False) -> int:
    now = timezone.localtime()
    users = User.objects.filter(notify_enabled=True)
    processed = 0
    if not dry_run and not SMTPHandler.is_configured():
        logger.info("SMTP is not configured. Skipping email delivery features.")
        return 0

    smtp_context = nullcontext(None) if dry_run else SMTPHandler.from_env()

    try:
        with smtp_context as smtp_handler:
            if smtp_handler:
                logger.info("SMTP connection established successfully.")

            for user in users:
                recipient = (user.notification_email or user.email or "").strip()
                if not recipient:
                    logger.warning(f"Skipping {user.username}: no notification email configured.")
                    continue
                if not should_notify_now(user, now, force):
                    continue

                deadlines = upcoming_deadlines_for_user(user)
                portal_discoveries = refresh_portals_for_user(user)
                if not deadlines and not portal_discoveries and not force:
                    logger.info(f"Skipping {user.username}: no deadlines or new portal jobs to report.")
                    continue

                subject = f"Job Tracker Update for {now.date().isoformat()}"
                body = build_email_body(user, deadlines, portal_discoveries)

                if dry_run:
                    logger.info(f"Dry run for {user.username} -> {recipient}")
                    logger.info(body)
                else:
                    assert smtp_handler is not None
                    smtp_handler.send_email(recipient, subject, body)
                    user.last_notification_sent_at = timezone.now()
                    user.save(update_fields=["last_notification_sent_at"])
                    logger.info(f"Notification sent to {user.username} at {recipient}")
                processed += 1
    except Exception as exc:
        logger.error(f"Notifier failed: {exc}")
        raise

    logger.info(f"Finished. Users processed: {processed}")
    return processed
