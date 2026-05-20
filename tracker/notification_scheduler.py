from __future__ import annotations

import os
import threading

from loguru import logger

from .notifications import notification_due_users, run_notifications

_lock = threading.Lock()
_started = False
_stop_event = threading.Event()


def notification_scheduler_poll_seconds() -> int:
    try:
        return max(int(os.getenv("NOTIFICATION_SCHEDULER_POLL_SECONDS", "60")), 10)
    except (TypeError, ValueError):
        return 60


def _notifications_loop() -> None:
    while True:
        try:
            due_users = notification_due_users()
            if due_users:
                logger.info(f"Scheduled notification check found {len(due_users)} due user(s).")
                run_notifications(force=False, dry_run=False)
        except Exception as exc:
            logger.exception(f"Scheduled notification check failed: {exc}")
        _stop_event.wait(notification_scheduler_poll_seconds())


def start_notification_scheduler() -> None:
    global _started
    if os.getenv("DISABLE_NOTIFICATION_SCHEDULER", "").lower() in {"1", "true", "yes", "on"}:
        logger.info("Notification scheduler disabled by DISABLE_NOTIFICATION_SCHEDULER.")
        return
    with _lock:
        if _started:
            return
        worker = threading.Thread(target=_notifications_loop, name="job-tracker-notifier", daemon=True)
        worker.start()
        _started = True
        logger.info(f"Notification scheduler started. Polling every {notification_scheduler_poll_seconds()} seconds.")
