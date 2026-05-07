from __future__ import annotations

import os
import threading
from datetime import timedelta

from django.utils import timezone

from .notifications import run_notifications

_lock = threading.Lock()
_started = False


def _notifications_loop() -> None:
    while True:
        try:
            run_notifications(force=False, dry_run=False)
        except Exception:
            pass
        next_run = timezone.now() + timedelta(hours=1)
        sleep_seconds = max((next_run - timezone.now()).total_seconds(), 300)
        threading.Event().wait(sleep_seconds)


def start_notification_scheduler() -> None:
    global _started
    if os.getenv("DISABLE_NOTIFICATION_SCHEDULER", "").lower() in {"1", "true", "yes", "on"}:
        return
    with _lock:
        if _started:
            return
        worker = threading.Thread(target=_notifications_loop, name="job-tracker-notifier", daemon=True)
        worker.start()
        _started = True
