from django.apps import AppConfig


START_SCHEDULER_COMMANDS = {"runserver", "serve"}
SKIP_SCHEDULER_COMMANDS = {
    "check",
    "collectstatic",
    "compilemessages",
    "createsuperuser",
    "dbshell",
    "diffsettings",
    "dumpdata",
    "findstatic",
    "flush",
    "inspectdb",
    "loaddata",
    "makemessages",
    "makemigrations",
    "migrate",
    "notify",
    "send_notifications.py",
    "shell",
    "showmigrations",
    "sqlflush",
    "sqlmigrate",
    "test",
}


def should_start_notification_scheduler(argv: list[str] | None = None) -> bool:
    import os
    import sys
    from pathlib import Path

    if os.getenv("DISABLE_NOTIFICATION_SCHEDULER", "").lower() in {"1", "true", "yes", "on"}:
        return False

    args = argv if argv is not None else sys.argv
    command_names = {Path(arg).name for arg in args[:2]}
    command_names.update(arg for arg in args[:2] if not arg.startswith("-"))
    if command_names & SKIP_SCHEDULER_COMMANDS:
        return False

    if command_names & START_SCHEDULER_COMMANDS:
        return os.environ.get("RUN_MAIN") == "true"

    if os.environ.get("RUN_MAIN") == "false":
        return False

    return True


class TrackerConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "tracker"
    verbose_name = "Job Tracker"

    def ready(self):
        if not should_start_notification_scheduler():
            return
        from .notification_scheduler import start_notification_scheduler

        start_notification_scheduler()
