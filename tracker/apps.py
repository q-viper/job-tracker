from django.apps import AppConfig


class TrackerConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "tracker"
    verbose_name = "Job Tracker"

    def ready(self):
        import os

        if os.environ.get("RUN_MAIN") != "true":
            return
        from .notification_scheduler import start_notification_scheduler

        start_notification_scheduler()
