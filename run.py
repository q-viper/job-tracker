from __future__ import annotations

import os
from pathlib import Path

import typer
from dotenv import load_dotenv
from django.core.management.base import CommandError

PROJECT_ROOT = Path(__file__).resolve().parent
app = typer.Typer(
    help="Run and maintain the Job Tracker application.",
    no_args_is_help=True,
    pretty_exceptions_enable=False,
)


def setup_django() -> None:
    load_dotenv(PROJECT_ROOT / ".env")
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")
    import django

    django.setup()


def get_workbook_path(excel: Path | None) -> Path:
    workbook = excel or (PROJECT_ROOT / "Applications.xlsx")
    if not workbook.exists():
        raise typer.BadParameter(f"Workbook not found: {workbook}")
    return workbook.resolve()


@app.command()
def serve(
    host: str = typer.Option("127.0.0.1", help="Bind address for the Django dev server."),
    port: int = typer.Option(8000, min=1, max=65535, help="Port for the Django dev server."),
) -> None:
    """Run the Django development server."""
    setup_django()
    from django.core.management import call_command

    try:
        call_command("runserver", f"{host}:{port}")
    except CommandError as exc:
        message = str(exc)
        if "permission" in message.lower() and "port" in message.lower():
            raise typer.BadParameter(
                f"{message} Try a higher unprivileged port, for example: "
                f"`python3 run.py serve --port 8080`."
            ) from exc
        raise


@app.command()
def makemigrations(
    app_label: str = typer.Option("tracker", help="Django app label to create migrations for."),
) -> None:
    """Generate Django migrations."""
    setup_django()
    from django.core.management import call_command

    call_command("makemigrations", app_label)


@app.command()
def migrate() -> None:
    """Apply Django migrations."""
    setup_django()
    from django.core.management import call_command

    call_command("migrate")


@app.command()
def bootstrap(
    excel: Path | None = typer.Option(None, help="Workbook path for the seeded test user import."),
    replace: bool = typer.Option(False, help="Replace previously imported test-user data before import."),
) -> None:
    """Create local users and import workbook data."""
    setup_django()
    from django.core.management import call_command

    workbook = get_workbook_path(excel)
    call_command("bootstrap_job_tracker", excel=str(workbook), replace=replace)


@app.command()
def notify(
    force: bool = typer.Option(False, help="Send even if the configured time has not been reached."),
    dry_run: bool = typer.Option(False, help="Build notifications without sending email."),
) -> None:
    """Run the SMTP notification flow."""
    setup_django()
    from tracker.notifications import run_notifications

    run_notifications(force=force, dry_run=dry_run)


@app.command()
def check() -> None:
    """Run Django system checks."""
    setup_django()
    from django.core.management import call_command

    call_command("check")


@app.command()
def test(
    app_label: str = typer.Option("tracker", help="Optional Django app label to scope the test run."),
) -> None:
    """Run the Django test suite."""
    setup_django()
    from django.core.management import call_command

    call_command("test", app_label)


def main() -> None:
    app()


if __name__ == "__main__":
    main()
