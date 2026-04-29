from pathlib import Path

from django.core.management.base import BaseCommand, CommandError

from tracker.services import bootstrap_local_users


class Command(BaseCommand):
    help = "Create local admin/test users from .env, seed tables, and import Applications.xlsx for the test user."

    def add_arguments(self, parser):
        parser.add_argument(
            "--excel",
            default="Applications.xlsx",
            help="Path to the workbook to import for the seeded test user.",
        )
        parser.add_argument(
            "--replace",
            action="store_true",
            help="Replace previously imported test user data before importing again.",
        )

    def handle(self, *args, **options):
        excel_path = Path(options["excel"]).resolve()
        if not excel_path.exists():
            raise CommandError(f"Workbook not found: {excel_path}")

        result = bootstrap_local_users(excel_path=excel_path, replace=options["replace"])
        self.stdout.write(self.style.SUCCESS("Local users bootstrapped successfully."))
        self.stdout.write(f"Admin user: {result['admin'].username}")
        self.stdout.write(f"Test user: {result['test_user'].username}")
        for table_slug, count in result["imported_counts"].items():
            self.stdout.write(f" - {table_slug}: {count} imported rows")
