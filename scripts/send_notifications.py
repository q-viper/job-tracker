from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

load_dotenv(PROJECT_ROOT / ".env")
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")

import django  # noqa: E402

django.setup()

from tracker.notifications import run_notifications  # noqa: E402


def parse_args():
    parser = argparse.ArgumentParser(description="Send deadline and portal notifications to users.")
    parser.add_argument("--force", action="store_true", help="Send even if the configured time has not been reached.")
    parser.add_argument("--dry-run", action="store_true", help="Build notifications without sending email.")
    return parser.parse_args()


def main():
    args = parse_args()
    run_notifications(force=args.force, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
