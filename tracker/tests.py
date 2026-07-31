import json
from datetime import datetime, time, timedelta
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from types import SimpleNamespace
from unittest.mock import patch
from uuid import uuid4

import requests
from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone
from openpyxl import Workbook, load_workbook

from .apps import should_start_notification_scheduler
from .chat_handler import ChatHandler
from .forms import AppliedCreateForm, AppliedUpdateForm
from .models import (
    ApplicationHistoryEvent,
    ApplicationTag,
    ApplicationTagAssignment,
    ChatModelConfig,
    DashboardBoard,
    DashboardBoardItem,
    DashboardPanelPreference,
    DynamicColumn,
    DynamicTable,
    PortalJobSuggestion,
)
from .notifications import notification_due_users, run_notifications, send_notification_to_user
from .rag import chat_with_rag
from .services import (
    aggregate_records,
    add_portal_job_to_to_apply,
    build_applied_flow_chart,
    build_active_applied_pipeline_chart,
    build_dashboard,
    delete_record,
    ensure_user_workspace,
    extract_jobs_from_portal_html,
    get_record,
    get_records,
    infer_fields_from_url_details,
    import_workbook_for_user,
    matches_any_keyword,
    move_to_applied,
    replace_applied_history_events,
    save_record,
    scrape_jobs_from_portal,
    seed_sample_jobs_for_user,
    sync_application_tags,
)
from .views import ToApplyListView


class DummySMTPContext:
    def __init__(self):
        self.sent_messages: list[tuple[str, str, str]] = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return None

    def send_email(self, recipient: str, subject: str, body: str) -> None:
        self.sent_messages.append((recipient, subject, body))


class DummySelectiveFailSMTPContext(DummySMTPContext):
    def __init__(self, failed_recipient: str):
        super().__init__()
        self.failed_recipient = failed_recipient

    def send_email(self, recipient: str, subject: str, body: str) -> None:
        if recipient == self.failed_recipient:
            raise RuntimeError("recipient refused")
        super().send_email(recipient, subject, body)


class DynamicRecordServiceTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="tester",
            password="TestPassword123!",
        )
        self.to_apply = DynamicTable.objects.create(
            owner=self.user,
            name="To Apply",
            slug="to-apply",
        )
        self.applied = DynamicTable.objects.create(
            owner=self.user,
            name="Applied",
            slug="applied",
        )
        self.portals = DynamicTable.objects.create(
            owner=self.user,
            name="Portals",
            slug="portals",
        )

        for position, (table, name, slug, data_type) in enumerate(
            [
                (self.to_apply, "Role / Title", "role-title", DynamicColumn.DataType.TEXT),
                (self.to_apply, "URL", "url", DynamicColumn.DataType.URL),
                (self.to_apply, "Type", "type", DynamicColumn.DataType.TEXT),
                (self.to_apply, "Deadline", "deadline", DynamicColumn.DataType.DATE),
                (self.portals, "URL", "url", DynamicColumn.DataType.URL),
                (self.portals, "Name", "name", DynamicColumn.DataType.TEXT),
                (self.applied, "Date", "date", DynamicColumn.DataType.DATE),
                (self.applied, "Role / Title", "role-title", DynamicColumn.DataType.TEXT),
                (self.applied, "URL", "url", DynamicColumn.DataType.URL),
                (self.applied, "Type", "type", DynamicColumn.DataType.TEXT),
                (self.applied, "Result", "result", DynamicColumn.DataType.STATUS),
                (self.applied, "AI Reliance", "ai-reliance", DynamicColumn.DataType.TEXT),
                (
                    self.applied,
                    "Application Progress",
                    "application-progress",
                    DynamicColumn.DataType.STATUS,
                ),
            ],
            start=1,
        ):
            DynamicColumn.objects.create(
                table=table,
                name=name,
                slug=slug,
                data_type=data_type,
                position=position,
            )

    def build_uploaded_workbook(self) -> SimpleUploadedFile:
        workbook = Workbook()
        applied_sheet = workbook.active
        applied_sheet.title = "Applied"
        applied_sheet.append(["Date", "URL", "Name", "Source", "Type", "Result", "Result On", "AI Reliance", "Completed Stages", "Remarks"])
        applied_sheet.append(["04/05/2026", "https://example.com/imported-applied", "Imported Company", "Example Board", "Job", "", "", "", "", ""])

        portals_sheet = workbook.create_sheet("Portals")
        portals_sheet.append(["S.N", "URL", "Name"])
        portals_sheet.append([1, "https://example.com/imported-portal", "Imported Portal"])

        to_apply_sheet = workbook.create_sheet("To Apply")
        to_apply_sheet.append(["S.N.", "URL", "Deadline", "Type", "English Test"])
        to_apply_sheet.append([1, "https://example.com/imported-to-apply", "09/05/2026", "PhD", ""])

        payload = BytesIO()
        workbook.save(payload)
        payload.seek(0)
        return SimpleUploadedFile(
            "Applications.xlsx",
            payload.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    def test_save_record_groups_cells_into_single_row(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "ML Engineer",
                "url": "https://example.com/jobs/ml-engineer",
                "type": "Job",
            },
        )
        records = get_records(self.to_apply, self.user)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].values["role-title"], "ML Engineer")
        self.assertEqual(records[0].values["type"], "Job")

    def test_applied_forms_allow_missing_url(self):
        DynamicColumn.objects.filter(table=self.applied, slug="url").update(is_required=True)

        create_form = AppliedCreateForm(
            data={"role-title": "Offline Referral Application", "type": "Job"},
            table=self.applied,
            user=self.user,
            source_records=[],
        )
        update_form = AppliedUpdateForm(
            data={"role-title": "Offline Referral Application", "type": "Job"},
            table=self.applied,
            user=self.user,
        )

        self.assertTrue(create_form.is_valid(), create_form.errors)
        self.assertTrue(update_form.is_valid(), update_form.errors)
        self.assertFalse(create_form.fields["url"].required)
        self.assertFalse(update_form.fields["url"].required)

    def test_move_to_applied_removes_source_row(self):
        source_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
            },
        )
        move_to_applied(
            user=self.user,
            applied_table=self.applied,
            to_apply_table=self.to_apply,
            values={"date": "2026-04-29"},
            source_row_key=source_row_key,
        )
        self.assertEqual(len(get_records(self.to_apply, self.user)), 0)
        applied_records = get_records(self.applied, self.user)
        self.assertEqual(len(applied_records), 1)
        self.assertEqual(applied_records[0].values["role-title"], "Vision PhD")
        self.assertEqual(applied_records[0].values["application-progress"], "Applied")

    def test_delete_record_removes_job_history_and_tag_assignments(self):
        row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
            },
        )
        sync_application_tags(self.user, "to-apply", row_key, ["Priority", "Dummy"])

        delete_record(self.to_apply, self.user, row_key)

        self.assertFalse(
            ApplicationHistoryEvent.objects.filter(
                user=self.user,
                table_slug="to-apply",
                record_row_key=row_key,
            ).exists()
        )
        self.assertFalse(
            ApplicationTagAssignment.objects.filter(
                user=self.user,
                table_slug="to-apply",
                record_row_key=row_key,
            ).exists()
        )

    def test_move_to_applied_preserves_tags_and_history_on_same_row_key(self):
        source_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
            },
        )
        sync_application_tags(self.user, "to-apply", source_row_key, ["Priority", "Dummy"])

        moved_row_key = move_to_applied(
            user=self.user,
            applied_table=self.applied,
            to_apply_table=self.to_apply,
            values={"date": "2026-04-29"},
            source_row_key=source_row_key,
        )

        self.assertEqual(moved_row_key, source_row_key)
        self.assertEqual(len(get_records(self.to_apply, self.user)), 0)
        applied_record = get_records(self.applied, self.user)[0]
        self.assertEqual(applied_record.row_key, source_row_key)
        self.assertCountEqual(
            list(
                ApplicationTagAssignment.objects.filter(
                    user=self.user,
                    table_slug="applied",
                    record_row_key=source_row_key,
                ).values_list("tag__name", flat=True)
            ),
            ["Priority", "Dummy"],
        )
        history = list(
            ApplicationHistoryEvent.objects.filter(
                user=self.user,
                table_slug="applied",
                record_row_key=source_row_key,
            ).order_by("occurred_at", "id")
        )
        self.assertEqual(history[0].to_stage, "To Apply")
        self.assertEqual(history[-1].to_stage, "Applied")
        self.assertLess(history[0].occurred_at, history[-1].occurred_at)

    def test_replacing_applied_history_dates_preserves_to_apply_event(self):
        source_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
            },
        )
        moved_row_key = move_to_applied(
            user=self.user,
            applied_table=self.applied,
            to_apply_table=self.to_apply,
            values={"date": "2026-04-29"},
            source_row_key=source_row_key,
        )

        replace_applied_history_events(
            self.user,
            moved_row_key,
            {
                "Applied": "2026-04-29",
                "Rejected": "2026-05-03",
            },
        )

        history = list(
            ApplicationHistoryEvent.objects.filter(
                user=self.user,
                table_slug="applied",
                record_row_key=moved_row_key,
            )
        )
        self.assertTrue(
            any(
                event.event_type == ApplicationHistoryEvent.EventType.ADDED_TO_APPLY
                and event.to_stage == "To Apply"
                for event in history
            )
        )
        self.assertTrue(
            any(
                event.event_type == ApplicationHistoryEvent.EventType.STAGE_CHANGE
                and event.to_stage == "Rejected"
                for event in history
            )
        )
        ordered_history = sorted(history, key=lambda event: (event.occurred_at, event.id))
        self.assertEqual(ordered_history[0].to_stage, "To Apply")
        self.assertEqual(ordered_history[1].to_stage, "Applied")

    def test_import_workbook_uses_stage_date_columns_for_applied_history(self):
        workbook = Workbook()
        applied_sheet = workbook.active
        applied_sheet.title = "Applied"
        applied_sheet.append(
            [
                "Date",
                "URL",
                "Name",
                "Type",
                "Application Progress",
                "Screening Date",
                "Interview Date",
                "Assignment Date",
            ]
        )
        applied_sheet.append(
            [
                "2026-06-24",
                "https://example.com/assignment",
                "Assignment Company",
                "PhD",
                "Assignment",
                "2026-06-25",
                "2026-06-26",
                "2026-06-27",
            ]
        )

        temp_path: Path | None = None
        try:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_path = Path(temp_file.name)
            workbook.save(temp_path)

            import_workbook_for_user(temp_path, self.user, replace=True)
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink()

        record = get_records(self.applied, self.user)[0]
        history = list(
            ApplicationHistoryEvent.objects.filter(
                user=self.user,
                table_slug="applied",
                record_row_key=record.row_key,
            ).order_by("occurred_at", "id")
        )

        self.assertEqual(
            [event.to_stage for event in history],
            ["Applied", "Screening", "Interview", "Assignment"],
        )
        self.assertEqual(history[-1].from_stage, "Interview")

    def test_import_workbook_routes_sparse_assignment_progress_through_prior_stages(self):
        workbook = Workbook()
        applied_sheet = workbook.active
        applied_sheet.title = "Applied"
        applied_sheet.append(["Date", "URL", "Name", "Type", "Application Progress"])
        applied_sheet.append(
            [
                "2026-06-24",
                "https://example.com/sparse-assignment",
                "Sparse Assignment Company",
                "PhD",
                "Assignment",
            ]
        )

        temp_path: Path | None = None
        try:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_path = Path(temp_file.name)
            workbook.save(temp_path)

            import_workbook_for_user(temp_path, self.user, replace=True)
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink()

        record = get_records(self.applied, self.user)[0]
        history = list(
            ApplicationHistoryEvent.objects.filter(
                user=self.user,
                table_slug="applied",
                record_row_key=record.row_key,
            ).order_by("occurred_at", "id")
        )

        self.assertEqual(
            [event.to_stage for event in history],
            ["Applied", "Screening", "Interview", "Assignment"],
        )
        self.assertEqual(history[-1].from_stage, "Interview")

    def test_move_to_applied_preserves_deadline_when_applied_tracks_it(self):
        tables = ensure_user_workspace(self.user)
        source_row_key = save_record(
            tables["to-apply"],
            self.user,
            {
                "role-title": "Deadline Role",
                "url": "https://example.com/deadline-role",
                "type": "Job",
                "deadline": "2026-05-15",
            },
        )

        move_to_applied(
            user=self.user,
            applied_table=tables["applied"],
            to_apply_table=tables["to-apply"],
            values={"date": "2026-05-01"},
            source_row_key=source_row_key,
        )

        applied_record = next(record for record in get_records(tables["applied"], self.user) if record.row_key == source_row_key)
        self.assertEqual(applied_record.values["deadline"], "2026-05-15")

    def test_move_to_applied_does_not_infer_or_scrape_url_fields(self):
        tables = ensure_user_workspace(self.user)
        source_row_key = save_record(
            tables["to-apply"],
            self.user,
            {
                "role-title": "Portal Role",
                "name": "Saved Organization",
                "source": "Saved Portal",
                "url": "https://example.com/jobs/wrong-title-if-scraped",
                "type": "Job",
            },
            apply_inference=False,
        )

        with patch("tracker.services.infer_fields_from_url") as infer_mock:
            move_to_applied(
                user=self.user,
                applied_table=tables["applied"],
                to_apply_table=tables["to-apply"],
                values={"date": "2026-05-01"},
                source_row_key=source_row_key,
            )

        infer_mock.assert_not_called()
        applied_record = next(record for record in get_records(tables["applied"], self.user) if record.row_key == source_row_key)
        self.assertEqual(applied_record.values["role-title"], "Portal Role")
        self.assertEqual(applied_record.values["name"], "Saved Organization")
        self.assertEqual(applied_record.values["source"], "Saved Portal")

    def test_portal_suggestion_adds_context_to_to_apply(self):
        tables = ensure_user_workspace(self.user)
        portal_row_key = save_record(
            tables["portals"],
            self.user,
            {
                "url": "https://jobs.example.com",
                "name": "Example Careers",
            },
            apply_inference=False,
        )
        suggestion = PortalJobSuggestion.objects.create(
            user=self.user,
            portal_table=tables["portals"],
            portal_row_key=portal_row_key,
            portal_name="Example Careers",
            portal_url="https://jobs.example.com",
            title="Computer Vision Researcher",
            url="https://jobs.example.com/computer-vision-researcher",
            summary="Research role",
        )

        row_key = add_portal_job_to_to_apply(self.user, suggestion)

        record = get_record(tables["to-apply"], self.user, row_key)
        self.assertIsNotNone(record)
        self.assertEqual(record.values["role-title"], "Computer Vision Researcher")
        self.assertEqual(record.values["name"], "Example Careers")
        self.assertEqual(record.values["source"], "Example Careers")

    def test_applied_progress_change_creates_history_event(self):
        row_key = save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
                "application-progress": "Applied",
            },
        )

        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
                "application-progress": "Interview",
            },
            row_key=row_key,
        )

        history = list(
            ApplicationHistoryEvent.objects.filter(
                user=self.user,
                table_slug="applied",
                record_row_key=row_key,
            ).order_by("occurred_at", "id")
        )
        self.assertEqual([event.to_stage for event in history], ["Applied", "Screening", "Interview"])
        self.assertEqual(history[-1].from_stage, "Screening")
        self.assertEqual(history[-1].to_stage, "Interview")

    def test_application_progress_aggregation_uses_inferred_result_status(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
                "result": "Rejected",
                "application-progress": "Applied",
            },
        )

        labels, values = aggregate_records(
            get_records(self.applied, self.user),
            group_by_column="application-progress",
            aggregation="count-records",
            metric_column="",
            limit=10,
        )

        self.assertEqual(dict(zip(labels, values)), {"Rejected": 1.0})

    def test_to_apply_list_shows_move_to_applied_button(self):
        source_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("to_apply_list"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Move to Applied")
        self.assertContains(response, reverse("to_apply_promote", args=[source_row_key]))

    def test_promote_to_applied_view_moves_record_and_opens_applied_edit(self):
        source_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
            },
        )
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("to_apply_promote", args=[source_row_key]),
            data={"next": "/to-apply/?view=table"},
        )

        self.assertRedirects(response, reverse("applied_edit", args=[source_row_key]))
        self.assertEqual(len(get_records(self.to_apply, self.user)), 0)
        applied_records = get_records(self.applied, self.user)
        self.assertEqual(len(applied_records), 1)
        self.assertEqual(applied_records[0].values["role-title"], "Vision PhD")
        self.assertEqual(applied_records[0].values["application-progress"], "Applied")

    def test_job_detail_view_shows_history_and_tags(self):
        row_key = save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
                "application-progress": "Interview",
            },
        )
        sync_application_tags(self.user, "applied", row_key, ["Priority", "PhD"])
        self.client.force_login(self.user)

        response = self.client.get(reverse("applied_detail", args=[row_key]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Progress Timeline")
        self.assertContains(response, "Priority")
        self.assertContains(response, "Applied")

    def test_dashboard_cards_link_to_job_detail_pages(self):
        to_apply_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Future PhD",
                "url": "https://example.com/future-phd",
                "type": "PhD",
                "deadline": "2026-05-10",
            },
        )
        applied_row_key = save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
                "application-progress": "Applied",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("dashboard"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, reverse("to_apply_detail", args=[to_apply_row_key]))
        self.assertContains(response, reverse("applied_detail", args=[applied_row_key]))

    def test_bulk_to_apply_actions_can_add_tags_and_move_selected_rows(self):
        first_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Role One",
                "url": "https://example.com/one",
                "type": "Job",
            },
        )
        second_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Role Two",
                "url": "https://example.com/two",
                "type": "PhD",
            },
        )
        self.client.force_login(self.user)

        add_tags_response = self.client.post(
            reverse("to_apply_bulk"),
            data={
                "bulk_action": "add_tags",
                "bulk_tags": "Priority, Dummy",
                "selected_row_keys": [str(first_row_key), str(second_row_key)],
            },
        )
        self.assertEqual(add_tags_response.status_code, 302)
        self.assertEqual(
            ApplicationTagAssignment.objects.filter(user=self.user, table_slug="to-apply").count(),
            4,
        )

        move_response = self.client.post(
            reverse("to_apply_bulk"),
            data={
                "bulk_action": "move_to_applied",
                "selected_row_keys": [str(first_row_key), str(second_row_key)],
            },
        )
        self.assertEqual(move_response.status_code, 302)
        self.assertEqual(len(get_records(self.to_apply, self.user)), 0)
        self.assertEqual(len(get_records(self.applied, self.user)), 2)
        self.assertEqual(
            ApplicationTagAssignment.objects.filter(user=self.user, table_slug="applied").count(),
            4,
        )

    def test_bulk_move_to_applied_with_single_selection_opens_applied_edit(self):
        row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Single Role",
                "url": "https://example.com/single",
                "type": "Job",
            },
        )
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("to_apply_bulk"),
            data={
                "bulk_action": "move_to_applied",
                "selected_row_keys": [str(row_key)],
                "next": "/to-apply/?view=table",
            },
        )

        self.assertRedirects(response, reverse("applied_edit", args=[row_key]))
        self.assertEqual(len(get_records(self.to_apply, self.user)), 0)
        self.assertEqual(len(get_records(self.applied, self.user)), 1)

    def test_applied_bulk_action_can_remove_tags(self):
        row_key = save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/phd",
                "type": "PhD",
                "application-progress": "Applied",
            },
        )
        sync_application_tags(self.user, "applied", row_key, ["Priority", "PhD"])
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("applied_bulk"),
            data={
                "bulk_action": "remove_tags",
                "bulk_tags": "Priority",
                "selected_row_keys": [str(row_key)],
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertCountEqual(
            list(
                ApplicationTagAssignment.objects.filter(
                    user=self.user,
                    table_slug="applied",
                    record_row_key=row_key,
                ).values_list("tag__name", flat=True)
            ),
            ["PhD"],
        )

    def test_seed_sample_jobs_for_user_creates_history_ready_demo_records(self):
        seed_sample_jobs_for_user(self.user)

        applied_records = get_records(self.applied, self.user)
        to_apply_records = get_records(self.to_apply, self.user)
        self.assertEqual(len(applied_records), 5)
        self.assertEqual(len(to_apply_records), 0)
        self.assertTrue(ApplicationTag.objects.filter(user=self.user, name="Dummy").exists())
        self.assertTrue(all(record.values.get("number-of-positions") for record in applied_records))
        self.assertTrue(all(record.values.get("deadline") for record in applied_records))

        flow_chart = build_applied_flow_chart(applied_records, self.user)
        sankey_data = flow_chart["chart_config"]["data"][0]
        self.assertIn("To Apply", "".join(sankey_data["node"]["label"]))
        self.assertTrue(sankey_data["link"]["value"])

    def test_applied_flow_includes_to_apply_and_reconciles_current_progress(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Queued Role",
                "url": "https://example.com/queued",
                "type": "Job",
            },
        )
        applied_row_key = save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Rejected Role",
                "url": "https://example.com/rejected",
                "type": "Job",
                "result": "Rejected",
                "application-progress": "Applied",
            },
        )
        ApplicationHistoryEvent.objects.filter(
            user=self.user,
            table_slug="applied",
            record_row_key=applied_row_key,
            to_stage="Applied",
        ).delete()

        flow_chart = build_applied_flow_chart(
            get_records(self.applied, self.user),
            self.user,
            to_apply_records=get_records(self.to_apply, self.user),
        )

        sankey_data = flow_chart["chart_config"]["data"][0]
        node_labels = sankey_data["node"]["label"]
        self.assertIn("To Apply (2)", node_labels)
        self.assertIn("Applied (1)", node_labels)
        self.assertIn("Rejected (1)", node_labels)
        links = sankey_data["link"]
        edges = {
            (node_labels[source], node_labels[target]): value
            for source, target, value in zip(links["source"], links["target"], links["value"])
        }
        self.assertEqual(edges[("To Apply (2)", "Applied (1)")], 1)
        self.assertEqual(edges[("Applied (1)", "Rejected (1)")], 1)

    def test_applied_flow_routes_withdrawn_records_through_applied(self):
        row_key = save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Withdrawn Role",
                "url": "https://example.com/withdrawn",
                "type": "Job",
                "result": "Withdrawn",
                "application-progress": "Applied",
            },
        )
        ApplicationHistoryEvent.objects.filter(
            user=self.user,
            table_slug="applied",
            record_row_key=row_key,
        ).delete()
        record_application_events = [
            (
                ApplicationHistoryEvent.EventType.APPLIED,
                "",
                "Applied",
                timezone.now() - timedelta(days=4),
            ),
            (
                ApplicationHistoryEvent.EventType.ADDED_TO_APPLY,
                "",
                "To Apply",
                timezone.now() - timedelta(days=3),
            ),
            (
                ApplicationHistoryEvent.EventType.STAGE_CHANGE,
                "To Apply",
                "Withdrawn",
                timezone.now() - timedelta(days=1),
            ),
        ]
        for event_type, from_stage, to_stage, occurred_at in record_application_events:
            ApplicationHistoryEvent.objects.create(
                user=self.user,
                table_slug="applied",
                record_row_key=row_key,
                event_type=event_type,
                from_stage=from_stage,
                to_stage=to_stage,
                occurred_at=occurred_at,
            )

        flow_chart = build_applied_flow_chart(get_records(self.applied, self.user), self.user)

        sankey_data = flow_chart["chart_config"]["data"][0]
        node_labels = sankey_data["node"]["label"]
        links = sankey_data["link"]
        edges = {
            (node_labels[source], node_labels[target]): value
            for source, target, value in zip(links["source"], links["target"], links["value"])
        }
        self.assertEqual(edges[("To Apply (1)", "Applied (1)")], 1)
        self.assertEqual(edges[("Applied (1)", "Withdrawn (1)")], 1)
        self.assertNotIn(("To Apply (1)", "Withdrawn (1)"), edges)

    def test_applied_flow_routes_sparse_assignment_history_through_interview(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-06-24",
                "role-title": "Assignment Role",
                "url": "https://example.com/assignment-role",
                "type": "PhD",
                "application-progress": "Assignment",
            },
        )

        flow_chart = build_applied_flow_chart(get_records(self.applied, self.user), self.user)

        sankey_data = flow_chart["chart_config"]["data"][0]
        node_labels = sankey_data["node"]["label"]
        links = sankey_data["link"]
        edges = {
            (node_labels[source], node_labels[target]): value
            for source, target, value in zip(links["source"], links["target"], links["value"])
        }
        self.assertEqual(edges[("To Apply (1)", "Applied (1)")], 1)
        self.assertEqual(edges[("Applied (1)", "Screening (1)")], 1)
        self.assertEqual(edges[("Screening (1)", "Interview (1)")], 1)
        self.assertEqual(edges[("Interview (1)", "Assignment (1)")], 1)
        self.assertNotIn(("Applied (1)", "Interview (1)"), edges)
        self.assertNotIn(("Applied (1)", "Assignment (1)"), edges)

    def test_ai_reliance_chart_forces_values_into_categories(self):
        tables = ensure_user_workspace(self.user)
        applied_table = tables["applied"]
        save_record(
            applied_table,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "No AI Role",
                "url": "https://example.com/no-ai",
                "type": "Job",
                "source": "LinkedIn",
                "ai-reliance": "0%",
            },
        )
        save_record(
            applied_table,
            self.user,
            {
                "date": "2026-04-30",
                "role-title": "Proofread Role",
                "url": "https://example.com/proofread",
                "type": "Job",
                "source": "Referral",
                "ai-reliance": "just grammar proofreading",
            },
        )
        save_record(
            applied_table,
            self.user,
            {
                "date": "2026-05-01",
                "role-title": "Heavy AI Role",
                "url": "https://example.com/heavy-ai",
                "type": "Job",
                "source": "University Careers",
                "ai-reliance": "mostly generated with AI",
            },
        )

        labels, values = aggregate_records(
            get_records(applied_table, self.user),
            group_by_column="ai-reliance",
            aggregation="count-records",
            metric_column="",
            limit=10,
        )

        self.assertEqual(dict(zip(labels, values)), {"None": 1.0, "Low": 1.0, "High": 1.0})
        dashboard_payload = build_dashboard(self.user)
        widgets = {widget.template_key: widget for widget in dashboard_payload["all_dashboard_widgets"]}
        ai_dataset = widgets["applied-ai-reliance"].chart_config["datasets"][0]
        source_dataset = widgets["applied-source-mix"].chart_config["datasets"][0]
        self.assertEqual(
            len(set(ai_dataset["backgroundColor"])),
            len(ai_dataset["backgroundColor"]),
        )
        self.assertEqual(
            len(set(source_dataset["backgroundColor"])),
            len(source_dataset["backgroundColor"]),
        )

    def test_dashboard_includes_stage_timing_and_deadline_lead_widgets(self):
        ensure_user_workspace(self.user)
        seed_sample_jobs_for_user(self.user)

        dashboard_payload = build_dashboard(self.user)
        widgets = {widget.template_key: widget for widget in dashboard_payload["dashboard_widgets"]}

        self.assertIn("applied-stage-timing", widgets)
        self.assertIn("applied-deadline-lead", widgets)
        self.assertIn("applied-ai-reliance", widgets)
        self.assertIn("applied-over-time", widgets)
        self.assertIn("applied-outcome-mix", widgets)
        self.assertIn("applied-source-mix", widgets)
        self.assertIn("to-apply-type-mix", widgets)
        self.assertIn("Applied", widgets["applied-stage-timing"].chart_config["labels"])
        self.assertTrue(widgets["applied-stage-timing"].chart_config["datasets"][0]["data"])
        self.assertEqual(
            sum(widgets["applied-deadline-lead"].chart_config["datasets"][0]["data"]),
            5,
        )
        self.assertIn("Applied late", widgets["applied-deadline-lead"].chart_config["labels"])

    def test_applications_over_time_sorts_date_axis_ascending(self):
        tables = ensure_user_workspace(self.user)
        tables["applied"].columns.filter(slug="date").update(data_type=DynamicColumn.DataType.TEXT)
        for label, applied_on in (
            ("Latest Role", "2026-06-10"),
            ("Earliest Role", "2026-06-01"),
            ("Middle Role", "2026-06-05"),
        ):
            save_record(
                tables["applied"],
                self.user,
                {
                    "date": applied_on,
                    "role-title": label,
                    "url": f"https://example.com/{label.lower().replace(' ', '-')}",
                    "type": "Job",
                },
            )

        dashboard_payload = build_dashboard(self.user)
        widget = next(
            item
            for item in dashboard_payload["all_dashboard_widgets"]
            if item.template_key == "applied-over-time"
        )

        self.assertEqual(
            widget.chart_config["labels"],
            ["2026-06-01", "2026-06-05", "2026-06-10"],
        )

    def test_build_dashboard_includes_default_panel_layout(self):
        ensure_user_workspace(self.user)

        payload = build_dashboard(self.user)

        self.assertEqual(
            [panel["key"] for panel in payload["dashboard_panel_layout"]["left"]],
            ["stage-breakdown", "quick-actions", "plot-templates"],
        )
        self.assertEqual(
            [panel["key"] for panel in payload["dashboard_panel_layout"]["main"]],
            ["metrics-strip", "widget-canvas"],
        )
        self.assertEqual(
            [panel["key"] for panel in payload["dashboard_panel_layout"]["right"]],
            ["hidden-widgets", "upcoming-deadlines", "recent-submissions"],
        )

    def test_dashboard_customize_view_updates_panel_visibility(self):
        ensure_user_workspace(self.user)
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("dashboard_customize"),
            data={"zone_stage-breakdown": "hidden"},
        )

        self.assertEqual(response.status_code, 302)
        preference = DashboardPanelPreference.objects.get(user=self.user, panel_key="stage-breakdown")
        self.assertEqual(preference.zone, "hidden")

        dashboard_response = self.client.get(reverse("dashboard"))
        self.assertNotContains(dashboard_response, "Current progress mix")

    def test_build_dashboard_creates_default_saved_boards(self):
        ensure_user_workspace(self.user)

        payload = build_dashboard(self.user)

        self.assertEqual(DashboardBoard.objects.filter(user=self.user).count(), 3)
        self.assertEqual([board.name for board in payload["dashboard_boards"]], ["Home Lens", "Pipeline Lab", "Deadline Radar"])
        self.assertEqual(payload["selected_homepage_board"].name, "Home Lens")
        self.assertTrue(payload["selected_homepage_widgets"])

    def test_dashboard_boards_page_lists_saved_boards(self):
        ensure_user_workspace(self.user)
        build_dashboard(self.user)
        self.client.force_login(self.user)

        response = self.client.get(reverse("dashboard_boards"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Saved boards for home and deep work")
        self.assertContains(response, "Home Lens")
        self.assertContains(response, "Pipeline Lab")

    def test_dashboard_designer_view_is_separate_from_homepage(self):
        ensure_user_workspace(self.user)
        dashboard_payload = build_dashboard(self.user)
        board = dashboard_payload["dashboard_boards"][0]
        self.client.force_login(self.user)

        response = self.client.get(reverse("dashboard_designer", args=[board.id]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Drag cards by the dotted handle")
        self.assertContains(response, "Save board settings")
        self.assertContains(response, board.name)

        homepage = self.client.get(reverse("dashboard"))
        self.assertNotContains(homepage, "Drag cards by the dotted handle")
        self.assertContains(homepage, "A calmer view of your job search momentum")

    def test_dashboard_board_item_api_can_add_update_and_remove_widget_layout(self):
        ensure_user_workspace(self.user)
        dashboard_payload = build_dashboard(self.user)
        board = dashboard_payload["dashboard_boards"][0]
        existing_item_count = DashboardBoardItem.objects.filter(board=board).count()
        self.client.force_login(self.user)

        create_response = self.client.post(
            reverse("api_dashboard_board_items", args=[board.id]),
            data=json.dumps({"template_key": "applied-outcomes-waterfall"}),
            content_type="application/json",
        )

        self.assertEqual(create_response.status_code, 201)
        created_widget = create_response.json()["widget"]
        created_item = DashboardBoardItem.objects.get(board=board, id=created_widget["layout_item_id"])
        self.assertEqual(DashboardBoardItem.objects.filter(board=board).count(), existing_item_count + 1)

        update_response = self.client.generic(
            "PATCH",
            reverse("api_dashboard_board_item_detail", args=[board.id, created_item.id]),
            data=json.dumps({"card_size": "Full", "chart_height": "XL", "sort_order": 95}),
            content_type="application/json",
        )

        self.assertEqual(update_response.status_code, 200)
        created_item.refresh_from_db()
        self.assertEqual(created_item.card_size, "Full")
        self.assertEqual(created_item.chart_height, "XL")
        self.assertEqual(created_item.sort_order, 95)

        delete_response = self.client.delete(reverse("api_dashboard_board_item_detail", args=[board.id, created_item.id]))

        self.assertEqual(delete_response.status_code, 200)
        self.assertFalse(DashboardBoardItem.objects.filter(board=board, id=created_item.id).exists())

    def test_previewing_an_empty_saved_board_stays_empty(self):
        ensure_user_workspace(self.user)
        build_dashboard(self.user)
        empty_board = DashboardBoard.objects.create(
            user=self.user,
            name="Empty Board",
            description="No plots yet.",
            show_on_homepage=True,
            sort_order=99,
        )
        self.client.force_login(self.user)

        payload = build_dashboard(self.user, selected_board_id=empty_board.id)
        response = self.client.get(reverse("dashboard"), {"board": empty_board.id})

        self.assertEqual(payload["selected_homepage_board"], empty_board)
        self.assertEqual(payload["selected_homepage_widgets"], [])
        self.assertContains(response, "This board does not have any saved plots yet.")

    def test_workspace_bootstrap_adds_number_of_positions_to_job_tables(self):
        tables = ensure_user_workspace(self.user)

        to_apply_column = tables["to-apply"].columns.get(slug="number-of-positions")
        applied_column = tables["applied"].columns.get(slug="number-of-positions")

        self.assertEqual(to_apply_column.data_type, DynamicColumn.DataType.NUMBER)
        self.assertEqual(applied_column.data_type, DynamicColumn.DataType.NUMBER)
        self.assertEqual(
            tables["applied"].columns.get(slug="deadline").data_type,
            DynamicColumn.DataType.DATE,
        )
        self.assertEqual(tables["to-apply"].columns.get(slug="name").name, "Name")
        self.assertEqual(tables["to-apply"].columns.get(slug="source").name, "Source")

    def test_job_forms_show_number_of_positions_field(self):
        self.client.force_login(self.user)

        to_apply_response = self.client.get(reverse("to_apply_create"))
        applied_response = self.client.get(reverse("applied_create"))

        self.assertEqual(to_apply_response.status_code, 200)
        self.assertEqual(applied_response.status_code, 200)
        self.assertContains(to_apply_response, 'name="number-of-positions"')
        self.assertContains(applied_response, 'name="number-of-positions"')
        self.assertContains(to_apply_response, 'name="name"')
        self.assertContains(to_apply_response, 'name="source"')

    def test_applied_form_uses_clear_labels_for_ambiguous_fields(self):
        DynamicColumn.objects.create(
            table=self.applied,
            name="Name",
            slug="name",
            data_type=DynamicColumn.DataType.TEXT,
            position=999,
        )
        DynamicColumn.objects.create(
            table=self.applied,
            name="Source",
            slug="source",
            data_type=DynamicColumn.DataType.TEXT,
            position=1000,
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("applied_create"))

        self.assertEqual(response.status_code, 200)
        form = response.context["form"]
        self.assertEqual(form.fields["name"].label, "Organization Name")
        self.assertEqual(form.fields["source"].label, "Application Source")
        self.assertEqual(form.fields["url"].label, "Job Posting URL")
        self.assertEqual(form.fields["source_row_key"].label, "Start from a To Apply entry")

    def test_applied_create_from_source_row_opens_applied_edit(self):
        source_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Applied ML Role",
                "url": "https://example.com/applied-ml-role",
                "type": "Job",
            },
        )
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("applied_create"),
            data={"source_row_key": str(source_row_key)},
        )

        self.assertRedirects(response, reverse("applied_edit", args=[source_row_key]))
        self.assertEqual(len(get_records(self.to_apply, self.user)), 0)
        applied_records = get_records(self.applied, self.user)
        self.assertEqual(len(applied_records), 1)
        self.assertEqual(applied_records[0].row_key, source_row_key)

    def test_to_apply_list_supports_tag_filter_and_sort_option(self):
        first_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Alpha Role",
                "url": "https://example.com/alpha",
                "type": "Job",
            },
        )
        second_row_key = save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Beta Role",
                "url": "https://example.com/beta",
                "type": "PhD",
            },
        )
        sync_application_tags(self.user, "to-apply", first_row_key, ["Priority"])
        sync_application_tags(self.user, "to-apply", second_row_key, ["Dummy"])
        self.client.force_login(self.user)

        response = self.client.get(reverse("to_apply_list"), {"sort": "tags", "dir": "asc", "tag": "priority"})

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Tags")
        self.assertContains(response, 'name="tag"')
        self.assertContains(response, "Alpha Role")
        self.assertNotContains(response, "Beta Role")

    def test_applied_dynamic_filter_ignores_blank_value_and_defaults_to_progress(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-01",
                "role-title": "Applied Role",
                "url": "https://example.com/applied-role",
                "type": "Job",
                "application-progress": "Applied",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-02",
                "role-title": "Interview Role",
                "url": "https://example.com/interview-role",
                "type": "Job",
                "application-progress": "Interview",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(
            reverse("applied_list"),
            {"filter_column": "application-progress", "filter_operator": "eq", "filter_value": ""},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["filter_column"], "application-progress")
        self.assertEqual(response.context["filter_operator"], "eq")
        self.assertContains(response, "Applied Role")
        self.assertContains(response, "Interview Role")
        self.assertIn("Rejected", response.context["filter_value_options"]["application-progress"])
        self.assertIn("Withdrawn", response.context["filter_value_options"]["application-progress"])

    def test_applied_dynamic_filter_uses_inferred_application_progress(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-01",
                "role-title": "Plain Applied Role",
                "url": "https://example.com/plain-applied",
                "type": "Job",
                "application-progress": "Applied",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-02",
                "role-title": "Rejected Applied Role",
                "url": "https://example.com/rejected-applied",
                "type": "Job",
                "result": "Rejected",
                "application-progress": "Applied",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(
            reverse("applied_list"),
            {"filter_column": "application-progress", "filter_operator": "eq", "filter_value": "Rejected"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Rejected Applied Role")
        self.assertNotContains(response, "Plain Applied Role")

    def test_applied_list_paginates_large_result_sets(self):
        for index in range(55):
            save_record(
                self.applied,
                self.user,
                {
                    "date": f"2026-05-{(index % 28) + 1:02d}",
                    "role-title": f"Paged Role {index:02d}",
                    "type": "Job",
                    "application-progress": "Applied",
                },
            )
        self.client.force_login(self.user)

        first_page = self.client.get(reverse("applied_list"))

        self.assertEqual(first_page.status_code, 200)
        self.assertEqual(first_page.context["record_count"], 55)
        self.assertEqual(first_page.context["page_record_count"], 25)
        self.assertEqual(len(first_page.context["table_rows"]), 25)
        self.assertTrue(first_page.context["pagination"]["has_next"])

        second_page = self.client.get(reverse("applied_list"), {"page": "2"})

        self.assertEqual(second_page.status_code, 200)
        self.assertEqual(second_page.context["record_count"], 55)
        self.assertEqual(second_page.context["page_record_count"], 25)
        self.assertEqual(len(second_page.context["table_rows"]), 25)
        self.assertTrue(second_page.context["pagination"]["has_previous"])
        self.assertTrue(second_page.context["pagination"]["has_next"])

        last_page = self.client.get(reverse("applied_list"), {"page": "3"})

        self.assertEqual(last_page.status_code, 200)
        self.assertEqual(last_page.context["page_record_count"], 5)
        self.assertEqual(len(last_page.context["table_rows"]), 5)
        self.assertTrue(last_page.context["pagination"]["has_previous"])
        self.assertFalse(last_page.context["pagination"]["has_next"])

    def test_table_filter_trigger_does_not_listen_to_bulk_checkboxes(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-01",
                "role-title": "Checkbox Role",
                "type": "Job",
                "application-progress": "Applied",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("applied_list"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "from:.table-control-text")
        self.assertContains(response, "from:.table-control-select")
        self.assertNotContains(response, "from:input")
        self.assertContains(response, "bulk-record-checkbox")

    def test_to_apply_dynamic_filter_defaults_to_virtual_application_progress(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Queued Role",
                "url": "https://example.com/queued-role",
                "type": "Job",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("to_apply_list"))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["filter_column"], "application-progress")
        self.assertEqual(response.context["filter_operator"], "eq")
        self.assertEqual(response.context["filter_value"], "")
        self.assertIn("To Apply", response.context["filter_value_options"]["application-progress"])
        self.assertContains(response, "Queued Role")

    def test_to_apply_dynamic_filter_supports_categorical_options_and_not_equals(self):
        DynamicColumn.objects.create(
            table=self.to_apply,
            name="Country",
            slug="country",
            data_type=DynamicColumn.DataType.TEXT,
            position=99,
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Germany Job",
                "url": "https://example.com/germany-job",
                "type": "Job",
                "country": "Germany",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "France PhD",
                "url": "https://example.com/france-phd",
                "type": "PhD",
                "country": "France",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(
            reverse("to_apply_list"),
            {"filter_column": "type", "filter_operator": "ne", "filter_value": "PhD"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Germany Job")
        self.assertNotContains(response, "France PhD")
        self.assertEqual(response.context["filter_column"], "type")
        self.assertIn("Germany", response.context["filter_value_options"]["country"])
        self.assertIn("France", response.context["filter_value_options"]["country"])

    def test_rag_returns_matching_records_from_user_tables(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "PhD Candidate in Computer Vision",
                "url": "https://example.com/phd-cv",
                "type": "PhD",
            },
        )

        with patch("tracker.rag.get_local_generation_pipeline", return_value=None):
            payload = chat_with_rag(self.user, "What are the PhD positions available with computer vision?")

        self.assertTrue(payload["default_prompts"])
        self.assertTrue(payload["sources"])
        self.assertIn("PhD Candidate in Computer Vision", payload["sources"][0]["title"])

    def test_active_applied_pipeline_counts_progress_without_rejected_or_withdrawn(self):
        for values in [
            {
                "date": "2026-04-29",
                "role-title": "Applied Role",
                "type": "Job",
                "application-progress": "Applied",
                "result": "Pending",
            },
            {
                "date": "2026-04-29",
                "role-title": "Interview Role",
                "type": "Job",
                "application-progress": "Interview",
                "result": "Pending",
            },
            {
                "date": "2026-04-29",
                "role-title": "Second Interview Role",
                "type": "Job",
                "application-progress": "Interview",
                "result": "",
            },
            {
                "date": "2026-04-29",
                "role-title": "Rejected Role",
                "type": "Job",
                "application-progress": "Interview",
                "result": "Rejected",
            },
            {
                "date": "2026-04-29",
                "role-title": "Withdrawn Role",
                "type": "Job",
                "application-progress": "Withdrawn",
                "result": "Withdrawn",
            },
        ]:
            save_record(self.applied, self.user, values)

        chart = build_active_applied_pipeline_chart(get_records(self.applied, self.user))

        self.assertEqual(chart["labels"], ["Applied", "Interview"])
        self.assertEqual(chart["values"], [1.0, 2.0])

    def test_chat_handler_creates_default_chat_models(self):
        available_models = ChatHandler.available_models()

        self.assertGreaterEqual(len(available_models), 2)
        self.assertTrue(ChatModelConfig.objects.filter(slug="google-flan-t5-small", enabled=True).exists())
        self.assertTrue(ChatModelConfig.objects.filter(slug="google-flan-t5-base", enabled=True).exists())

    def test_rag_handles_simple_greeting_without_search_error(self):
        with patch("tracker.rag.get_local_generation_pipeline", return_value=None):
            payload = chat_with_rag(self.user, "Hi")

        self.assertIn("I can help", payload["answer"])
        self.assertFalse(payload["sources"])

    def test_rag_uses_uploaded_cv_for_cover_letter_requests(self):
        with patch("tracker.rag.get_local_generation_pipeline", return_value=None):
            payload = chat_with_rag(
                self.user,
                "Write a simple cover letter for me.",
                uploaded_document_name="cv.txt",
                uploaded_document_text=(
                    "Jane Doe\n"
                    "Computer Vision Researcher\n"
                    "Worked on deep learning, Python, image segmentation, and academic collaboration.\n"
                ),
            )

        self.assertIn("Dear Hiring Team", payload["answer"])
        self.assertFalse(payload["sources"])

    def test_rag_chat_api_returns_answer_for_authenticated_user(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Research Engineer",
                "url": "https://example.com/research-engineer",
                "type": "Job",
            },
        )
        self.client.force_login(self.user)

        with patch("tracker.rag.get_local_generation_pipeline", return_value=None):
            response = self.client.post(
                reverse("api_rag_chat"),
                data=json.dumps({"query": "Show me research engineer roles", "model_slug": "google-flan-t5-small"}),
                content_type="application/json",
            )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertIn("answer", payload)
        self.assertEqual(payload["model_slug"], "google-flan-t5-small")
        self.assertTrue(payload["sources"])

    def test_rag_counts_filtered_phd_applied_roles(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "PhD Candidate in Vision",
                "url": "https://example.com/phd-vision",
                "type": "PhD",
                "application-progress": "Applied",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "ML Engineer",
                "url": "https://example.com/ml-engineer",
                "type": "Job",
                "application-progress": "Applied",
            },
        )

        with patch("tracker.rag.get_local_generation_pipeline", return_value=None):
            payload = chat_with_rag(self.user, "What is my PhD apply count?")

        self.assertIn("applied to 1", payload["answer"].lower())
        self.assertIn("phd", payload["answer"].lower())
        self.assertFalse(payload["sources"])

    def test_rag_counts_ai_reliance_queries_with_short_tokens_and_numbers(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Research Scientist",
                "url": "https://example.com/research-scientist",
                "type": "Job",
                "application-progress": "Applied",
                "ai-reliance": "100",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/vision-phd",
                "type": "PhD",
                "application-progress": "Applied",
                "ai-reliance": "100",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "ML Engineer",
                "url": "https://example.com/ml-engineer",
                "type": "Job",
                "application-progress": "Applied",
                "ai-reliance": "50",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Role with 50 in the URL",
                "url": "https://example.com/roles/50-in-url",
                "type": "Job",
                "application-progress": "Applied",
                "ai-reliance": "10",
            },
        )

        with patch("tracker.rag.get_local_generation_pipeline", return_value=None):
            full_ai_payload = chat_with_rag(self.user, "How many times did I use AI 100%")
            partial_ai_payload = chat_with_rag(self.user, "How many times did I use AI 50%")

        self.assertIn("2 applied roles", full_ai_payload["answer"])
        self.assertIn("AI reliance 100", full_ai_payload["answer"])
        self.assertFalse(full_ai_payload["sources"])
        self.assertIn("1 applied roles", partial_ai_payload["answer"])
        self.assertIn("AI reliance 50", partial_ai_payload["answer"])
        self.assertFalse(partial_ai_payload["sources"])

    def test_to_apply_date_sort_handles_blank_and_invalid_deadlines(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Early deadline",
                "url": "https://example.com/early",
                "type": "PhD",
                "deadline": "2026-05-01",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Invalid deadline",
                "url": "https://example.com/invalid",
                "type": "Job",
                "deadline": "not-a-date",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Missing deadline",
                "url": "https://example.com/missing",
                "type": "Internship",
                "deadline": "",
            },
        )

        view = ToApplyListView()
        view.request = SimpleNamespace(GET={"sort": "deadline", "dir": "asc"})

        records = view.sort_records_for_display(get_records(self.to_apply, self.user), self.to_apply)

        self.assertEqual(records[0].values["role-title"], "Early deadline")
        self.assertCountEqual(
            [records[-2].values["role-title"], records[-1].values["role-title"]],
            ["Invalid deadline", "Missing deadline"],
        )

    def test_to_apply_defaults_to_deadline_ascending_with_missing_last(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Later deadline",
                "url": "https://example.com/later",
                "type": "PhD",
                "deadline": "2026-05-10",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Missing deadline",
                "url": "https://example.com/missing",
                "type": "Job",
                "deadline": "",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Earlier deadline",
                "url": "https://example.com/earlier",
                "type": "Internship",
                "deadline": "2026-05-01",
            },
        )

        view = ToApplyListView()
        view.request = SimpleNamespace(GET={})

        records = view.sort_records_for_display(get_records(self.to_apply, self.user), self.to_apply)

        self.assertEqual(view.get_sort_field(self.to_apply), "deadline")
        self.assertEqual(view.get_sort_direction(self.to_apply), "asc")
        self.assertEqual(records[0].values["role-title"], "Earlier deadline")
        self.assertEqual(records[1].values["role-title"], "Later deadline")
        self.assertEqual(records[-1].values["role-title"], "Missing deadline")

    def test_applied_defaults_to_date_descending_with_missing_last(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-10",
                "role-title": "Latest application",
                "url": "https://example.com/latest",
                "type": "Job",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "",
                "role-title": "Missing date application",
                "url": "https://example.com/missing-date",
                "type": "PhD",
            },
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-01",
                "role-title": "Older application",
                "url": "https://example.com/older",
                "type": "Internship",
            },
        )

        view = ToApplyListView()
        view.request = SimpleNamespace(GET={})

        records = view.sort_records_for_display(get_records(self.applied, self.user), self.applied)

        self.assertEqual(view.get_sort_field(self.applied), "date")
        self.assertEqual(view.get_sort_direction(self.applied), "desc")
        self.assertEqual(records[0].values["role-title"], "Latest application")
        self.assertEqual(records[1].values["role-title"], "Older application")
        self.assertEqual(records[-1].values["role-title"], "Missing date application")

    def test_date_grouped_aggregations_stay_chronological_and_keep_unspecified_last(self):
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Middle",
                "url": "https://example.com/middle",
                "type": "PhD",
                "deadline": "2026-05-10",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Earliest",
                "url": "https://example.com/earliest",
                "type": "PhD",
                "deadline": "2026-05-01",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Latest duplicate",
                "url": "https://example.com/latest-duplicate",
                "type": "Job",
                "deadline": "2026-05-10",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "No deadline",
                "url": "https://example.com/no-deadline",
                "type": "Job",
                "deadline": "",
            },
        )

        labels, values = aggregate_records(
            get_records(self.to_apply, self.user),
            group_by_column="deadline",
            aggregation="Count Records",
            metric_column="",
            limit=10,
        )

        self.assertEqual(labels, ["2026-05-01", "2026-05-10", "Unspecified"])
        self.assertEqual(values, [1.0, 2.0, 1.0])

    def test_all_portal_jobs_page_lists_scraped_jobs_in_table(self):
        PortalJobSuggestion.objects.create(
            user=self.user,
            portal_table=self.portals,
            portal_row_key=uuid4(),
            portal_name="Example Research Portal",
            portal_url="https://example.com/careers",
            title="Computer Vision Scientist",
            url="https://example.com/careers/cv-scientist",
            summary="Research role focused on vision and machine learning.",
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("all_portal_jobs"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Computer Vision Scientist")
        self.assertContains(response, "Example Research Portal")

    def test_portals_list_fetch_jobs_link_opens_live_fetch_page(self):
        row_key = save_record(
            self.portals,
            self.user,
            {
                "url": "https://example.com/portal",
                "name": "Example Portal",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("portals_list"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, reverse("portal_fetch_live", args=[row_key]))

    def test_portal_fetch_live_page_renders_and_includes_stream_url(self):
        row_key = save_record(
            self.portals,
            self.user,
            {
                "url": "https://example.com/portal",
                "name": "Example Portal",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("portal_fetch_live", args=[row_key]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Live Portal Fetch")
        self.assertContains(response, reverse("api_portal_fetch_stream", args=[row_key]))

    def test_portal_fetch_stream_emits_job_and_done_events(self):
        row_key = save_record(
            self.portals,
            self.user,
            {
                "url": "https://example.com/portal",
                "name": "Example Portal",
            },
        )
        job = PortalJobSuggestion.objects.create(
            user=self.user,
            portal_table=self.portals,
            portal_row_key=row_key,
            portal_name="Example Portal",
            portal_url="https://example.com/portal",
            title="Computer Vision Scientist",
            url="https://example.com/portal/cv-scientist",
            summary="Research role focused on vision and machine learning.",
        )
        self.client.force_login(self.user)

        with patch("tracker.views.refresh_portal_jobs", return_value=[job]):
            response = self.client.get(reverse("api_portal_fetch_stream", args=[row_key]))
            body = b"".join(response.streaming_content).decode("utf-8", errors="ignore")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "text/event-stream")
        self.assertIn("event: job", body)
        self.assertIn("Computer Vision Scientist", body)
        self.assertIn("event: done", body)

    def test_portal_sweep_stream_emits_portal_done_and_done_events(self):
        row_key = save_record(
            self.portals,
            self.user,
            {
                "url": "https://example.com/portal",
                "name": "Example Portal",
            },
        )
        job = PortalJobSuggestion.objects.create(
            user=self.user,
            portal_table=self.portals,
            portal_row_key=row_key,
            portal_name="Example Portal",
            portal_url="https://example.com/portal",
            title="Computer Vision Scientist",
            url="https://example.com/portal/cv-scientist",
            summary="Research role focused on vision and machine learning.",
        )
        self.client.force_login(self.user)

        with patch("tracker.views.refresh_portal_jobs", return_value=[job]):
            response = self.client.get(reverse("api_portal_fetch_all_stream"))
            body = b"".join(response.streaming_content).decode("utf-8", errors="ignore")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "text/event-stream")
        self.assertIn("event: portal_done", body)
        self.assertIn("event: done", body)

    def test_url_inference_reads_jobposting_json_ld_without_raw_employment_type(self):
        html = """
        <html>
            <head>
                <meta property="og:site_name" content="Example Careers">
                <script type="application/ld+json">
                {
                    "@context": "https://schema.org",
                    "@graph": [
                        {
                            "@type": "JobPosting",
                            "title": "Computer Vision Research Engineer",
                            "employmentType": "FULL_TIME",
                            "validThrough": "2026-08-31T23:59:00Z",
                            "description": "<p>Build reliable perception systems.</p>",
                            "hiringOrganization": {"@type": "Organization", "name": "Example Robotics"},
                            "jobLocation": {
                                "@type": "Place",
                                "address": {"@type": "PostalAddress", "addressCountry": "DE"}
                            }
                        }
                    ]
                }
                </script>
            </head>
        </html>
        """

        with (
            patch("tracker.services.safe_get_url", return_value=SimpleNamespace(text=html)),
            patch("tracker.services.fetch_portal_html_in_browser") as browser_fetch,
        ):
            result = infer_fields_from_url_details("https://jobs.example.com/research-engineer")

        values = result["values"]
        self.assertTrue(result["scrape_succeeded"])
        self.assertEqual(result["source"], "page")
        self.assertEqual(values["role-title"], "Computer Vision Research Engineer")
        self.assertEqual(values["name"], "Example Robotics")
        self.assertEqual(values["source"], "Example Careers")
        self.assertEqual(values["type"], "Job")
        self.assertEqual(values["deadline"], "2026-08-31")
        self.assertEqual(values["country"], "Germany")
        browser_fetch.assert_not_called()

    def test_url_inference_api_only_returns_empty_field_suggestions(self):
        tables = ensure_user_workspace(self.user)
        self.client.force_login(self.user)

        with patch(
            "tracker.api.infer_fields_from_url_details",
            return_value={
                "values": {
                    "role-title": "Scraped Role",
                    "name": "Scraped Organization",
                    "source": "Scraped Source",
                    "type": "Job",
                    "country": "Germany",
                },
                "scrape_succeeded": True,
                "source": "page",
                "detail": "",
            },
        ):
            response = self.client.post(
                reverse("api_infer_record_from_url"),
                data=json.dumps(
                    {
                        "table_slug": tables["to-apply"].slug,
                        "url": "https://example.com/jobs/scraped-role",
                        "values": {
                            "role-title": "Already Typed",
                            "source": "",
                            "type": "",
                            "country": "",
                        },
                    }
                ),
                content_type="application/json",
            )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertNotIn("role-title", payload["values"])
        self.assertEqual(payload["values"]["source"], "Scraped Source")
        self.assertEqual(payload["values"]["type"], "Job")
        self.assertEqual(payload["values"]["country"], "Germany")
        self.assertTrue(payload["scrape_succeeded"])

    def test_url_inference_api_reports_when_no_fields_can_be_inferred(self):
        tables = ensure_user_workspace(self.user)
        self.client.force_login(self.user)

        with patch(
            "tracker.api.infer_fields_from_url_details",
            return_value={
                "values": {},
                "scrape_succeeded": False,
                "source": "url",
                "detail": "blocked",
            },
        ):
            response = self.client.post(
                reverse("api_infer_record_from_url"),
                data=json.dumps(
                    {
                        "table_slug": tables["to-apply"].slug,
                        "url": "https://example.com/jobs/no-details",
                        "values": {},
                    }
                ),
                content_type="application/json",
            )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(payload["values"], {})
        self.assertFalse(payload["scrape_succeeded"])
        self.assertIn("Could not read enough job details", payload["detail"])
        self.assertEqual(payload["error_detail"], "blocked")

    def test_scrape_jobs_from_portal_falls_back_to_selenium_after_403(self):
        blocked_error = requests.HTTPError("403 Client Error: Forbidden for url")
        blocked_error.response = SimpleNamespace(status_code=403)
        browser_html = """
        <html>
            <body>
                <a href="/careers/cv-scientist">Computer Vision Scientist</a>
            </body>
        </html>
        """

        with (
            patch("tracker.services.portal_fetch_mode", return_value="auto"),
            patch("tracker.services.safe_get_url", side_effect=blocked_error),
            patch("tracker.services.fetch_portal_html_in_browser", return_value=browser_html) as browser_fetch,
        ):
            jobs = scrape_jobs_from_portal("https://example.com/careers", ["vision"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "Computer Vision Scientist")
        self.assertEqual(jobs[0]["url"], "https://example.com/careers/cv-scientist")
        browser_fetch.assert_called_once()

    def test_scrape_jobs_from_portal_falls_back_to_selenium_after_ssl_handshake_failure(self):
        ssl_error = requests.exceptions.SSLError(
            "[SSL: SSLV3_ALERT_HANDSHAKE_FAILURE] sslv3 alert handshake failure (_ssl.c:1000)"
        )
        browser_html = """
        <html>
            <body>
                <a href="/stellen/research-assistant">Research Assistant Computer Vision</a>
            </body>
        </html>
        """

        with (
            patch("tracker.services.portal_fetch_mode", return_value="auto"),
            patch("tracker.services.safe_get_url", side_effect=ssl_error),
            patch("tracker.services.fetch_portal_html_in_browser", return_value=browser_html) as browser_fetch,
        ):
            jobs = scrape_jobs_from_portal("https://adb.zuv.uni-heidelberg.de/info/INFO_LS$.Startup", ["vision"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "Research Assistant Computer Vision")
        self.assertEqual(
            jobs[0]["url"],
            "https://adb.zuv.uni-heidelberg.de/stellen/research-assistant",
        )
        browser_fetch.assert_called_once()

    def test_scrape_jobs_from_portal_retries_with_selenium_for_javascript_pages(self):
        static_response = SimpleNamespace(text="<html><body><div id='app-root'></div></body></html>")
        browser_html = """
        <html>
            <body>
                <a href="/jobs/ml-researcher">Machine Learning Researcher</a>
            </body>
        </html>
        """

        with (
            patch("tracker.services.portal_fetch_mode", return_value="auto"),
            patch("tracker.services.safe_get_url", return_value=static_response),
            patch("tracker.services.fetch_portal_html_in_browser", return_value=browser_html) as browser_fetch,
        ):
            jobs = scrape_jobs_from_portal("https://example.com/portal", ["researcher"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "Machine Learning Researcher")
        self.assertEqual(jobs[0]["url"], "https://example.com/jobs/ml-researcher")
        browser_fetch.assert_called_once()

    def test_scrape_jobs_from_portal_filters_keywords_by_words_not_substrings(self):
        html = """
        <html>
            <body>
                <div>
                    <a href="/jobs/physics-chair">Research Chair in Physics</a>
                </div>
                <div>
                    <a href="/jobs/ai-researcher">AI Research Engineer</a>
                </div>
            </body>
        </html>
        """

        with patch("tracker.services.safe_get_url", return_value=SimpleNamespace(text=html)):
            jobs = scrape_jobs_from_portal("https://example.com/careers", ["AI"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "AI Research Engineer")
        self.assertEqual(jobs[0]["url"], "https://example.com/jobs/ai-researcher")

    def test_scrape_jobs_from_portal_matches_keywords_from_summary_text(self):
        html = """
        <html>
            <body>
                <div>
                    <a href="/jobs/open-position">Open Position</a>
                    <span>Computer-Vision Laboratory and robotics team</span>
                </div>
            </body>
        </html>
        """

        with patch("tracker.services.safe_get_url", return_value=SimpleNamespace(text=html)):
            jobs = scrape_jobs_from_portal("https://example.com/careers", ["computer vision"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "Open Position")
        self.assertIn("Computer-Vision Laboratory", jobs[0]["summary"])

    def test_scrape_jobs_from_portal_matches_plural_forms_of_keywords(self):
        html = """
        <html>
            <body>
                <div>
                    <a href="/jobs/applied-scientists">Applied Scientists</a>
                </div>
            </body>
        </html>
        """

        with patch("tracker.services.safe_get_url", return_value=SimpleNamespace(text=html)):
            jobs = scrape_jobs_from_portal("https://example.com/careers", ["scientist"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "Applied Scientists")

    def test_scrape_jobs_from_portal_supports_combined_keyword_entries(self):
        html = """
        <html>
            <body>
                <div>
                    <a href="/jobs/ml-engineer">ML Engineer</a>
                </div>
            </body>
        </html>
        """

        with patch("tracker.services.safe_get_url", return_value=SimpleNamespace(text=html)):
            jobs = scrape_jobs_from_portal("https://example.com/careers", ["AI/ML"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "ML Engineer")

    def test_portal_keyword_matching_understands_common_acronyms(self):
        self.assertTrue(matches_any_keyword("Artificial Intelligence Research Engineer", ["AI"]))
        self.assertTrue(matches_any_keyword("Machine Learning Scientist", ["ML"]))
        self.assertTrue(matches_any_keyword("Computer Vision Researcher", ["CV"]))
        self.assertFalse(matches_any_keyword("Research Chair in Physics", ["AI"]))

    def test_portal_keyword_matching_allows_non_adjacent_phrase_tokens(self):
        html = """
        <html>
            <body>
                <div>
                    <a href="/jobs/vision-lab">Computer Perception and Vision Research Engineer</a>
                </div>
            </body>
        </html>
        """

        jobs = extract_jobs_from_portal_html("https://example.com/careers", html, ["computer vision"])

        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["title"], "Computer Perception and Vision Research Engineer")

    def test_dashboard_recent_submissions_show_submission_date(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-02",
                "role-title": "Vision PhD",
                "url": "https://example.com/vision-phd",
                "type": "PhD",
                "application-progress": "Applied",
                "result": "Rejected",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("dashboard"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Submitted 02/05/2026")
        self.assertContains(response, "Rejected")

    def test_export_applications_downloads_user_tables_as_workbook(self):
        follow_up = DynamicTable.objects.create(
            owner=self.user,
            name="Follow Up",
            slug="follow-up",
        )
        DynamicColumn.objects.create(
            table=follow_up,
            name="Status Note",
            slug="status-note",
            data_type=DynamicColumn.DataType.TEXT,
            position=1,
        )
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-29",
                "role-title": "Vision PhD",
                "url": "https://example.com/vision-phd",
                "name": "Example University",
                "type": "PhD",
                "application-progress": "Applied",
            },
        )
        save_record(
            self.portals,
            self.user,
            {
                "url": "https://example.com/portal",
                "name": "Example Portal",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "url": "https://example.com/to-apply",
                "deadline": "2026-05-10",
                "type": "Job",
            },
        )
        save_record(
            follow_up,
            self.user,
            {
                "status-note": "Waiting for recruiter reply",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("export_applications"))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response["Content-Type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.assertIn('filename="Applications.xlsx"', response["Content-Disposition"])

        workbook = load_workbook(BytesIO(response.content))
        self.assertEqual(workbook.sheetnames[:3], ["Applied", "Portals", "To Apply"])
        self.assertIn("Follow Up", workbook.sheetnames)

        applied_sheet = workbook["Applied"]
        self.assertEqual(applied_sheet["B2"].value, "Vision PhD")
        self.assertEqual(applied_sheet["B2"].hyperlink.target, "https://example.com/vision-phd")
        self.assertEqual(applied_sheet["A2"].value.strftime("%d/%m/%Y"), "29/04/2026")

        follow_up_sheet = workbook["Follow Up"]
        self.assertEqual(follow_up_sheet["A1"].value, "Status Note")
        self.assertEqual(follow_up_sheet["A2"].value, "Waiting for recruiter reply")

    def test_manage_tables_can_append_uploaded_workbook(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-01",
                "role-title": "Existing application",
                "url": "https://example.com/existing",
                "type": "Job",
            },
        )
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("manage_tables"),
            data={
                "import_workbook": "1",
                "import_mode": "append",
                "workbook": self.build_uploaded_workbook(),
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(len(get_records(self.applied, self.user)), 2)
        self.assertEqual(len(get_records(self.to_apply, self.user)), 1)
        self.assertEqual(len(get_records(self.portals, self.user)), 1)

    def test_manage_tables_can_replace_with_uploaded_workbook(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-04-01",
                "role-title": "Existing application",
                "url": "https://example.com/existing",
                "type": "Job",
            },
        )
        save_record(
            self.to_apply,
            self.user,
            {
                "role-title": "Existing to apply",
                "url": "https://example.com/existing-to-apply",
                "type": "PhD",
                "deadline": "2026-04-15",
            },
        )
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("manage_tables"),
            data={
                "import_workbook": "1",
                "import_mode": "replace",
                "workbook": self.build_uploaded_workbook(),
            },
        )

        self.assertEqual(response.status_code, 302)
        applied_records = get_records(self.applied, self.user)
        to_apply_records = get_records(self.to_apply, self.user)
        portal_records = get_records(self.portals, self.user)
        self.assertEqual(len(applied_records), 1)
        self.assertEqual(len(to_apply_records), 1)
        self.assertEqual(len(portal_records), 1)
        self.assertEqual(applied_records[0].values["url"], "https://example.com/imported-applied")
        self.assertEqual(to_apply_records[0].values["url"], "https://example.com/imported-to-apply")


class NotificationWorkflowTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="notifier",
            email="notifier@example.com",
            password="TestPassword123!",
        )
        self.user.notification_email = "notifier@example.com"
        self.user.notify_enabled = True
        self.user.auto_fetch_portals = True
        self.user.save()

    def test_run_notifications_sends_daily_summary_even_without_new_items(self):
        smtp_context = DummySMTPContext()
        self.user.last_notification_error = "old smtp error"
        self.user.last_notification_error_at = timezone.make_aware(datetime(2026, 5, 18, 9, 30))
        self.user.save(update_fields=["last_notification_error", "last_notification_error_at"])

        with (
            patch("tracker.notifications.SMTPHandler.is_configured", return_value=True),
            patch("tracker.notifications.SMTPHandler.from_env", return_value=smtp_context),
            patch("tracker.notifications.refresh_portals_for_user", return_value=[]),
            patch("tracker.notifications.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 5, 20, 9, 0))),
        ):
            processed = run_notifications(force=False, dry_run=False)

        self.user.refresh_from_db()
        self.assertEqual(processed, 1)
        self.assertEqual(len(smtp_context.sent_messages), 1)
        recipient, subject, body = smtp_context.sent_messages[0]
        self.assertEqual(recipient, "notifier@example.com")
        self.assertIn("Job Tracker Update", subject)
        self.assertIn("No upcoming deadlines were found", body)
        self.assertIn("No new portal matches were found", body)
        self.assertIsNotNone(self.user.last_notification_sent_at)
        self.assertEqual(self.user.last_notification_error, "")
        self.assertIsNone(self.user.last_notification_error_at)

    def test_run_notifications_continues_after_one_recipient_fails(self):
        blocked_user = get_user_model().objects.create_user(
            username="blocked",
            email="blocked@example.com",
            password="TestPassword123!",
        )
        blocked_user.notification_email = "blocked@example.com"
        blocked_user.notify_enabled = True
        blocked_user.auto_fetch_portals = False
        blocked_user.save()
        self.user.auto_fetch_portals = False
        self.user.save(update_fields=["auto_fetch_portals"])
        smtp_context = DummySelectiveFailSMTPContext("blocked@example.com")

        with (
            patch("tracker.notifications.SMTPHandler.is_configured", return_value=True),
            patch("tracker.notifications.SMTPHandler.from_env", return_value=smtp_context),
            patch("tracker.notifications.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 5, 20, 9, 0))),
        ):
            processed = run_notifications(force=False, dry_run=False)

        self.user.refresh_from_db()
        blocked_user.refresh_from_db()
        self.assertEqual(processed, 1)
        self.assertEqual(len(smtp_context.sent_messages), 1)
        self.assertEqual(smtp_context.sent_messages[0][0], "notifier@example.com")
        self.assertIsNotNone(self.user.last_notification_sent_at)
        self.assertIsNone(blocked_user.last_notification_sent_at)
        self.assertEqual(self.user.last_notification_error, "")
        self.assertIsNone(self.user.last_notification_error_at)
        self.assertIn("recipient refused", blocked_user.last_notification_error)
        self.assertIsNotNone(blocked_user.last_notification_error_at)

    def test_send_notification_to_user_runs_portal_refresh_once_per_send(self):
        smtp_context = DummySMTPContext()
        self.user.last_notification_sent_at = None
        self.user.save(update_fields=["last_notification_sent_at"])

        with (
            patch("tracker.notifications.SMTPHandler.from_env", return_value=smtp_context),
            patch("tracker.notifications.refresh_portals_for_user", return_value=[]) as refresh_mock,
        ):
            result = send_notification_to_user(
                self.user,
                force=True,
                dry_run=False,
                smtp_handler=None,
                include_portal_refresh=True,
            )

        self.assertTrue(result["sent"])
        self.assertEqual(refresh_mock.call_count, 1)
        self.assertEqual(len(smtp_context.sent_messages), 1)

    def test_send_notification_to_user_does_not_clear_automatic_error_state(self):
        smtp_context = DummySMTPContext()
        error_at = timezone.make_aware(datetime(2026, 5, 19, 8, 30))
        self.user.last_notification_error = "scheduled smtp timeout"
        self.user.last_notification_error_at = error_at
        self.user.save(update_fields=["last_notification_error", "last_notification_error_at"])

        with patch("tracker.notifications.SMTPHandler.from_env", return_value=smtp_context):
            result = send_notification_to_user(
                self.user,
                force=True,
                dry_run=False,
                smtp_handler=None,
                include_portal_refresh=False,
            )

        self.user.refresh_from_db()
        self.assertTrue(result["sent"])
        self.assertEqual(self.user.last_notification_error, "scheduled smtp timeout")
        self.assertEqual(self.user.last_notification_error_at, error_at)

    def test_notification_due_users_respects_configured_daily_time(self):
        early = timezone.make_aware(datetime(2026, 5, 19, 7, 59))
        due = timezone.make_aware(datetime(2026, 5, 19, 8, 0))
        self.user.notify_time = time(hour=8, minute=0)
        self.user.last_notification_sent_at = None
        self.user.save(update_fields=["notify_time", "last_notification_sent_at"])

        self.assertEqual(notification_due_users(early), [])
        self.assertEqual(notification_due_users(due), [self.user])

        self.user.last_notification_sent_at = due
        self.user.save(update_fields=["last_notification_sent_at"])

        self.assertEqual(notification_due_users(due), [])

    def test_scheduler_starts_for_app_runtime_but_not_management_commands(self):
        self.assertTrue(should_start_notification_scheduler(["gunicorn", "config.wsgi"]))
        with patch.dict("os.environ", {"RUN_MAIN": "true"}):
            self.assertTrue(should_start_notification_scheduler(["manage.py", "runserver"]))
        with patch.dict("os.environ", {}, clear=True):
            self.assertFalse(should_start_notification_scheduler(["manage.py", "runserver"]))
            self.assertFalse(should_start_notification_scheduler(["manage.py", "migrate"]))
            self.assertFalse(should_start_notification_scheduler(["manage.py", "test"]))
            self.assertFalse(should_start_notification_scheduler(["run.py", "notify"]))
        with patch.dict("os.environ", {"DISABLE_NOTIFICATION_SCHEDULER": "1"}):
            self.assertFalse(should_start_notification_scheduler(["gunicorn", "config.wsgi"]))

    def test_account_settings_send_notification_now_uses_manual_send_path(self):
        self.client.force_login(self.user)

        with (
            patch("tracker.views.SMTPHandler.is_configured", return_value=True),
            patch("tracker.views.send_notification_to_user", return_value={"sent": True, "recipient": "notifier@example.com"}) as send_mock,
        ):
            response = self.client.post(
                reverse("account_settings"),
                data={
                    "notification_email": "notifier@example.com",
                    "notify_enabled": "on",
                    "notify_time": "08:00",
                    "auto_fetch_portals": "on",
                    "deadline_notification_window_days": "7",
                    "chat_history_retention_days": "1",
                    "send_notification_now": "1",
                },
            )

        self.assertEqual(response.status_code, 302)
        send_mock.assert_called_once()
        _, kwargs = send_mock.call_args
        self.assertEqual(kwargs["force"], True)
        self.assertEqual(kwargs["dry_run"], False)

    def test_account_settings_shows_last_automatic_notification_error(self):
        self.user.last_notification_error = "SMTP handshake failed"
        self.user.last_notification_error_at = timezone.make_aware(datetime(2026, 5, 20, 10, 45))
        self.user.save(update_fields=["last_notification_error", "last_notification_error_at"])
        self.client.force_login(self.user)

        response = self.client.get(reverse("account_settings"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Last automatic error")
        self.assertContains(response, "SMTP handshake failed")

    def test_account_settings_time_change_allows_new_same_day_reminder(self):
        sent_at = timezone.make_aware(datetime(2026, 5, 20, 11, 5))
        self.user.notify_time = time(hour=8, minute=0)
        self.user.last_notification_sent_at = sent_at
        self.user.save(update_fields=["notify_time", "last_notification_sent_at"])
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("account_settings"),
            data={
                "notification_email": "notifier@example.com",
                "notify_enabled": "on",
                "notify_time": "12:15",
                "auto_fetch_portals": "on",
                "deadline_notification_window_days": "7",
                "chat_history_retention_days": "1",
            },
        )

        self.assertEqual(response.status_code, 302)
        self.user.refresh_from_db()
        self.assertEqual(self.user.notify_time, time(hour=12, minute=15))
        self.assertIsNone(self.user.last_notification_sent_at)
