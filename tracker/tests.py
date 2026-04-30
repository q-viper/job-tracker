import json
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import patch
from uuid import uuid4

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse
from openpyxl import load_workbook

from .chat_handler import ChatHandler
from .models import ChatModelConfig, DynamicColumn, DynamicTable, PortalJobSuggestion
from .rag import chat_with_rag
from .services import aggregate_records, get_records, move_to_applied, save_record
from .views import ToApplyListView


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
        self.assertEqual(records[-1].values["role-title"], "Missing deadline")

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

    def test_dashboard_recent_submissions_show_submission_date(self):
        save_record(
            self.applied,
            self.user,
            {
                "date": "2026-05-02",
                "role-title": "Vision PhD",
                "url": "https://example.com/vision-phd",
                "type": "PhD",
                "application-progress": "Interview",
            },
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("dashboard"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Submitted 02 May 2026")

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
