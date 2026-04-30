import json
from types import SimpleNamespace
from unittest.mock import patch
from uuid import uuid4

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from .chat_handler import ChatHandler
from .models import ChatModelConfig, DynamicColumn, DynamicTable, PortalJobSuggestion
from .rag import chat_with_rag
from .services import get_records, move_to_applied, save_record
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
