from __future__ import annotations

import uuid
from datetime import time

from django.conf import settings
from django.contrib.auth.models import AbstractUser
from django.core.exceptions import ValidationError
from django.db import models


class User(AbstractUser):
    display_name = models.CharField(max_length=120, blank=True)
    notification_email = models.EmailField(blank=True)
    notify_enabled = models.BooleanField(default=True)
    notify_time = models.TimeField(default=time(hour=8, minute=0))
    auto_fetch_portals = models.BooleanField(default=True)
    deadline_notification_window_days = models.PositiveSmallIntegerField(default=7)
    chat_history_retention_days = models.PositiveSmallIntegerField(default=1)
    email_verified_at = models.DateTimeField(null=True, blank=True)
    verification_email_sent_at = models.DateTimeField(null=True, blank=True)
    last_notification_sent_at = models.DateTimeField(null=True, blank=True)

    def __str__(self) -> str:
        return self.display_name or self.get_full_name() or self.username


class DynamicTable(models.Model):
    owner = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="tables",
    )
    name = models.CharField(max_length=120)
    slug = models.SlugField(max_length=120)
    description = models.TextField(blank=True)
    display_order = models.PositiveIntegerField(default=0)
    show_in_navigation = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("display_order", "name")
        constraints = [
            models.UniqueConstraint(
                fields=("owner", "slug"),
                name="unique_table_slug_per_owner",
            ),
        ]

    def __str__(self) -> str:
        return f"{self.owner.username} - {self.name}"


class DynamicColumn(models.Model):
    class DataType(models.TextChoices):
        TEXT = "text", "Text"
        LONG_TEXT = "long_text", "Long text"
        URL = "url", "URL"
        DATE = "date", "Date"
        NUMBER = "number", "Number"
        STATUS = "status", "Status"

    table = models.ForeignKey(
        DynamicTable,
        on_delete=models.CASCADE,
        related_name="columns",
    )
    name = models.CharField(max_length=120)
    slug = models.SlugField(max_length=120)
    data_type = models.CharField(
        max_length=20,
        choices=DataType.choices,
        default=DataType.TEXT,
    )
    position = models.PositiveIntegerField(default=0)
    is_required = models.BooleanField(default=False)
    is_visible = models.BooleanField(default=True)
    placeholder = models.CharField(max_length=150, blank=True)
    help_text = models.CharField(max_length=255, blank=True)
    choices_text = models.TextField(
        blank=True,
        help_text="Comma-separated options for status-like fields.",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("position", "id")
        constraints = [
            models.UniqueConstraint(
                fields=("table", "slug"),
                name="unique_column_slug_per_table",
            ),
        ]

    @property
    def choices_list(self) -> list[str]:
        return [
            item.strip()
            for item in self.choices_text.replace("\n", ",").split(",")
            if item.strip()
        ]

    def __str__(self) -> str:
        return f"{self.table.name} -> {self.name}"


class DynamicData(models.Model):
    table = models.ForeignKey(
        DynamicTable,
        on_delete=models.CASCADE,
        related_name="data_entries",
    )
    column = models.ForeignKey(
        DynamicColumn,
        on_delete=models.CASCADE,
        related_name="data_entries",
    )
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="data_entries",
    )
    row_key = models.UUIDField(default=uuid.uuid4, db_index=True)
    value = models.TextField(blank=True)
    metadata = models.JSONField(default=dict, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("table_id", "row_key", "column_id")
        constraints = [
            models.UniqueConstraint(
                fields=("table", "column", "user", "row_key"),
                name="unique_data_cell_per_row",
            ),
        ]
        indexes = [
            models.Index(fields=("table", "user", "row_key")),
        ]

    def clean(self) -> None:
        if self.column_id and self.table_id and self.column.table_id != self.table_id:
            raise ValidationError("Column must belong to the selected table.")
        if self.table_id and self.user_id and self.table.owner_id != self.user_id:
            raise ValidationError("Data must belong to the owner of the selected table.")

    def __str__(self) -> str:
        return f"{self.table.slug}:{self.row_key}:{self.column.slug}"


class PortalJobSuggestion(models.Model):
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="portal_job_suggestions",
    )
    portal_table = models.ForeignKey(
        DynamicTable,
        on_delete=models.CASCADE,
        related_name="portal_job_suggestions",
    )
    portal_row_key = models.UUIDField()
    portal_name = models.CharField(max_length=255, blank=True)
    portal_url = models.URLField(max_length=1000)
    title = models.CharField(max_length=255)
    url = models.URLField(max_length=1000)
    summary = models.TextField(blank=True)
    added_to_apply = models.BooleanField(default=False)
    fetched_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ("-fetched_at", "title")
        constraints = [
            models.UniqueConstraint(
                fields=("user", "portal_row_key", "url"),
                name="unique_portal_job_per_user",
            ),
        ]

    def clean(self) -> None:
        if self.portal_table_id and self.user_id and self.portal_table.owner_id != self.user_id:
            raise ValidationError("Portal jobs must belong to the portal owner.")

    def __str__(self) -> str:
        return f"{self.portal_name or self.portal_url} -> {self.title}"


class ChatMessage(models.Model):
    class Role(models.TextChoices):
        USER = "user", "User"
        ASSISTANT = "assistant", "Assistant"

    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="chat_messages",
    )
    role = models.CharField(max_length=20, choices=Role.choices)
    content = models.TextField()
    attachment_name = models.CharField(max_length=255, blank=True)
    attachment_text = models.TextField(blank=True)
    metadata = models.JSONField(default=dict, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ("created_at", "id")
        indexes = [
            models.Index(fields=("user", "created_at")),
        ]

    def __str__(self) -> str:
        return f"{self.user.username} [{self.role}] {self.created_at:%Y-%m-%d %H:%M}"


class ChatModelConfig(models.Model):
    class Mode(models.TextChoices):
        LIST = "list", "List / Retrieval"
        CHAT = "chat", "Chat"
        RAG = "rag", "RAG"

    name = models.CharField(max_length=120)
    slug = models.SlugField(max_length=120, unique=True)
    mode = models.CharField(max_length=20, choices=Mode.choices, default=Mode.LIST)
    model_name = models.CharField(max_length=255, blank=True)
    enabled = models.BooleanField(default=True)
    is_default = models.BooleanField(default=False)
    notes = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("-is_default", "name")

    def __str__(self) -> str:
        return self.name
