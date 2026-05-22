from __future__ import annotations

import uuid
from datetime import time

from django.conf import settings
from django.contrib.auth.models import AbstractUser
from django.core.exceptions import ValidationError
from django.db import models
from django.utils import timezone


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
    last_notification_error = models.TextField(blank=True)
    last_notification_error_at = models.DateTimeField(null=True, blank=True)

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


class ApplicationTag(models.Model):
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="application_tags",
    )
    name = models.CharField(max_length=80)
    slug = models.SlugField(max_length=80)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("name", "id")
        constraints = [
            models.UniqueConstraint(
                fields=("user", "slug"),
                name="unique_application_tag_slug_per_user",
            ),
        ]

    def __str__(self) -> str:
        return f"{self.user.username} -> {self.name}"


class ApplicationTagAssignment(models.Model):
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="application_tag_assignments",
    )
    tag = models.ForeignKey(
        ApplicationTag,
        on_delete=models.CASCADE,
        related_name="assignments",
    )
    table_slug = models.SlugField(max_length=120)
    record_row_key = models.UUIDField(db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ("tag__name", "id")
        constraints = [
            models.UniqueConstraint(
                fields=("user", "tag", "table_slug", "record_row_key"),
                name="unique_application_tag_assignment",
            ),
        ]
        indexes = [
            models.Index(fields=("user", "table_slug", "record_row_key")),
        ]

    def clean(self) -> None:
        if self.tag_id and self.user_id and self.tag.user_id != self.user_id:
            raise ValidationError("Tag assignments must belong to the tag owner.")

    def __str__(self) -> str:
        return f"{self.tag.name} -> {self.table_slug}:{self.record_row_key}"


class ApplicationHistoryEvent(models.Model):
    class EventType(models.TextChoices):
        ADDED_TO_APPLY = "added_to_apply", "Added to To Apply"
        APPLIED = "applied", "Applied"
        STAGE_CHANGE = "stage_change", "Stage Change"

    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="application_history_events",
    )
    table_slug = models.SlugField(max_length=120)
    record_row_key = models.UUIDField(db_index=True)
    event_type = models.CharField(max_length=40, choices=EventType.choices)
    from_stage = models.CharField(max_length=120, blank=True)
    to_stage = models.CharField(max_length=120, blank=True)
    occurred_at = models.DateTimeField(default=timezone.now)
    notes = models.CharField(max_length=255, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ("occurred_at", "id")
        indexes = [
            models.Index(fields=("user", "table_slug", "record_row_key", "occurred_at")),
        ]

    def __str__(self) -> str:
        return f"{self.table_slug}:{self.record_row_key} {self.event_type}"


class DashboardPanelPreference(models.Model):
    class Zone(models.TextChoices):
        LEFT = "left", "Left rail"
        MAIN = "main", "Main column"
        RIGHT = "right", "Right rail"
        HIDDEN = "hidden", "Hidden"

    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="dashboard_panel_preferences",
    )
    panel_key = models.SlugField(max_length=80)
    zone = models.CharField(max_length=20, choices=Zone.choices, default=Zone.MAIN)
    sort_order = models.PositiveIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("zone", "sort_order", "panel_key")
        constraints = [
            models.UniqueConstraint(
                fields=("user", "panel_key"),
                name="unique_dashboard_panel_preference_per_user",
            ),
        ]

    def __str__(self) -> str:
        return f"{self.user.username} -> {self.panel_key} ({self.zone})"


class DashboardBoard(models.Model):
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="dashboard_boards",
    )
    name = models.CharField(max_length=120)
    description = models.TextField(blank=True)
    show_on_homepage = models.BooleanField(default=True)
    autoplay_enabled = models.BooleanField(default=False)
    autoplay_interval_seconds = models.PositiveSmallIntegerField(default=8)
    sort_order = models.PositiveIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("sort_order", "name", "id")

    def clean(self) -> None:
        if self.autoplay_interval_seconds < 3:
            raise ValidationError("Autoplay interval should be at least 3 seconds.")

    def __str__(self) -> str:
        return f"{self.user.username} -> {self.name}"


class DashboardBoardItem(models.Model):
    class CardSize(models.TextChoices):
        COMPACT = "Compact", "Compact"
        STANDARD = "Standard", "Standard"
        WIDE = "Wide", "Wide"
        FULL = "Full", "Full"

    class ChartHeight(models.TextChoices):
        SHORT = "Short", "Short"
        MEDIUM = "Medium", "Medium"
        TALL = "Tall", "Tall"
        XL = "XL", "XL"

    board = models.ForeignKey(
        DashboardBoard,
        on_delete=models.CASCADE,
        related_name="items",
    )
    widget_row_key = models.UUIDField(db_index=True)
    sort_order = models.PositiveIntegerField(default=0)
    card_size = models.CharField(
        max_length=20,
        choices=CardSize.choices,
        default=CardSize.STANDARD,
    )
    chart_height = models.CharField(
        max_length=20,
        choices=ChartHeight.choices,
        default=ChartHeight.MEDIUM,
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("sort_order", "id")
        constraints = [
            models.UniqueConstraint(
                fields=("board", "widget_row_key"),
                name="unique_dashboard_widget_per_board",
            ),
        ]
        indexes = [
            models.Index(fields=("board", "sort_order")),
        ]

    def __str__(self) -> str:
        return f"{self.board.name} -> {self.widget_row_key}"


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
