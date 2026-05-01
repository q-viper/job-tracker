from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as DjangoUserAdmin
from datetime import timedelta

from .models import (
    ApplicationHistoryEvent,
    ApplicationTag,
    ApplicationTagAssignment,
    ChatMessage,
    ChatModelConfig,
    DashboardBoard,
    DashboardBoardItem,
    DashboardPanelPreference,
    DynamicColumn,
    DynamicData,
    DynamicTable,
    PortalJobSuggestion,
    User,
)


class DynamicColumnInline(admin.TabularInline):
    model = DynamicColumn
    extra = 0


@admin.register(User)
class UserAdmin(DjangoUserAdmin):
    fieldsets = DjangoUserAdmin.fieldsets + (
        (
            "Profile",
            {
                "fields": (
                    "display_name",
                    "notification_email",
                    "notify_enabled",
                    "notify_time",
                    "auto_fetch_portals",
                    "deadline_notification_window_days",
                    "chat_history_retention_days",
                    "email_verified_at",
                    "verification_email_sent_at",
                    "last_notification_sent_at",
                )
            },
        ),
    )
    list_display = (
        "username",
        "email",
        "notification_email",
        "display_name",
        "notify_enabled",
        "auto_fetch_portals",
        "chat_history_retention_days",
        "email_verified_at",
        "is_staff",
        "is_active",
    )


@admin.register(DynamicTable)
class DynamicTableAdmin(admin.ModelAdmin):
    list_display = ("name", "owner", "slug", "display_order", "show_in_navigation")
    list_filter = ("show_in_navigation", "owner")
    search_fields = ("name", "slug", "owner__username")
    inlines = [DynamicColumnInline]


@admin.register(DynamicColumn)
class DynamicColumnAdmin(admin.ModelAdmin):
    list_display = (
        "name",
        "table",
        "slug",
        "data_type",
        "position",
        "is_required",
        "is_visible",
    )
    list_filter = ("table", "data_type", "is_required", "is_visible")
    search_fields = ("name", "slug", "table__name", "table__owner__username")


@admin.register(DynamicData)
class DynamicDataAdmin(admin.ModelAdmin):
    list_display = ("table", "column", "user", "row_key", "short_value", "updated_at")
    list_filter = ("table", "column", "user")
    search_fields = ("value", "row_key", "table__name", "user__username")

    @staticmethod
    def short_value(obj: DynamicData) -> str:
        return obj.value[:60]


@admin.register(PortalJobSuggestion)
class PortalJobSuggestionAdmin(admin.ModelAdmin):
    list_display = (
        "title",
        "user",
        "portal_name",
        "added_to_apply",
        "fetched_at",
    )
    list_filter = ("user", "added_to_apply", "portal_table")
    search_fields = ("title", "url", "portal_name", "user__username")


@admin.register(ApplicationTag)
class ApplicationTagAdmin(admin.ModelAdmin):
    list_display = ("name", "user", "slug", "updated_at")
    list_filter = ("user",)
    search_fields = ("name", "slug", "user__username")


@admin.register(ApplicationTagAssignment)
class ApplicationTagAssignmentAdmin(admin.ModelAdmin):
    list_display = ("tag", "user", "table_slug", "record_row_key", "created_at")
    list_filter = ("user", "table_slug", "tag")
    search_fields = ("tag__name", "user__username", "record_row_key")


@admin.register(ApplicationHistoryEvent)
class ApplicationHistoryEventAdmin(admin.ModelAdmin):
    list_display = ("user", "table_slug", "record_row_key", "event_type", "from_stage", "to_stage", "occurred_at")
    list_filter = ("user", "table_slug", "event_type")
    search_fields = ("user__username", "record_row_key", "from_stage", "to_stage", "notes")


@admin.register(DashboardPanelPreference)
class DashboardPanelPreferenceAdmin(admin.ModelAdmin):
    list_display = ("panel_key", "user", "zone", "sort_order", "updated_at")
    list_filter = ("zone", "user")
    search_fields = ("panel_key", "user__username")


@admin.register(DashboardBoard)
class DashboardBoardAdmin(admin.ModelAdmin):
    list_display = ("name", "user", "show_on_homepage", "autoplay_enabled", "sort_order", "updated_at")
    list_filter = ("show_on_homepage", "autoplay_enabled", "user")
    search_fields = ("name", "description", "user__username")


@admin.register(DashboardBoardItem)
class DashboardBoardItemAdmin(admin.ModelAdmin):
    list_display = ("board", "widget_row_key", "card_size", "chart_height", "sort_order", "updated_at")
    list_filter = ("card_size", "chart_height", "board__user")
    search_fields = ("board__name", "board__user__username", "widget_row_key")


@admin.action(description="Purge expired chat history for selected users")
def purge_expired_chat_history(modeladmin, request, queryset):
    from django.utils import timezone

    deleted = 0
    for user in queryset:
        cutoff = timezone.now() - timedelta(days=max(user.chat_history_retention_days, 1))
        deleted += ChatMessage.objects.filter(user=user, created_at__lt=cutoff).delete()[0]
    modeladmin.message_user(request, f"Deleted {deleted} expired chat messages.")


UserAdmin.actions = [purge_expired_chat_history]


@admin.register(ChatMessage)
class ChatMessageAdmin(admin.ModelAdmin):
    list_display = ("user", "role", "short_content", "attachment_name", "created_at")
    list_filter = ("role", "user")
    search_fields = ("content", "attachment_name", "user__username")
    actions = ("delete_selected",)

    @staticmethod
    def short_content(obj: ChatMessage) -> str:
        return obj.content[:80]


@admin.register(ChatModelConfig)
class ChatModelConfigAdmin(admin.ModelAdmin):
    list_display = ("name", "slug", "mode", "model_name", "enabled", "is_default", "updated_at")
    list_filter = ("mode", "enabled", "is_default")
    search_fields = ("name", "slug", "model_name", "notes")
