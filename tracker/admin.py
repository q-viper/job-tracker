from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as DjangoUserAdmin
from datetime import timedelta

from .models import ChatMessage, ChatModelConfig, DynamicColumn, DynamicData, DynamicTable, PortalJobSuggestion, User


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
