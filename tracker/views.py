from __future__ import annotations

from io import BytesIO
import json
from pathlib import Path
from tempfile import NamedTemporaryFile

from django.conf import settings
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.decorators import login_required
from django.contrib.auth.mixins import LoginRequiredMixin
from django.contrib.auth.views import LoginView
from django.http import Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme
from django.utils import timezone
from django.views import View
from django.views.generic import FormView, TemplateView

from .forms import (
    AccountSettingsForm,
    AppliedCreateForm,
    AppliedUpdateForm,
    ColumnDefinitionForm,
    DashboardBoardForm,
    DynamicRecordForm,
    EmailAwareAuthenticationForm,
    SignupForm,
    TableDefinitionForm,
    ToApplyForm,
    WorkbookImportForm,
)
from .email_verification import send_verification_email, verify_user_from_token
from .models import (
    ApplicationTag,
    DashboardBoard,
    DashboardBoardItem,
    DashboardPanelPreference,
    DynamicTable,
    PortalJobSuggestion,
)
from .services import (
    add_application_tags,
    add_dashboard_widget_to_board,
    add_portal_job_to_to_apply,
    build_export_workbook,
    build_application_history_timeline,
    build_dashboard_board_widgets,
    build_dashboard,
    dashboard_panel_catalog,
    count_records,
    dashboard_template_initial_values,
    default_sort_direction_for_field,
    default_sort_field_for_table,
    delete_record,
    default_portal_keywords_text,
    ensure_dashboard_widget_from_template,
    ensure_user_workspace,
    ensure_dashboard_panel_preferences,
    get_record,
    get_records,
    get_application_tags,
    get_application_tags_map,
    get_user_table,
    infer_progress,
    import_workbook_for_user,
    make_unique_column_slug,
    make_unique_table_slug,
    move_to_applied,
    parse_tag_names,
    replace_applied_history_events,
    refresh_all_portal_jobs_for_user,
    refresh_portal_jobs,
    remove_application_tags,
    save_record,
    sync_application_tags,
    sort_records_with_field,
)
from .notifications import build_email_body, send_notification_to_user, upcoming_deadlines_for_user
from .serializers import serialize_dashboard_widget
from .smtp_handler import SMTPHandler


class WorkspaceReadyMixin:
    def dispatch(self, request, *args, **kwargs):
        workbook = Path(settings.BASE_DIR) / "Applications.xlsx"
        ensure_user_workspace(request.user, workbook if workbook.exists() else None)
        return super().dispatch(request, *args, **kwargs)


class AuthLoginView(LoginView):
    template_name = "registration/login.html"
    redirect_authenticated_user = True
    authentication_form = EmailAwareAuthenticationForm

    def get_form(self, form_class=None):
        form = super().get_form(form_class)
        for field in form.fields.values():
            field.widget.attrs["class"] = (
                "mt-2 w-full rounded-2xl border border-slate-300/80 bg-white/85 px-4 py-3 "
                "text-sm text-slate-900 shadow-sm outline-none transition "
                "placeholder:text-slate-400 focus:border-sky-500 focus:ring-2 focus:ring-sky-200"
            )
        return form


class SignupView(FormView):
    template_name = "registration/signup.html"
    form_class = SignupForm

    def form_valid(self, form):
        user = form.save(commit=False)
        smtp_configured = SMTPHandler.is_configured()
        if smtp_configured:
            user.is_active = False
            user.save()
            try:
                send_verification_email(self.request, user)
            except Exception as exc:
                user.delete()
                form.add_error(None, f"Could not send the verification email right now: {exc}")
                return self.form_invalid(form)
            return self.render_to_response(
                self.get_context_data(
                    form=self.form_class(),
                    verification_pending=True,
                    verification_email=user.email,
                )
            )

        user.is_active = True
        user.email_verified_at = timezone.now()
        user.save()
        messages.success(
            self.request,
            "Your account is ready. SMTP is not configured, so email verification was skipped for this local setup.",
        )
        return redirect("login")


class VerifyEmailView(View):
    def get(self, request, uidb64: str, token: str):
        user = verify_user_from_token(uidb64, token)
        if user is None:
            messages.error(request, "That verification link is no longer valid.")
            return redirect("login")
        messages.success(request, "Your email is verified. You can sign in now.")
        return redirect("login")


class DashboardView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/dashboard.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        selected_board_id = self.request.GET.get("board")
        try:
            parsed_board_id = int(selected_board_id) if selected_board_id else None
        except ValueError:
            parsed_board_id = None
        dashboard_payload = build_dashboard(self.request.user, selected_board_id=parsed_board_id)
        context.update(dashboard_payload)
        context["nav_section"] = "dashboard"
        context["dashboard_create_url"] = reverse("workspace_record_create", args=["dashboard"])
        context["dashboard_manage_url"] = reverse("workspace_table", args=["dashboard"])
        context["dashboard_customize_url"] = reverse("dashboard_customize")
        context["dashboard_boards_url"] = reverse("dashboard_boards")
        context["dashboard_widgets"] = dashboard_payload["selected_homepage_widgets"]
        context["dashboard_designer_url"] = (
            reverse("dashboard_designer", args=[dashboard_payload["selected_homepage_board"].id])
            if dashboard_payload["selected_homepage_board"]
            else ""
        )
        context["dashboard_widgets_json"] = json.dumps(
            [serialize_dashboard_widget(widget) for widget in dashboard_payload["selected_homepage_widgets"]]
        )
        context["dashboard_secondary_panels"] = [
            *dashboard_payload["dashboard_panel_layout"]["left"],
            *dashboard_payload["dashboard_panel_layout"]["right"],
        ]
        context["dashboard_templates"] = [
            {
                **item,
                "row_key": str(item["row_key"]) if item["row_key"] else "",
                "create_url": f"{reverse('workspace_record_create', args=['dashboard'])}?template={item['key']}",
                "edit_url": reverse("workspace_record_edit", args=["dashboard", item["row_key"]]) if item["row_key"] else "",
            }
            for item in dashboard_payload["dashboard_templates"]
        ]
        return context


class DashboardCustomizeView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/dashboard_customize.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        ensure_dashboard_panel_preferences(self.request.user)
        catalog = dashboard_panel_catalog()
        preferences = list(
            DashboardPanelPreference.objects.filter(user=self.request.user).order_by("zone", "sort_order", "panel_key")
        )
        context.update(
            {
                "nav_section": "dashboard",
                "dashboard_url": reverse("dashboard"),
                "panel_zone_choices": DashboardPanelPreference.Zone.choices,
                "panel_preferences": [
                    {
                        **catalog.get(preference.panel_key, {"key": preference.panel_key, "label": preference.panel_key.title()}),
                        "zone": preference.zone,
                        "sort_order": preference.sort_order,
                    }
                    for preference in preferences
                ],
            }
        )
        return context

    def post(self, request, *args, **kwargs):
        ensure_dashboard_panel_preferences(request.user)
        allowed_zones = {choice[0] for choice in DashboardPanelPreference.Zone.choices}
        preferences = {
            preference.panel_key: preference
            for preference in DashboardPanelPreference.objects.filter(user=request.user)
        }
        for panel in dashboard_panel_catalog().values():
            preference = preferences.get(panel["key"])
            if preference is None:
                continue
            zone = (request.POST.get(f"zone_{panel['key']}") or preference.zone).strip()
            if zone not in allowed_zones:
                zone = preference.zone
            try:
                sort_order = max(1, int(request.POST.get(f"order_{panel['key']}") or preference.sort_order))
            except ValueError:
                sort_order = preference.sort_order
            changed_fields: list[str] = []
            if preference.zone != zone:
                preference.zone = zone
                changed_fields.append("zone")
            if preference.sort_order != sort_order:
                preference.sort_order = sort_order
                changed_fields.append("sort_order")
            if changed_fields:
                preference.save(update_fields=changed_fields + ["updated_at"])
        messages.success(request, "Dashboard layout updated.")
        return redirect("dashboard_customize")


class DashboardBoardListView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/dashboard_boards.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        payload = build_dashboard(self.request.user)
        boards = payload["dashboard_boards"]
        context.update(
            {
                "nav_section": "dashboard",
                "dashboard_url": reverse("dashboard"),
                "dashboard_customize_url": reverse("dashboard_customize"),
                "board_form": kwargs.get("board_form") or DashboardBoardForm(),
                "board_cards": [
                    {
                        "board": board,
                        "widget_count": board.items.count(),
                        "designer_url": reverse("dashboard_designer", args=[board.id]),
                        "preview_url": f"{reverse('dashboard')}?board={board.id}",
                        "delete_url": reverse("dashboard_board_delete", args=[board.id]),
                    }
                    for board in boards
                ],
            }
        )
        return context

    def post(self, request, *args, **kwargs):
        form = DashboardBoardForm(request.POST)
        if form.is_valid():
            board = form.save(commit=False)
            board.user = request.user
            board.sort_order = (
                DashboardBoard.objects.filter(user=request.user).order_by("-sort_order").values_list("sort_order", flat=True).first()
                or 0
            ) + 10
            board.save()
            messages.success(request, "Dashboard board created. You can start arranging plots now.")
            return redirect("dashboard_designer", board.id)
        return self.render_to_response(self.get_context_data(board_form=form))


class DashboardDesignerView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/dashboard_designer.html"

    def get_board(self) -> DashboardBoard:
        return get_object_or_404(
            DashboardBoard.objects.filter(user=self.request.user).prefetch_related("items"),
            id=self.kwargs["board_id"],
        )

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        board = self.get_board()
        payload = build_dashboard(self.request.user)
        widget_map = {widget.row_key: widget for widget in payload["all_dashboard_widgets"]}
        board_widgets = build_dashboard_board_widgets(board, widget_map)
        board_widget_row_keys = {widget.row_key for widget in board_widgets}
        available_widgets = [
            widget
            for widget in payload["all_dashboard_widgets"]
            if widget.row_key not in board_widget_row_keys
        ]
        context.update(
            {
                "nav_section": "dashboard",
                "board": board,
                "board_form": kwargs.get("board_form") or DashboardBoardForm(instance=board),
                "dashboard_url": f"{reverse('dashboard')}?board={board.id}",
                "dashboard_boards_url": reverse("dashboard_boards"),
                "dashboard_customize_url": reverse("dashboard_customize"),
                "dashboard_create_url": reverse("workspace_record_create", args=["dashboard"]),
                "dashboard_manage_url": reverse("workspace_table", args=["dashboard"]),
                "board_delete_url": reverse("dashboard_board_delete", args=[board.id]),
                "board_widgets": board_widgets,
                "board_widgets_json": json.dumps([serialize_dashboard_widget(widget) for widget in board_widgets]),
                "available_widgets": available_widgets,
                "available_widgets_json": json.dumps([serialize_dashboard_widget(widget) for widget in available_widgets]),
                "dashboard_templates": [
                    {
                        **item,
                        "row_key": str(item["row_key"]) if item["row_key"] else "",
                        "create_url": f"{reverse('workspace_record_create', args=['dashboard'])}?template={item['key']}",
                        "edit_url": reverse("workspace_record_edit", args=["dashboard", item["row_key"]]) if item["row_key"] else "",
                        "on_board": any(widget.template_key == item["key"] for widget in board_widgets),
                    }
                    for item in payload["dashboard_templates"]
                ],
            }
        )
        return context

    def post(self, request, *args, **kwargs):
        board = self.get_board()
        form = DashboardBoardForm(request.POST, instance=board)
        if form.is_valid():
            form.save()
            messages.success(request, "Dashboard board settings updated.")
            return redirect("dashboard_designer", board.id)
        return self.render_to_response(self.get_context_data(board_form=form))


class DashboardBoardDeleteView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    def post(self, request, board_id: int):
        board = get_object_or_404(DashboardBoard, user=request.user, id=board_id)
        board.delete()
        messages.success(request, "Dashboard board deleted.")
        return redirect("dashboard_boards")


class DashboardBoardWidgetCollectionView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    def post(self, request, board_id: int):
        board = get_object_or_404(DashboardBoard.objects.filter(user=request.user).prefetch_related("items"), id=board_id)
        try:
            payload = json.loads(request.body.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            return JsonResponse({"detail": "Invalid JSON payload."}, status=400)

        widget_row_key = (payload.get("widget_row_key") or "").strip()
        template_key = (payload.get("template_key") or "").strip()
        dashboard_table = get_user_table(request.user, "dashboard")
        if template_key:
            widget_record = ensure_dashboard_widget_from_template(request.user, dashboard_table, template_key)
            widget_row_key = str(widget_record.row_key)
        if not widget_row_key:
            return JsonResponse({"detail": "widget_row_key or template_key is required."}, status=400)

        board_item = add_dashboard_widget_to_board(
            board,
            widget_row_key,
            card_size=payload.get("card_size") or DashboardBoardItem.CardSize.STANDARD,
            chart_height=payload.get("chart_height") or DashboardBoardItem.ChartHeight.MEDIUM,
        )
        dashboard_payload = build_dashboard(request.user, selected_board_id=board.id)
        widget_map = {widget.row_key: widget for widget in dashboard_payload["all_dashboard_widgets"]}
        board_widgets = build_dashboard_board_widgets(board, widget_map)
        widget = next((item for item in board_widgets if item.layout_item_id == board_item.id), None)
        if widget is None:
            return JsonResponse({"detail": "Widget was added but could not be rendered."}, status=500)
        return JsonResponse({"widget": serialize_dashboard_widget(widget)}, status=201)


class DashboardBoardWidgetDetailView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    def patch(self, request, board_id: int, item_id: int):
        board = get_object_or_404(DashboardBoard, user=request.user, id=board_id)
        board_item = get_object_or_404(DashboardBoardItem, board=board, id=item_id)
        try:
            payload = json.loads(request.body.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            return JsonResponse({"detail": "Invalid JSON payload."}, status=400)

        allowed_card_sizes = {choice[0] for choice in DashboardBoardItem.CardSize.choices}
        allowed_chart_heights = {choice[0] for choice in DashboardBoardItem.ChartHeight.choices}
        changed_fields: list[str] = []

        card_size = (payload.get("card_size") or "").strip()
        if card_size and card_size in allowed_card_sizes and board_item.card_size != card_size:
            board_item.card_size = card_size
            changed_fields.append("card_size")

        chart_height = (payload.get("chart_height") or "").strip()
        if chart_height and chart_height in allowed_chart_heights and board_item.chart_height != chart_height:
            board_item.chart_height = chart_height
            changed_fields.append("chart_height")

        if "sort_order" in payload:
            try:
                sort_order = max(1, int(payload.get("sort_order") or board_item.sort_order))
            except (TypeError, ValueError):
                sort_order = board_item.sort_order
            if board_item.sort_order != sort_order:
                board_item.sort_order = sort_order
                changed_fields.append("sort_order")

        if changed_fields:
            board_item.save(update_fields=changed_fields + ["updated_at"])

        dashboard_payload = build_dashboard(request.user, selected_board_id=board.id)
        widget_map = {widget.row_key: widget for widget in dashboard_payload["all_dashboard_widgets"]}
        board_widgets = build_dashboard_board_widgets(board, widget_map)
        widget = next((item for item in board_widgets if item.layout_item_id == board_item.id), None)
        if widget is None:
            return JsonResponse({"detail": "Updated layout could not be rendered."}, status=500)
        return JsonResponse({"widget": serialize_dashboard_widget(widget)})

    def delete(self, request, board_id: int, item_id: int):
        board = get_object_or_404(DashboardBoard, user=request.user, id=board_id)
        board_item = get_object_or_404(DashboardBoardItem, board=board, id=item_id)
        board_item.delete()
        return JsonResponse({"deleted": True})


class BaseTableListView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/table_page.html"
    table_slug = ""
    page_title = ""
    page_description = ""
    nav_section = ""

    def get_table(self) -> DynamicTable:
        return get_user_table(self.request.user, self.table_slug)

    def get_detail_url(self, table: DynamicTable, record) -> str:
        if table.slug == "to-apply":
            return reverse("to_apply_detail", args=[record.row_key])
        if table.slug == "applied":
            return reverse("applied_detail", args=[record.row_key])
        return ""

    def get_record_tags_map(self, table: DynamicTable, records) -> dict:
        request_user = getattr(self.request, "user", None)
        if table.slug not in {"to-apply", "applied"} or request_user is None:
            return {}
        return get_application_tags_map(request_user, table.slug, [record.row_key for record in records])

    def get_tag_filter(self, table: DynamicTable) -> str:
        request_user = getattr(self.request, "user", None)
        if table.slug not in {"to-apply", "applied"} or request_user is None:
            return ""
        requested = (self.request.GET.get("tag") or "").strip()
        allowed = set(
            ApplicationTag.objects.filter(user=request_user).values_list("slug", flat=True)
        )
        return requested if requested in allowed else ""

    def get_tag_filter_options(self, table: DynamicTable) -> list[dict[str, str]]:
        request_user = getattr(self.request, "user", None)
        if table.slug not in {"to-apply", "applied"} or request_user is None:
            return []
        return [
            {"value": tag.slug, "label": tag.name}
            for tag in ApplicationTag.objects.filter(user=request_user).order_by("name")
        ]

    def filter_records_for_display(self, records, table: DynamicTable):
        keyword = (self.request.GET.get("q") or "").strip().lower()
        if keyword:
            tokens = [token for token in keyword.split() if token]
            filtered = []
            for record in records:
                haystack_parts = [record.label]
                haystack_parts.extend(str(value) for value in record.values.values() if value is not None)
                haystack = " ".join(haystack_parts).lower()
                if all(token in haystack for token in tokens):
                    filtered.append(record)
            records = filtered
        tag_filter = self.get_tag_filter(table)
        if not tag_filter:
            return records
        tags_map = self.get_record_tags_map(table, records)
        return [
            record
            for record in records
            if any(tag.slug == tag_filter for tag in tags_map.get(record.row_key, []))
        ]

    def build_rows(self, table: DynamicTable, records):
        tags_map = self.get_record_tags_map(table, records)
        rows = []
        for record in records:
            if table.slug == "to-apply":
                edit_url = reverse("to_apply_edit", args=[record.row_key])
                delete_url = reverse("to_apply_delete", args=[record.row_key])
                promote_url = reverse("to_apply_promote", args=[record.row_key])
            elif table.slug == "applied":
                edit_url = reverse("applied_edit", args=[record.row_key])
                delete_url = reverse("applied_delete", args=[record.row_key])
                promote_url = ""
            elif table.slug == "portals":
                edit_url = reverse("workspace_record_edit", args=[table.slug, record.row_key])
                delete_url = reverse("workspace_record_delete", args=[table.slug, record.row_key])
                promote_url = ""
            else:
                edit_url = reverse("workspace_record_edit", args=[table.slug, record.row_key])
                delete_url = reverse("workspace_record_delete", args=[table.slug, record.row_key])
                promote_url = ""
            primary_cells, secondary_cells = self.split_cells(table, record)
            rows.append(
                {
                    "record": record,
                    "edit_url": edit_url,
                    "delete_url": delete_url,
                    "promote_url": promote_url,
                    "detail_url": self.get_detail_url(table, record),
                    "tags": tags_map.get(record.row_key, []),
                    "jobs_url": reverse("portal_jobs", args=[record.row_key]) if table.slug == "portals" else "",
                    "primary_cells": primary_cells,
                    "secondary_cells": secondary_cells,
                }
            )
        return rows

    def priority_slugs(self, table: DynamicTable) -> list[str]:
        if table.slug == "applied":
            return ["application-progress", "result", "type", "deadline", "number-of-positions", "date", "source", "url"]
        if table.slug == "to-apply":
            return ["type", "deadline", "number-of-positions", "source", "english-test", "url"]
        if table.slug == "portals":
            return ["url", "fetch-keywords"]
        return ["type", "status", "date", "url"]

    def split_cells(self, table: DynamicTable, record):
        visible_cells = [cell for cell in record.cells if cell.value]
        by_slug = {cell.column.slug: cell for cell in visible_cells}
        primary_cells = []
        seen: set[str] = set()
        for slug in self.priority_slugs(table):
            cell = by_slug.get(slug)
            if cell and slug not in seen:
                primary_cells.append(cell)
                seen.add(slug)
        for cell in visible_cells:
            if len(primary_cells) >= 4:
                break
            if cell.column.slug not in seen:
                primary_cells.append(cell)
                seen.add(cell.column.slug)
        secondary_cells = [cell for cell in record.cells if cell.column.slug not in seen and cell.value]
        return primary_cells, secondary_cells

    def get_create_url(self, table: DynamicTable) -> str:
        if table.slug == "to-apply":
            return reverse("to_apply_create")
        if table.slug == "applied":
            return reverse("applied_create")
        return reverse("workspace_record_create", args=[table.slug])

    def get_view_mode(self) -> str:
        return self.request.GET.get("view", "cards") if self.request.GET.get("view") in {"cards", "table"} else "cards"

    def get_sort_direction(self, table: DynamicTable) -> str:
        requested = self.request.GET.get("dir")
        if requested in {"asc", "desc"}:
            return requested
        return default_sort_direction_for_field(table, self.get_sort_field(table))

    def get_sort_field(self, table: DynamicTable) -> str:
        allowed = {"label", "created_at", "updated_at", *table.columns.filter(is_visible=True).values_list("slug", flat=True)}
        if table.slug in {"to-apply", "applied"}:
            allowed.add("tags")
        requested = self.request.GET.get("sort")
        if requested in allowed:
            return requested
        return default_sort_field_for_table(table)

    def get_sort_options(self, table: DynamicTable) -> list[dict[str, str]]:
        options = [
            {"value": "updated_at", "label": "Recently updated"},
            {"value": "created_at", "label": "Created time"},
            {"value": "label", "label": "Title / label"},
        ]
        if table.slug in {"to-apply", "applied"}:
            options.append({"value": "tags", "label": "Tags"})
        for column in table.columns.filter(is_visible=True):
            options.append({"value": column.slug, "label": column.name})
        return options

    def sort_records_for_display(self, records, table: DynamicTable):
        filtered_records = self.filter_records_for_display(records, table)
        if self.get_sort_field(table) == "tags":
            tags_map = self.get_record_tags_map(table, filtered_records)
            populated = []
            missing = []
            for record in filtered_records:
                tag_names = [tag.name for tag in tags_map.get(record.row_key, [])]
                if not tag_names:
                    missing.append(record)
                    continue
                populated.append(
                    (
                        ", ".join(sorted(tag_names, key=str.lower)).lower(),
                        record.label.lower(),
                        record,
                    )
                )
            ordered = sorted(
                populated,
                key=lambda item: (item[0], item[1]),
                reverse=self.get_sort_direction(table) == "desc",
            )
            return [item[2] for item in ordered] + missing
        return sort_records_with_field(
            filtered_records,
            table,
            sort_field=self.get_sort_field(table),
            direction=self.get_sort_direction(table),
        )

    def get_extra_actions(self, table: DynamicTable) -> list[dict[str, str]]:
        return []

    def get_bulk_actions(self, table: DynamicTable) -> list[dict[str, str]]:
        if table.slug == "to-apply":
            return [
                {"value": "move_to_applied", "label": "Move to Applied"},
                {"value": "add_tags", "label": "Add tags"},
                {"value": "remove_tags", "label": "Remove tags"},
                {"value": "delete", "label": "Delete selected"},
            ]
        if table.slug == "applied":
            return [
                {"value": "add_tags", "label": "Add tags"},
                {"value": "remove_tags", "label": "Remove tags"},
                {"value": "delete", "label": "Delete selected"},
            ]
        return []

    def get_bulk_action_url(self, table: DynamicTable) -> str:
        if table.slug == "to-apply":
            return reverse("to_apply_bulk")
        if table.slug == "applied":
            return reverse("applied_bulk")
        return ""

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        table = self.get_table()
        records = self.sort_records_for_display(get_records(table, self.request.user), table)
        context.update(
            {
                "table": table,
                "columns": table.columns.filter(is_visible=True),
                "table_rows": self.build_rows(table, records),
                "record_count": len(records),
                "page_title": self.page_title or table.name,
                "page_description": self.page_description or table.description,
                "create_url": self.get_create_url(table),
                "nav_section": self.nav_section or table.slug,
                "view_mode": self.get_view_mode(),
                "sort_field": self.get_sort_field(table),
                "sort_direction": self.get_sort_direction(table),
                "sort_options": self.get_sort_options(table),
                "tag_filter": self.get_tag_filter(table),
                "search_query": (self.request.GET.get("q") or "").strip(),
                "tag_filter_options": self.get_tag_filter_options(table),
                "extra_actions": self.get_extra_actions(table),
                "bulk_actions": self.get_bulk_actions(table),
                "bulk_action_url": self.get_bulk_action_url(table),
            }
        )
        return context


class ToApplyListView(BaseTableListView):
    table_slug = "to-apply"
    page_title = "To Apply"
    page_description = "Queue up roles and programs you still want to apply for."
    nav_section = "to-apply"


class AppliedListView(BaseTableListView):
    table_slug = "applied"
    page_title = "Applied"
    page_description = "Track results, progress, cover letter text, and CV text."
    nav_section = "applied"


class PortalsListView(BaseTableListView):
    table_slug = "portals"
    page_title = "Portals"
    page_description = "Fetch openings from saved portals and push good matches into To Apply."
    nav_section = "portals"

    def get_extra_actions(self, table: DynamicTable) -> list[dict[str, str]]:
        return [
            {
                "label": "All fetched jobs",
                "url": reverse("all_portal_jobs"),
                "classes": "rounded-full bg-white px-4 py-2 text-sm font-medium text-slate-700 ring-1 ring-slate-200 transition hover:bg-slate-50 dark:bg-slate-800 dark:text-slate-100 dark:ring-slate-700 dark:hover:bg-slate-700",
            },
            {
                "label": "Fetch all portals",
                "url": reverse("fetch_all_portal_jobs"),
                "method": "post",
                "classes": "rounded-full bg-sky-100 px-4 py-2 text-sm font-medium text-sky-800 transition hover:bg-sky-200 dark:bg-sky-500/20 dark:text-sky-200 dark:hover:bg-sky-500/30",
                "confirm_message": "Fetch jobs from all saved portals now?",
            },
        ]


class WorkspaceTableView(BaseTableListView):
    def get_table(self) -> DynamicTable:
        return get_user_table(self.request.user, self.kwargs["table_slug"])


class BaseRecordFormView(LoginRequiredMixin, WorkspaceReadyMixin, FormView):
    template_name = "tracker/record_form.html"
    table_slug = ""
    form_class = DynamicRecordForm
    success_view_name = ""
    page_title = ""
    page_description = ""
    nav_section = ""

    def get_table(self) -> DynamicTable:
        slug = self.table_slug or self.kwargs.get("table_slug")
        return get_user_table(self.request.user, slug)

    def get_initial(self) -> dict[str, Any]:
        initial = super().get_initial()
        row_key = self.kwargs.get("row_key")
        if row_key:
            record = get_record(self.get_table(), self.request.user, row_key)
            if record:
                initial.update(record.values)
        elif self.get_table().slug == "dashboard":
            template_key = (self.request.GET.get("template") or "").strip()
            if template_key:
                try:
                    initial.update(dashboard_template_initial_values(template_key))
                except ValueError:
                    pass
        return initial

    def get_form_kwargs(self):
        kwargs = super().get_form_kwargs()
        kwargs["table"] = self.get_table()
        kwargs["user"] = self.request.user
        kwargs["row_key"] = self.kwargs.get("row_key")
        return kwargs

    def get_success_url(self) -> str:
        if self.success_view_name == "workspace_table":
            return reverse(self.success_view_name, args=[self.get_table().slug])
        return reverse(self.success_view_name)

    def get_return_url(self) -> str:
        table = self.get_table()
        if table.slug == "to-apply":
            return reverse("to_apply_list")
        if table.slug == "applied":
            return reverse("applied_list")
        return reverse("workspace_table", args=[table.slug])

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        table = self.get_table()
        form = context.get("form")
        regular_fields = []
        date_fields = []
        if form is not None:
            for field in form:
                if self._is_date_like_field(field.name):
                    date_fields.append(field)
                else:
                    regular_fields.append(field)
            date_fields.sort(key=lambda field: self._date_field_rank(field.name))
        context.update(
            {
                "table": table,
                "page_title": self.page_title or f"Add {table.name} Entry",
                "page_description": self.page_description or table.description,
                "nav_section": self.nav_section or table.slug,
                "is_edit": bool(self.kwargs.get("row_key")),
                "return_url": self.get_return_url(),
                "available_job_tags": getattr(context.get("form"), "available_tags", []),
                "dashboard_columns_json": json.dumps(getattr(context.get("form"), "dashboard_columns_map", {})),
                "url_inference_api_url": reverse("api_infer_record_from_url"),
                "form_regular_fields": regular_fields,
                "form_date_fields": date_fields,
            }
        )
        return context

    @staticmethod
    def _is_date_like_field(field_name: str) -> bool:
        normalized = (field_name or "").strip().lower()
        return (
            normalized in {"date", "deadline", "result-on", "created_at", "updated_at", "application-date", "applied-date"}
            or normalized.endswith("_date")
            or normalized.endswith("-date")
        )

    @staticmethod
    def _date_field_rank(field_name: str) -> tuple[int, str]:
        normalized = (field_name or "").strip().lower()
        history_stage_order = [
            "history_applied_date",
            "history_screening_date",
            "history_interview_date",
            "history_technical_interview_date",
            "history_assignment_date",
            "history_offer_date",
            "history_accepted_date",
            "history_rejected_date",
            "history_withdrawn_date",
        ]
        priority = {
            "source_row_key": 0,
            "applied-date": 10,
            "application-date": 10,
            "date": 10,
            "deadline": 20,
            "result-on": 25,
            "created_at": 30,
            "updated_at": 40,
        }
        if normalized in history_stage_order:
            return (50 + history_stage_order.index(normalized), normalized)
        if normalized.startswith("history_") and normalized.endswith("_date"):
            return (80, normalized)
        return (priority.get(normalized, 90), normalized)

    def sync_form_tags(self, form, row_key) -> None:
        raw_tags = form.cleaned_data.get("tags")
        if raw_tags is None:
            return
        sync_application_tags(
            self.request.user,
            self.get_table().slug,
            row_key,
            parse_tag_names(raw_tags),
        )

    def after_record_saved(self, form, row_key) -> None:
        return

    def form_valid(self, form):
        table = self.get_table()
        if not table.columns.exists():
            messages.warning(self.request, "Add at least one column before creating records.")
            return redirect("manage_table_columns", table.slug)
        payload = {
            key: value
            for key, value in form.cleaned_data.items()
            if key != "tags" and not key.startswith("history_")
        }
        row_uuid = save_record(
            table,
            self.request.user,
            payload,
            row_key=self.kwargs.get("row_key"),
        )
        self.sync_form_tags(form, row_uuid)
        self.after_record_saved(form, row_uuid)
        messages.success(self.request, f"{table.name} entry saved.")
        return redirect(self.get_success_url())


class ToApplyCreateView(BaseRecordFormView):
    table_slug = "to-apply"
    form_class = ToApplyForm
    success_view_name = "to_apply_list"
    page_title = "Add To Apply Entry"
    page_description = "Capture an opportunity you want to revisit."
    nav_section = "to-apply"


class ToApplyUpdateView(ToApplyCreateView):
    page_title = "Edit To Apply Entry"
    page_description = "Update the opportunity details in your queue."


class AppliedCreateView(BaseRecordFormView):
    table_slug = "applied"
    form_class = AppliedCreateForm
    success_view_name = "applied_list"
    page_title = "Add Applied Entry"
    page_description = "Create a new applied record or promote one directly from To Apply."
    nav_section = "applied"

    def get_form_kwargs(self):
        kwargs = super().get_form_kwargs()
        kwargs["source_records"] = get_records(get_user_table(self.request.user, "to-apply"), self.request.user)
        return kwargs

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context["source_records_json"] = json.dumps(
            [
                {"row_key": str(record.row_key), "values": record.values}
                for record in get_records(get_user_table(self.request.user, "to-apply"), self.request.user)
            ]
        )
        return context

    def form_valid(self, form):
        payload = {
            key: value
            for key, value in form.cleaned_data.items()
            if key not in {"source_row_key", "tags"}
        }
        source_row_key = form.cleaned_data.get("source_row_key") or None
        row_uuid = move_to_applied(
            user=self.request.user,
            applied_table=self.get_table(),
            to_apply_table=get_user_table(self.request.user, "to-apply"),
            values=payload,
            source_row_key=source_row_key,
        )
        raw_tags = form.cleaned_data.get("tags") or ""
        if raw_tags.strip() or not source_row_key:
            sync_application_tags(
                self.request.user,
                self.get_table().slug,
                row_uuid,
                parse_tag_names(raw_tags),
            )
        messages.success(self.request, "Applied entry saved.")
        if source_row_key:
            return redirect("applied_edit", row_uuid)
        return redirect(self.get_success_url())


class AppliedUpdateView(BaseRecordFormView):
    table_slug = "applied"
    form_class = AppliedUpdateForm
    success_view_name = "applied_list"
    page_title = "Edit Applied Entry"
    page_description = "Update progress, results, cover letter text, or CV text."
    nav_section = "applied"

    def after_record_saved(self, form, row_key) -> None:
        history_dates = {}
        for stage in form.HISTORY_STAGE_CHOICES:
            field_name = form.history_field_name(stage)
            selected_date = form.cleaned_data.get(field_name)
            if selected_date:
                history_dates[stage] = selected_date
        fallback_applied_date = form.cleaned_data.get("date")
        if fallback_applied_date and "Applied" not in history_dates:
            history_dates["Applied"] = fallback_applied_date
        if history_dates:
            replace_applied_history_events(self.request.user, row_key, history_dates)


class BaseJobDetailView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/job_detail.html"
    table_slug = ""
    nav_section = ""

    def get_table(self) -> DynamicTable:
        return get_user_table(self.request.user, self.table_slug)

    def get_record(self):
        record = get_record(self.get_table(), self.request.user, self.kwargs["row_key"])
        if not record:
            raise Http404("Job entry not found.")
        return record

    def get_back_url(self) -> str:
        if self.table_slug == "to-apply":
            return reverse("to_apply_list")
        return reverse("applied_list")

    def get_edit_url(self, row_key) -> str:
        if self.table_slug == "to-apply":
            return reverse("to_apply_edit", args=[row_key])
        return reverse("applied_edit", args=[row_key])

    def get_delete_url(self, row_key) -> str:
        if self.table_slug == "to-apply":
            return reverse("to_apply_delete", args=[row_key])
        return reverse("applied_delete", args=[row_key])

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        table = self.get_table()
        record = self.get_record()
        timeline_items, timeline_plot = build_application_history_timeline(
            self.request.user,
            table.slug,
            record,
        )
        context.update(
            {
                "table": table,
                "record": record,
                "record_tags": get_application_tags(self.request.user, table.slug, record.row_key),
                "timeline_items": timeline_items,
                "timeline_plot_json": json.dumps(timeline_plot),
                "nav_section": self.nav_section or table.slug,
                "back_url": self.get_back_url(),
                "edit_url": self.get_edit_url(record.row_key),
                "delete_url": self.get_delete_url(record.row_key),
                "current_stage": "To Apply" if table.slug == "to-apply" else infer_progress(record),
                "external_url": record.values.get("url", ""),
            }
        )
        return context


class ToApplyDetailView(BaseJobDetailView):
    table_slug = "to-apply"
    nav_section = "to-apply"


class AppliedDetailView(BaseJobDetailView):
    table_slug = "applied"
    nav_section = "applied"


class BaseBulkJobActionView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    table_slug = ""

    def get_table(self) -> DynamicTable:
        return get_user_table(self.request.user, self.table_slug)

    def get_success_url(self) -> str:
        if self.table_slug == "to-apply":
            return reverse("to_apply_list")
        return reverse("applied_list")

    def get_redirect_response(self, request):
        redirect_to = (request.POST.get("next") or "").strip()
        if redirect_to and url_has_allowed_host_and_scheme(
            url=redirect_to,
            allowed_hosts={request.get_host()},
            require_https=request.is_secure(),
        ):
            return redirect(redirect_to)
        return redirect(self.get_success_url())

    def post(self, request):
        table = self.get_table()
        row_keys = [item for item in request.POST.getlist("selected_row_keys") if item]
        if not row_keys:
            messages.warning(request, "Select at least one item first.")
            return self.get_redirect_response(request)

        action = (request.POST.get("bulk_action") or "").strip()
        tag_names = parse_tag_names(request.POST.get("bulk_tags") or "")
        processed = 0

        if action == "delete":
            for row_key in row_keys:
                delete_record(table, request.user, row_key)
                processed += 1
            messages.success(request, f"Deleted {processed} selected item(s).")
            return self.get_redirect_response(request)

        if action == "move_to_applied" and table.slug == "to-apply":
            applied_table = get_user_table(request.user, "applied")
            moved_row_keys: list[str] = []
            for row_key in row_keys:
                moved_key = move_to_applied(
                    user=request.user,
                    applied_table=applied_table,
                    to_apply_table=table,
                    values={},
                    source_row_key=row_key,
                )
                moved_row_keys.append(str(moved_key))
                processed += 1
            messages.success(request, f"Moved {processed} selected item(s) to Applied.")
            if len(moved_row_keys) == 1:
                return redirect("applied_edit", moved_row_keys[0])
            return self.get_redirect_response(request)

        if action == "add_tags":
            if not tag_names:
                messages.error(request, "Enter at least one tag to add.")
                return self.get_redirect_response(request)
            for row_key in row_keys:
                add_application_tags(request.user, table.slug, row_key, tag_names)
                processed += 1
            messages.success(request, f"Added tags to {processed} selected item(s).")
            return self.get_redirect_response(request)

        if action == "remove_tags":
            if not tag_names:
                messages.error(request, "Enter at least one tag to remove.")
                return self.get_redirect_response(request)
            for row_key in row_keys:
                remove_application_tags(request.user, table.slug, row_key, tag_names)
                processed += 1
            messages.success(request, f"Removed tags from {processed} selected item(s).")
            return self.get_redirect_response(request)

        messages.error(request, "That bulk action is not available here.")
        return self.get_redirect_response(request)


class ToApplyBulkActionView(BaseBulkJobActionView):
    table_slug = "to-apply"


class AppliedBulkActionView(BaseBulkJobActionView):
    table_slug = "applied"


class WorkspaceRecordCreateView(BaseRecordFormView):
    success_view_name = "workspace_table"


class WorkspaceRecordUpdateView(WorkspaceRecordCreateView):
    pass


class ManageTablesView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/manage_tables.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        tables = DynamicTable.objects.filter(owner=self.request.user).prefetch_related("columns")
        context.update(
            {
                "nav_section": "tables",
                "table_form": kwargs.get("table_form") or TableDefinitionForm(),
                "workbook_import_form": kwargs.get("workbook_import_form") or WorkbookImportForm(),
                "tables": [
                    {
                        "table": table,
                        "record_count": count_records(table),
                    }
                    for table in tables
                ],
            }
        )
        return context

    def handle_workbook_import(self, request):
        form = WorkbookImportForm(request.POST, request.FILES)
        if not form.is_valid():
            return self.render_to_response(self.get_context_data(workbook_import_form=form))

        workbook = form.cleaned_data["workbook"]
        replace = form.cleaned_data["import_mode"] == WorkbookImportForm.ImportMode.REPLACE
        temp_path: Path | None = None

        try:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                for chunk in workbook.chunks():
                    temp_file.write(chunk)
                temp_path = Path(temp_file.name)

            imported_counts = import_workbook_for_user(temp_path, request.user, replace=replace)
        except Exception as exc:
            form.add_error(None, f"Could not import that workbook: {exc}")
            return self.render_to_response(self.get_context_data(workbook_import_form=form))
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink(missing_ok=True)

        mode_label = "replaced" if replace else "appended"
        counts_text = ", ".join(f"{slug}: {count}" for slug, count in imported_counts.items())
        messages.success(
            request,
            f"Workbook imported successfully. Existing data was {mode_label}. {counts_text}",
        )
        return redirect("manage_tables")

    def post(self, request, *args, **kwargs):
        if "import_workbook" in request.POST:
            return self.handle_workbook_import(request)

        form = TableDefinitionForm(request.POST)
        if form.is_valid():
            table = form.save(commit=False)
            table.owner = request.user
            table.slug = make_unique_table_slug(request.user, table.name)
            table.display_order = (
                DynamicTable.objects.filter(owner=request.user).order_by("-display_order").values_list("display_order", flat=True).first() or 0
            ) + 10
            table.save()
            messages.success(request, f"{table.name} created. Add columns to start using it.")
            return redirect("manage_table_columns", table.slug)
        return self.render_to_response(self.get_context_data(table_form=form))


class ManageTableColumnsView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/manage_columns.html"

    def get_table(self):
        return get_user_table(self.request.user, self.kwargs["table_slug"])

    def get_context_data(self, **kwargs):
        table = self.get_table()
        context = super().get_context_data(**kwargs)
        context.update(
            {
                "nav_section": "tables",
                "table": table,
                "column_form": kwargs.get("column_form") or ColumnDefinitionForm(),
                "columns": table.columns.all(),
            }
        )
        return context

    def post(self, request, *args, **kwargs):
        table = self.get_table()
        form = ColumnDefinitionForm(request.POST)
        if form.is_valid():
            column = form.save(commit=False)
            column.table = table
            column.slug = make_unique_column_slug(table, column.name)
            column.position = (table.columns.order_by("-position").values_list("position", flat=True).first() or 0) + 1
            column.save()
            messages.success(request, f"Column {column.name} added to {table.name}.")
            return redirect("manage_table_columns", table.slug)
        return self.render_to_response(self.get_context_data(column_form=form))


class AccountSettingsView(LoginRequiredMixin, WorkspaceReadyMixin, FormView):
    template_name = "tracker/account_settings.html"
    form_class = AccountSettingsForm

    def get_form_kwargs(self):
        kwargs = super().get_form_kwargs()
        kwargs["instance"] = get_user_model().objects.get(pk=self.request.user.pk)
        return kwargs

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context["nav_section"] = "settings"
        context["smtp_configured"] = SMTPHandler.is_configured()
        context["notification_preview"] = kwargs.get("notification_preview", "")
        context["last_notification_sent_at"] = self.request.user.last_notification_sent_at
        context["last_notification_error"] = self.request.user.last_notification_error
        context["last_notification_error_at"] = self.request.user.last_notification_error_at
        return context

    def get_success_url(self):
        return reverse("account_settings")

    def post(self, request, *args, **kwargs):
        form = self.get_form()
        if "preview_notification" in request.POST:
            if form.is_valid():
                user = form.save(commit=False)
                if not user.notification_email:
                    user.notification_email = user.email
                preview_body = build_email_body(user, upcoming_deadlines_for_user(user), [])
                return self.render_to_response(self.get_context_data(form=form, notification_preview=preview_body))
            return self.render_to_response(self.get_context_data(form=form))
        if "send_notification_now" in request.POST:
            if not SMTPHandler.is_configured():
                messages.error(request, "SMTP is not configured yet, so the notification email cannot be sent.")
                return self.render_to_response(self.get_context_data(form=form))
            if form.is_valid():
                user = form.save(commit=False)
                if not user.notification_email:
                    user.notification_email = user.email
                user.save()
                try:
                    result = send_notification_to_user(user, force=True, dry_run=False)
                except Exception as exc:
                    messages.error(request, f"Could not send the notification email right now: {exc}")
                    return redirect(self.get_success_url())
                if not result.get("sent"):
                    messages.error(request, "Notification email was not sent. Please check your notification settings.")
                    return redirect(self.get_success_url())
                message_id = (result.get("message_id") or "").strip()
                if message_id:
                    messages.success(
                        request,
                        f"Notification email accepted by SMTP for {result['recipient']} (Message-ID: {message_id}).",
                    )
                    return redirect(self.get_success_url())
                messages.success(
                    request,
                    f"Notification email accepted by SMTP for {result['recipient']}.",
                )
                return redirect(self.get_success_url())
            return self.render_to_response(self.get_context_data(form=form))
        if form.is_valid():
            return self.form_valid(form)
        return self.form_invalid(form)

    def form_valid(self, form):
        user = form.save(commit=False)
        previous_notify_time = (
            type(user)
            .objects.filter(pk=user.pk)
            .values_list("notify_time", flat=True)
            .first()
        )
        if not user.notification_email:
            user.notification_email = user.email
        if user.notify_time != previous_notify_time:
            user.last_notification_sent_at = None
        user.save()
        messages.success(self.request, "Notification settings updated.")
        return redirect(self.get_success_url())


class PortalJobsView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/portal_jobs.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        portal_table = get_user_table(self.request.user, "portals")
        portal_record = get_record(portal_table, self.request.user, self.kwargs["row_key"])
        if not portal_record:
            raise Http404("Portal entry not found.")
        jobs = PortalJobSuggestion.objects.filter(
            user=self.request.user,
            portal_table=portal_table,
            portal_row_key=self.kwargs["row_key"],
        ).order_by("title")
        context.update(
            {
                "nav_section": "portals",
                "portal_record": portal_record,
                "jobs": jobs,
                "current_keywords": portal_record.values.get("fetch-keywords") or default_portal_keywords_text(),
                "all_portal_jobs_url": reverse("all_portal_jobs"),
                "fetch_all_portal_jobs_url": reverse("fetch_all_portal_jobs"),
            }
        )
        return context


class AllPortalJobsView(LoginRequiredMixin, WorkspaceReadyMixin, TemplateView):
    template_name = "tracker/all_portal_jobs.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        portal_table = get_user_table(self.request.user, "portals")
        portal_records = get_records(portal_table, self.request.user)
        jobs = PortalJobSuggestion.objects.filter(user=self.request.user).order_by("-fetched_at", "portal_name", "title")
        context.update(
            {
                "nav_section": "portals",
                "jobs": jobs,
                "job_count": jobs.count(),
                "portal_count": len(portal_records),
                "fetch_all_portal_jobs_url": reverse("fetch_all_portal_jobs"),
                "portals_url": reverse("portals_list"),
            }
        )
        return context


class FetchPortalJobsView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    def post(self, request, row_key):
        try:
            refresh_portal_jobs(
                request.user,
                get_user_table(request.user, "portals"),
                row_key,
                raw_keywords=request.POST.get("keywords"),
            )
            messages.success(request, "Fetched jobs from the selected portal.")
        except Exception as exc:
            messages.error(request, f"Could not fetch jobs: {exc}")
        return redirect("portal_jobs", row_key)


class FetchAllPortalJobsView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    def post(self, request):
        redirect_to = (request.POST.get("next") or "").strip() or reverse("all_portal_jobs")
        try:
            result = refresh_all_portal_jobs_for_user(request.user)
            if result["refreshed"]:
                portal_label = "portal" if result["refreshed"] == 1 else "portals"
                messages.success(
                    request,
                    f"Fetched {result['total_jobs']} jobs across {result['refreshed']} {portal_label}.",
                )
            else:
                messages.info(request, "No saved portals were available to fetch.")
            if result["errors"]:
                sample = " | ".join(result["errors"][:3])
                if len(result["errors"]) > 3:
                    sample += f" | and {len(result['errors']) - 3} more"
                messages.warning(request, f"Some portals could not be fetched: {sample}")
        except Exception as exc:
            messages.error(request, f"Could not fetch jobs from all portals: {exc}")
        return redirect(redirect_to)


class AddPortalJobToApplyView(LoginRequiredMixin, WorkspaceReadyMixin, View):
    def post(self, request, job_id: int):
        suggestion = get_object_or_404(PortalJobSuggestion, id=job_id, user=request.user)
        add_portal_job_to_to_apply(request.user, suggestion)
        messages.success(request, "Job added to To Apply.")
        redirect_to = (request.POST.get("next") or "").strip()
        if redirect_to:
            return redirect(redirect_to)
        return redirect("portal_jobs", suggestion.portal_row_key)


@login_required
def promote_to_applied(request, row_key):
    redirect_to = (request.POST.get("next") or "").strip()
    if request.method == "POST":
        try:
            moved_row_key = move_to_applied(
                user=request.user,
                applied_table=get_user_table(request.user, "applied"),
                to_apply_table=get_user_table(request.user, "to-apply"),
                values={},
                source_row_key=row_key,
            )
            messages.success(request, "Job moved to Applied.")
            return redirect("applied_edit", moved_row_key)
        except ValueError as exc:
            messages.error(request, str(exc))

    if redirect_to and url_has_allowed_host_and_scheme(
        url=redirect_to,
        allowed_hosts={request.get_host()},
        require_https=request.is_secure(),
    ):
        return redirect(redirect_to)
    return redirect("to_apply_list")


@login_required
def delete_to_apply(request, row_key):
    if request.method == "POST":
        delete_record(get_user_table(request.user, "to-apply"), request.user, row_key)
        messages.success(request, "To Apply entry deleted.")
    return redirect("to_apply_list")


@login_required
def delete_applied(request, row_key):
    if request.method == "POST":
        delete_record(get_user_table(request.user, "applied"), request.user, row_key)
        messages.success(request, "Applied entry deleted.")
    return redirect("applied_list")


@login_required
def delete_workspace_record(request, table_slug, row_key):
    if request.method == "POST":
        delete_record(get_user_table(request.user, table_slug), request.user, row_key)
        messages.success(request, "Record deleted.")
    return redirect("workspace_table", table_slug)


@login_required
def export_applications(request):
    workbook_path = Path(settings.BASE_DIR) / "Applications.xlsx"
    ensure_user_workspace(request.user, workbook_path if workbook_path.exists() else None)
    workbook = build_export_workbook(
        request.user,
        template_path=workbook_path if workbook_path.exists() else None,
    )
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    response = HttpResponse(
        output.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="Applications.xlsx"'
    return response
