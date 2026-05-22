from __future__ import annotations

from datetime import date, datetime

from django import forms
from django.contrib.auth.forms import AuthenticationForm, UserCreationForm
from django.core.exceptions import ValidationError

from .models import ApplicationTag, DashboardBoard, DynamicColumn, DynamicTable, User
from .services import APPLICATION_STAGE_ORDER, get_application_history_events, get_application_tags

FORM_APPLICATION_STAGE_ORDER = [stage for stage in APPLICATION_STAGE_ORDER if stage != "To Apply"]

APPLICATION_RESULT_ORDER = [
    "Pending",
    "Offer",
    "Accepted",
    "Rejected",
    "Withdrawn",
]

FRIENDLY_JOB_FIELD_LABELS = {
    "to-apply": {
        "name": "Organization Name",
        "role-title": "Role / Position Title",
        "url": "Job Posting URL",
        "source": "Source / Portal",
        "type": "Opportunity Type",
        "deadline": "Application Deadline",
        "number-of-positions": "Open Positions",
        "english-test": "English Test Requirement",
    },
    "applied": {
        "name": "Organization Name",
        "role-title": "Role / Position Title",
        "date": "Applied On",
        "url": "Job Posting URL",
        "source": "Application Source",
        "type": "Opportunity Type",
        "deadline": "Posting Deadline",
        "result": "Outcome",
        "result-on": "Outcome Date",
        "application-progress": "Current Stage",
        "number-of-positions": "Open Positions",
        "ai-reliance": "AI Assistance Level",
        "completed-stages": "Completed Stages Notes",
        "remarks": "Notes",
        "cover-letter-text": "Cover Letter Draft",
        "cv-text": "CV / Resume Draft",
    },
}

FRIENDLY_JOB_FIELD_HELP_TEXTS = {
    "to-apply": {
        "name": "Company, university, or organization offering this opportunity.",
        "source": "Where you found this opening (for example LinkedIn, company careers page, or referral).",
    },
    "applied": {
        "name": "Company, university, or organization you applied to.",
        "source": "Where this application came from (for example LinkedIn, referral, or careers page).",
        "remarks": "Use this for interview notes, follow-up reminders, or context you do not want to lose.",
    },
}


def apply_widget_classes(field: forms.Field) -> None:
    if isinstance(field.widget, forms.CheckboxInput):
        field.widget.attrs["class"] = "mt-2 h-5 w-5 rounded border-slate-300 text-sky-600 focus:ring-sky-500"
        return
    if isinstance(field.widget, forms.RadioSelect):
        field.widget.attrs["class"] = "peer sr-only"
        return
    classes = (
        "mt-2 w-full rounded-2xl border border-slate-300/80 bg-white/85 px-4 py-3 "
        "text-sm text-slate-900 shadow-sm outline-none transition "
        "placeholder:text-slate-400 focus:border-sky-500 focus:ring-2 focus:ring-sky-200"
    )
    if isinstance(field.widget, forms.Textarea):
        field.widget.attrs["rows"] = field.widget.attrs.get("rows", 5)
    if isinstance(field.widget, (forms.DateInput, forms.TimeInput, forms.DateTimeInput)):
        field.widget.attrs["data-picker-only"] = "true"
        field.widget.attrs["autocomplete"] = "off"
    field.widget.attrs["class"] = classes


class TableDefinitionForm(forms.ModelForm):
    class Meta:
        model = DynamicTable
        fields = ("name", "description", "show_in_navigation")
        widgets = {"description": forms.Textarea(attrs={"rows": 3})}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            apply_widget_classes(field)


class WorkbookImportForm(forms.Form):
    class ImportMode:
        APPEND = "append"
        REPLACE = "replace"

    IMPORT_MODE_CHOICES = (
        (ImportMode.APPEND, "Append to current data"),
        (ImportMode.REPLACE, "Remove all and import again"),
    )

    workbook = forms.FileField(
        label="Workbook file",
        help_text="Upload an .xlsx file shaped like your Applications workbook.",
    )
    import_mode = forms.ChoiceField(
        label="Import mode",
        choices=IMPORT_MODE_CHOICES,
        initial=ImportMode.APPEND,
        help_text="Append keeps existing records. Remove all clears current records first.",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            apply_widget_classes(field)
        self.fields["workbook"].widget.attrs["accept"] = ".xlsx"

    def clean_workbook(self):
        workbook = self.cleaned_data["workbook"]
        if not workbook.name.lower().endswith(".xlsx"):
            raise forms.ValidationError("Upload an .xlsx workbook.")
        return workbook


class ColumnDefinitionForm(forms.ModelForm):
    class Meta:
        model = DynamicColumn
        fields = (
            "name",
            "data_type",
            "is_required",
            "is_visible",
            "placeholder",
            "help_text",
            "choices_text",
        )
        widgets = {"choices_text": forms.Textarea(attrs={"rows": 3})}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            apply_widget_classes(field)


class AccountSettingsForm(forms.ModelForm):
    class Meta:
        model = User
        fields = (
            "notification_email",
            "notify_enabled",
            "notify_time",
            "auto_fetch_portals",
            "deadline_notification_window_days",
            "chat_history_retention_days",
        )
        widgets = {
            "notify_time": forms.TimeInput(attrs={"type": "time"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            apply_widget_classes(field)


class DashboardBoardForm(forms.ModelForm):
    class Meta:
        model = DashboardBoard
        fields = (
            "name",
            "description",
            "show_on_homepage",
            "autoplay_enabled",
            "autoplay_interval_seconds",
        )
        widgets = {
            "description": forms.Textarea(attrs={"rows": 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["name"].help_text = "This is the label you will see in the dashboard library and homepage tabs."
        self.fields["show_on_homepage"].help_text = "Enable this if the board should appear as an option on the homepage."
        self.fields["autoplay_enabled"].help_text = "Rotate a spotlight highlight across plots automatically on the homepage."
        self.fields["autoplay_interval_seconds"].help_text = "Controls how often the homepage spotlight moves to the next plot."
        for field in self.fields.values():
            apply_widget_classes(field)


class EmailAwareAuthenticationForm(AuthenticationForm):
    error_messages = {
        **AuthenticationForm.error_messages,
        "inactive": "Verify your email before logging in.",
    }

    def confirm_login_allowed(self, user):
        if not user.is_active and getattr(user, "email_verified_at", None) is None:
            raise ValidationError(self.error_messages["inactive"], code="inactive")
        super().confirm_login_allowed(user)


class SignupForm(UserCreationForm):
    full_name = forms.CharField(max_length=120, label="Full name")
    email = forms.EmailField(label="Email")

    class Meta(UserCreationForm.Meta):
        model = User
        fields = ("full_name", "email", "username", "password1", "password2")

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["username"].help_text = "Use letters, numbers, and @/./+/-/_ only."
        for field in self.fields.values():
            apply_widget_classes(field)

    def clean_email(self):
        email = (self.cleaned_data.get("email") or "").strip().lower()
        if User.objects.filter(email__iexact=email).exists():
            raise forms.ValidationError("An account with this email already exists.")
        return email

    def save(self, commit=True):
        user = super().save(commit=False)
        full_name = (self.cleaned_data.get("full_name") or "").strip()
        user.display_name = full_name
        user.email = self.cleaned_data["email"]
        user.notification_email = self.cleaned_data["email"]
        if commit:
            user.save()
        return user


class DynamicRecordForm(forms.Form):
    def __init__(self, *args, table: DynamicTable, user=None, row_key=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.table = table
        self.user = user
        self.row_key = row_key
        self.dashboard_columns_map: dict[str, list[str]] = {}

        for column in table.columns.filter(is_visible=True):
            self.fields[column.slug] = self._build_field(column)

        if table.slug == "dashboard" and user is not None:
            self._configure_dashboard_fields()
        self._normalize_initial_dates()

    @staticmethod
    def _is_date_like_slug(slug: str) -> bool:
        normalized = (slug or "").strip().lower()
        return (
            normalized in {"date", "deadline", "result-on", "created_at", "updated_at"}
            or normalized.endswith("_date")
            or normalized.endswith("-date")
            or normalized.endswith("-on")
        )

    @staticmethod
    def _parse_date_like_value(raw_value):
        if raw_value in {None, ""}:
            return None
        if isinstance(raw_value, datetime):
            return raw_value.date()
        if isinstance(raw_value, date):
            return raw_value
        text = str(raw_value).strip()
        if not text:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(text).date()
        except ValueError:
            return None

    def _normalize_initial_dates(self) -> None:
        for field_name, field in self.fields.items():
            if not isinstance(field, forms.DateField):
                continue
            parsed = self._parse_date_like_value(self.initial.get(field_name))
            if parsed:
                self.initial[field_name] = parsed

    def _build_field(self, column: DynamicColumn) -> forms.Field:
        is_date_like = column.data_type == DynamicColumn.DataType.DATE or self._is_date_like_slug(column.slug)
        if column.data_type == DynamicColumn.DataType.LONG_TEXT:
            field = forms.CharField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                widget=forms.Textarea(attrs={"placeholder": column.placeholder or column.name}),
            )
        elif is_date_like:
            field = forms.DateField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                input_formats=["%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y"],
                widget=forms.DateInput(attrs={"type": "date", "placeholder": column.placeholder}),
            )
        elif column.data_type == DynamicColumn.DataType.STATUS and column.choices_list:
            field = forms.ChoiceField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                choices=[("", "Select an option")] + [(choice, choice) for choice in column.choices_list],
            )
        elif column.data_type == DynamicColumn.DataType.URL:
            field = forms.URLField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                widget=forms.URLInput(attrs={"placeholder": column.placeholder}),
            )
        elif column.data_type == DynamicColumn.DataType.NUMBER:
            field = forms.DecimalField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                decimal_places=0,
                max_digits=10,
                widget=forms.NumberInput(attrs={"placeholder": column.placeholder}),
            )
        else:
            field = forms.CharField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                widget=forms.TextInput(attrs={"placeholder": column.placeholder}),
            )
        apply_widget_classes(field)
        return field

    def _pick_source_table_slug(self, source_choices: list[tuple[str, str]]) -> str:
        if self.is_bound:
            return self.data.get("source-table") or (source_choices[1][0] if len(source_choices) > 1 else "")
        if self.initial.get("source-table"):
            return self.initial["source-table"]
        return source_choices[1][0] if len(source_choices) > 1 else ""

    def _configure_dashboard_fields(self) -> None:
        source_tables = list(
            DynamicTable.objects.filter(owner=self.user).exclude(slug="dashboard").prefetch_related("columns")
        )
        source_choices = [("", "Select source table")] + [(table.slug, table.name) for table in source_tables]
        self.dashboard_columns_map = {
            table.slug: [column.slug for column in table.columns.filter(is_visible=True)]
            for table in source_tables
        }
        source_slug = self._pick_source_table_slug(source_choices)
        column_choices = [("", "No grouping")] + [
            (column_slug, column_slug) for column_slug in self.dashboard_columns_map.get(source_slug, [])
        ]
        metric_choices = [("", "Not needed for count records")] + [
            (column_slug, column_slug) for column_slug in self.dashboard_columns_map.get(source_slug, [])
        ]

        if "source-table" in self.fields:
            self.fields["source-table"] = forms.ChoiceField(
                required=False,
                label=self.fields["source-table"].label,
                help_text=self.fields["source-table"].help_text,
                choices=source_choices,
            )
            apply_widget_classes(self.fields["source-table"])
        if "group-by-column" in self.fields:
            self.fields["group-by-column"] = forms.ChoiceField(
                required=False,
                label=self.fields["group-by-column"].label,
                help_text=self.fields["group-by-column"].help_text,
                choices=column_choices,
            )
            apply_widget_classes(self.fields["group-by-column"])
        if "metric-column" in self.fields:
            self.fields["metric-column"] = forms.ChoiceField(
                required=False,
                label=self.fields["metric-column"].label,
                help_text=self.fields["metric-column"].help_text,
                choices=metric_choices,
            )
            apply_widget_classes(self.fields["metric-column"])


class JobRecordForm(DynamicRecordForm):
    def __init__(self, *args, table: DynamicTable, user=None, row_key=None, **kwargs):
        super().__init__(*args, table=table, user=user, row_key=row_key, **kwargs)
        self._apply_friendly_labels_and_help()
        self.available_tags: list[str] = []
        if table.slug in {"to-apply", "applied"}:
            self._configure_job_state_fields()
        if table.slug not in {"to-apply", "applied"} or user is None:
            return

        existing_tags = []
        if row_key:
            existing_tags = [tag.name for tag in get_application_tags(user, table.slug, row_key)]
        tag_initial = ", ".join(existing_tags)
        self.fields["tags"] = forms.CharField(
            required=False,
            label="Tags",
            initial=tag_initial,
            help_text="Comma-separated. Existing tags can be reused, and new ones are created automatically.",
            widget=forms.TextInput(attrs={"placeholder": "Dummy, Priority"}),
        )
        apply_widget_classes(self.fields["tags"])
        self.available_tags = list(
            ApplicationTag.objects.filter(user=user).order_by("name").values_list("name", flat=True)
        )
        if self.fields["tags"].initial:
            self.initial["tags"] = self.fields["tags"].initial

    def _apply_friendly_labels_and_help(self) -> None:
        table_labels = FRIENDLY_JOB_FIELD_LABELS.get(self.table.slug, {})
        table_help_texts = FRIENDLY_JOB_FIELD_HELP_TEXTS.get(self.table.slug, {})
        for slug, field in self.fields.items():
            if slug in table_labels:
                field.label = table_labels[slug]
            if slug in table_help_texts:
                field.help_text = table_help_texts[slug]

    @staticmethod
    def _ordered_choices(choices: list[tuple[str, str]], preferred_order: list[str]) -> list[tuple[str, str]]:
        normalized_rank = {name.lower(): idx for idx, name in enumerate(preferred_order)}
        ranked = []
        unranked = []
        for value, label in choices:
            text = str(label or value or "").strip()
            rank = normalized_rank.get(text.lower())
            entry = (str(value), str(label))
            if rank is None:
                unranked.append(entry)
            else:
                ranked.append((rank, entry))
        ranked.sort(key=lambda item: item[0])
        return [entry for _, entry in ranked] + unranked

    def _set_state_field_widget(self, field_name: str, preferred_order: list[str]) -> None:
        field = self.fields.get(field_name)
        if not isinstance(field, forms.ChoiceField):
            return
        existing_choices = list(field.choices)
        non_empty = [(value, label) for value, label in existing_choices if str(value).strip()]
        field.choices = self._ordered_choices(non_empty, preferred_order)
        field.widget = forms.RadioSelect(choices=field.choices)
        apply_widget_classes(field)

    def _configure_job_state_fields(self) -> None:
        self._set_state_field_widget("application-progress", FORM_APPLICATION_STAGE_ORDER)
        self._set_state_field_widget("result", APPLICATION_RESULT_ORDER)


class ToApplyForm(JobRecordForm):
    def clean(self):
        cleaned_data = super().clean()
        if not cleaned_data.get("url"):
            self.add_error("url", "Add a URL so you can revisit this opportunity later.")
        return cleaned_data


class AppliedCreateForm(JobRecordForm):
    source_row_key = forms.ChoiceField(
        required=False,
        label="Start from a To Apply entry",
        choices=[("", "Enter a brand-new application")],
    )

    def __init__(self, *args, table: DynamicTable, user=None, source_records=None, row_key=None, **kwargs):
        super().__init__(*args, table=table, user=user, row_key=row_key, **kwargs)
        if "url" in self.fields:
            self.fields["url"].required = False
        choices = [("", "Enter a brand-new application")]
        for record in source_records or []:
            role_type = record.values.get("type")
            label = f"{record.label} ({role_type})" if role_type else record.label
            choices.append((str(record.row_key), label))
        self.fields["source_row_key"].choices = choices
        self.fields["source_row_key"].help_text = (
            "Pick an existing To Apply item to prefill this form and move it into Applied."
        )
        apply_widget_classes(self.fields["source_row_key"])


class AppliedUpdateForm(JobRecordForm):
    HISTORY_STAGE_CHOICES = FORM_APPLICATION_STAGE_ORDER

    @staticmethod
    def history_field_name(stage: str) -> str:
        return f"history_{stage.strip().lower().replace(' ', '_')}_date"

    def __init__(self, *args, table: DynamicTable, user=None, row_key=None, **kwargs):
        super().__init__(*args, table=table, user=user, row_key=row_key, **kwargs)
        if "url" in self.fields:
            self.fields["url"].required = False
        if not user or not row_key:
            return
        events = get_application_history_events(user, "applied", row_key)
        stage_dates: dict[str, date] = {}
        for event in events:
            stage = (event.to_stage or "").strip()
            if not stage or stage not in self.HISTORY_STAGE_CHOICES or stage in stage_dates:
                continue
            stage_dates[stage] = event.occurred_at.date()

        for stage in self.HISTORY_STAGE_CHOICES:
            field_name = self.history_field_name(stage)
            self.fields[field_name] = forms.DateField(
                required=False,
                label=f"{stage} on",
                help_text="Optional. Pick the date this stage was reached.",
                widget=forms.DateInput(attrs={"type": "date"}),
                initial=stage_dates.get(stage),
            )
            apply_widget_classes(self.fields[field_name])
            if stage_dates.get(stage):
                self.initial[field_name] = stage_dates[stage]
