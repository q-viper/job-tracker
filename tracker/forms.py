from __future__ import annotations

from django import forms
from django.contrib.auth.forms import AuthenticationForm, UserCreationForm
from django.core.exceptions import ValidationError

from .models import DynamicColumn, DynamicTable, User


def apply_widget_classes(field: forms.Field) -> None:
    if isinstance(field.widget, forms.CheckboxInput):
        field.widget.attrs["class"] = "mt-2 h-5 w-5 rounded border-slate-300 text-sky-600 focus:ring-sky-500"
        return
    classes = (
        "mt-2 w-full rounded-2xl border border-slate-300/80 bg-white/85 px-4 py-3 "
        "text-sm text-slate-900 shadow-sm outline-none transition "
        "placeholder:text-slate-400 focus:border-sky-500 focus:ring-2 focus:ring-sky-200"
    )
    if isinstance(field.widget, forms.Textarea):
        field.widget.attrs["rows"] = field.widget.attrs.get("rows", 5)
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
    def __init__(self, *args, table: DynamicTable, user=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.table = table
        self.user = user
        self.dashboard_columns_map: dict[str, list[str]] = {}

        for column in table.columns.filter(is_visible=True):
            self.fields[column.slug] = self._build_field(column)

        if table.slug == "dashboard" and user is not None:
            self._configure_dashboard_fields()

    def _build_field(self, column: DynamicColumn) -> forms.Field:
        if column.data_type == DynamicColumn.DataType.LONG_TEXT:
            field = forms.CharField(
                required=column.is_required,
                label=column.name,
                help_text=column.help_text,
                widget=forms.Textarea(attrs={"placeholder": column.placeholder or column.name}),
            )
        elif column.data_type == DynamicColumn.DataType.DATE:
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


class ToApplyForm(DynamicRecordForm):
    def clean(self):
        cleaned_data = super().clean()
        if not cleaned_data.get("url"):
            self.add_error("url", "Add a URL so you can revisit this opportunity later.")
        return cleaned_data


class AppliedCreateForm(DynamicRecordForm):
    source_row_key = forms.ChoiceField(
        required=False,
        label="Import from To Apply",
        choices=[("", "Enter a brand-new application")],
    )

    def __init__(self, *args, table: DynamicTable, user=None, source_records=None, **kwargs):
        super().__init__(*args, table=table, user=user, **kwargs)
        choices = [("", "Enter a brand-new application")]
        for record in source_records or []:
            role_type = record.values.get("type")
            label = f"{record.label} ({role_type})" if role_type else record.label
            choices.append((str(record.row_key), label))
        self.fields["source_row_key"].choices = choices
        apply_widget_classes(self.fields["source_row_key"])

    def clean(self):
        cleaned_data = super().clean()
        if not cleaned_data.get("source_row_key") and not cleaned_data.get("url"):
            self.add_error("url", "Provide a URL or choose an existing To Apply item.")
        return cleaned_data
