from rest_framework import serializers
from django.urls import reverse

from .models import DynamicColumn, DynamicTable, PortalJobSuggestion


class DynamicColumnSerializer(serializers.ModelSerializer):
    choices = serializers.SerializerMethodField()

    class Meta:
        model = DynamicColumn
        fields = (
            "id",
            "name",
            "slug",
            "data_type",
            "position",
            "is_required",
            "is_visible",
            "placeholder",
            "help_text",
            "choices",
        )

    def get_choices(self, obj: DynamicColumn) -> list[str]:
        return obj.choices_list


class DynamicTableSerializer(serializers.ModelSerializer):
    columns = DynamicColumnSerializer(many=True, read_only=True)

    class Meta:
        model = DynamicTable
        fields = (
            "id",
            "name",
            "slug",
            "description",
            "display_order",
            "show_in_navigation",
            "columns",
        )


class DisplayCellSerializer(serializers.Serializer):
    label = serializers.CharField(source="column.name")
    slug = serializers.CharField(source="column.slug")
    data_type = serializers.CharField(source="column.data_type")
    value = serializers.CharField()
    display_value = serializers.CharField()
    hyperlink = serializers.CharField(allow_null=True)


class GroupedRecordSerializer(serializers.Serializer):
    row_key = serializers.UUIDField()
    label = serializers.CharField()
    values = serializers.DictField(child=serializers.CharField(allow_blank=True))
    metadata = serializers.DictField()
    created_at = serializers.DateTimeField(allow_null=True)
    updated_at = serializers.DateTimeField(allow_null=True)
    cells = DisplayCellSerializer(many=True)


class RecordWriteSerializer(serializers.Serializer):
    values = serializers.DictField(required=False)


class PromoteToAppliedSerializer(serializers.Serializer):
    source_row_key = serializers.UUIDField(required=False)
    values = serializers.DictField(required=False)


class PortalJobSuggestionSerializer(serializers.ModelSerializer):
    class Meta:
        model = PortalJobSuggestion
        fields = (
            "id",
            "portal_row_key",
            "portal_name",
            "portal_url",
            "title",
            "url",
            "summary",
            "added_to_apply",
            "fetched_at",
        )


def serialize_dashboard_widget(widget) -> dict[str, object]:
    return {
        "row_key": str(widget.row_key),
        "layout_item_id": widget.layout_item_id,
        "title": widget.title,
        "template_key": widget.template_key,
        "chart_type": widget.chart_type,
        "accent": widget.accent,
        "notes": widget.notes,
        "metric_label": widget.metric_label,
        "source_table": widget.source_table,
        "group_by_column": widget.group_by_column,
        "metric_column": widget.metric_column,
        "aggregation": widget.aggregation,
        "filters": widget.filters,
        "labels": widget.labels,
        "values": widget.values,
        "chart_config": widget.chart_config,
        "kpi_value": widget.kpi_value,
        "table_columns": widget.table_columns,
        "show_on_dashboard": widget.show_on_dashboard,
        "sort_order": widget.sort_order,
        "card_size": widget.card_size,
        "chart_height": widget.chart_height,
        "grid_span": widget.grid_span,
        "chart_height_px": widget.chart_height_px,
        "edit_url": reverse("workspace_record_edit", args=["dashboard", widget.row_key]),
    }
