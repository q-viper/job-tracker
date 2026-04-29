from rest_framework import serializers

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
