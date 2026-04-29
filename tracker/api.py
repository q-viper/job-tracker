from django.shortcuts import get_object_or_404
from rest_framework import status
from rest_framework.parsers import FormParser, JSONParser, MultiPartParser
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from rest_framework.views import APIView

from .chat_handler import ChatHandler
from .models import ChatMessage, DynamicData, DynamicTable, PortalJobSuggestion
from .serializers import (
    DynamicTableSerializer,
    GroupedRecordSerializer,
    PortalJobSuggestionSerializer,
    PromoteToAppliedSerializer,
    RecordWriteSerializer,
)
from .services import (
    add_portal_job_to_to_apply,
    apply_inferred_record_values,
    build_dashboard,
    get_record,
    get_records,
    get_user_table,
    infer_application_type,
    move_to_applied,
    refresh_portal_jobs,
    save_record,
)


class DashboardAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request):
        payload = build_dashboard(request.user)
        return Response(
            {
                "metrics": payload["metrics"],
                "stage_breakdown": payload["stage_breakdown"],
                "upcoming_deadlines": GroupedRecordSerializer(
                    payload["upcoming_deadlines"],
                    many=True,
                ).data,
                "recent_applied": GroupedRecordSerializer(
                    payload["recent_applied"],
                    many=True,
                ).data,
                "custom_cards": GroupedRecordSerializer(
                    payload["custom_cards"],
                    many=True,
                ).data,
            }
        )


class TableListAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request):
        tables = DynamicTable.objects.filter(owner=request.user).prefetch_related("columns")
        return Response(DynamicTableSerializer(tables, many=True).data)


class TableRecordAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request, table_slug: str):
        table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug=table_slug,
        )
        records = get_records(table, request.user)
        return Response(
            {
                "table": DynamicTableSerializer(table).data,
                "records": GroupedRecordSerializer(records, many=True).data,
            }
        )

    def post(self, request, table_slug: str):
        table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug=table_slug,
        )
        serializer = RecordWriteSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        row_key = save_record(
            table,
            request.user,
            serializer.validated_data.get("values", {}),
        )
        record = get_record(table, request.user, row_key)
        return Response(
            GroupedRecordSerializer(record).data,
            status=status.HTTP_201_CREATED,
        )


class TableRecordDetailAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request, table_slug: str, row_key):
        table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug=table_slug,
        )
        record = get_record(table, request.user, row_key)
        if not record:
            return Response(status=status.HTTP_404_NOT_FOUND)
        return Response(GroupedRecordSerializer(record).data)

    def patch(self, request, table_slug: str, row_key):
        table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug=table_slug,
        )
        serializer = RecordWriteSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        row_key = save_record(
            table,
            request.user,
            serializer.validated_data.get("values", {}),
            row_key=row_key,
        )
        record = get_record(table, request.user, row_key)
        return Response(GroupedRecordSerializer(record).data)

    def delete(self, request, table_slug: str, row_key):
        table = get_object_or_404(DynamicTable, owner=request.user, slug=table_slug)
        from .services import delete_record

        delete_record(table, request.user, row_key)
        return Response(status=status.HTTP_204_NO_CONTENT)


class PromoteToAppliedAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request):
        serializer = PromoteToAppliedSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        applied_table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug="applied",
        )
        to_apply_table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug="to-apply",
        )
        row_key = move_to_applied(
            user=request.user,
            applied_table=applied_table,
            to_apply_table=to_apply_table,
            values=serializer.validated_data.get("values", {}),
            source_row_key=serializer.validated_data.get("source_row_key"),
        )
        record = get_record(applied_table, request.user, row_key)
        return Response(
            GroupedRecordSerializer(record).data,
            status=status.HTTP_201_CREATED,
        )


class PortalFetchAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request, row_key):
        portal_table = get_object_or_404(
            DynamicTable.objects.filter(owner=request.user).prefetch_related("columns"),
            slug="portals",
        )
        jobs = refresh_portal_jobs(
            request.user,
            portal_table,
            row_key,
            raw_keywords=request.data.get("keywords"),
        )
        return Response(PortalJobSuggestionSerializer(jobs, many=True).data)


class PortalJobAddToApplyAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request, job_id: int):
        suggestion = get_object_or_404(PortalJobSuggestion, id=job_id, user=request.user)
        row_key = add_portal_job_to_to_apply(request.user, suggestion)
        return Response({"row_key": row_key}, status=status.HTTP_201_CREATED)


class ChatAddToApplyAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request):
        portal_job_id = request.data.get("portal_job_id")
        if portal_job_id:
            suggestion = get_object_or_404(PortalJobSuggestion, id=portal_job_id, user=request.user)
            row_key = add_portal_job_to_to_apply(request.user, suggestion)
            return Response({"row_key": row_key}, status=status.HTTP_201_CREATED)

        title = (request.data.get("title") or "").strip()
        url = (request.data.get("url") or "").strip()
        application_type = (request.data.get("application_type") or "").strip()
        if not url:
            return Response({"detail": "A URL is required to add this entry to To Apply."}, status=status.HTTP_400_BAD_REQUEST)

        to_apply_table = get_user_table(request.user, "to-apply")
        existing = DynamicData.objects.filter(
            table=to_apply_table,
            user=request.user,
            column__slug="url",
            value=url,
        ).first()
        if existing:
            return Response({"row_key": existing.row_key}, status=status.HTTP_200_OK)

        row_key = save_record(
            to_apply_table,
            request.user,
            {
                "role-title": title,
                "url": url,
                "type": application_type or infer_application_type(title or url),
            },
        )
        return Response({"row_key": row_key}, status=status.HTTP_201_CREATED)


class ClearChatHistoryAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request):
        deleted, _ = ChatMessage.objects.filter(user=request.user).delete()
        return Response({"deleted": deleted}, status=status.HTTP_200_OK)


class InferRecordFromUrlAPIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request):
        table_slug = (request.data.get("table_slug") or "").strip()
        url = (request.data.get("url") or "").strip()
        if not table_slug or not url:
            return Response({"detail": "table_slug and url are required."}, status=status.HTTP_400_BAD_REQUEST)
        table = get_user_table(request.user, table_slug)
        values = request.data.get("values")
        if not isinstance(values, dict):
            values = {}
        inferred = apply_inferred_record_values(table, {"url": url, **values})
        return Response({"values": inferred})


class RagChatAPIView(APIView):
    permission_classes = [IsAuthenticated]
    parser_classes = [JSONParser, MultiPartParser, FormParser]

    def post(self, request):
        query = (request.data.get("query") or "").strip()
        if not query:
            return Response({"detail": "Query is required."}, status=status.HTTP_400_BAD_REQUEST)
        handler = ChatHandler(request.user, model_slug=(request.data.get("model_slug") or "").strip())
        return Response(handler.chat(query, uploaded_document=request.FILES.get("document")))
