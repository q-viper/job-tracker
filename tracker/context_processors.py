from django.db.utils import OperationalError, ProgrammingError

from .models import ChatMessage, DynamicTable


def navigation_tables(request):
    if not getattr(request, "user", None) or not request.user.is_authenticated:
        return {
            "navigation_tables": [],
            "rag_default_prompts": [],
            "quassi_history": [],
            "quassi_models": [],
            "smtp_configured": False,
        }
    try:
        tables = list(DynamicTable.objects.filter(
            owner=request.user,
            show_in_navigation=True,
        ).order_by("display_order", "name"))
    except (OperationalError, ProgrammingError):
        tables = []

    if "HTTP_HX_REQUEST" in request.META:
        return {
            "navigation_tables": tables,
            "rag_default_prompts": [],
            "quassi_history": [],
            "quassi_models": [],
            "smtp_configured": False,
        }

    try:
        from .rag import RAG_DEFAULT_PROMPTS
    except Exception:
        RAG_DEFAULT_PROMPTS = []
    try:
        from .rag import purge_expired_chat_history

        purge_expired_chat_history(request.user)
        history = list(
            ChatMessage.objects.filter(user=request.user)
            .order_by("-created_at", "-id")[:16]
            .values("role", "content", "attachment_name", "metadata", "created_at")
        )
        history.reverse()
    except (OperationalError, ProgrammingError):
        history = []
    try:
        from .smtp_handler import SMTPHandler

        smtp_configured = SMTPHandler.is_configured()
    except Exception:
        smtp_configured = False
    try:
        from .chat_handler import ChatHandler

        quassi_models = [
            {
                "name": model.name,
                "slug": model.slug,
                "mode": model.mode,
                "is_default": model.is_default,
            }
            for model in ChatHandler.available_models()
        ]
    except Exception:
        quassi_models = []
    return {
        "navigation_tables": tables,
        "rag_default_prompts": RAG_DEFAULT_PROMPTS,
        "quassi_history": history,
        "quassi_models": quassi_models,
        "smtp_configured": smtp_configured,
    }
