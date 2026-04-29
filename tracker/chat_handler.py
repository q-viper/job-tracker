from __future__ import annotations

from dataclasses import dataclass

from .models import ChatModelConfig
from .rag import chat_with_rag, extract_text_from_uploaded_file


LEGACY_CHAT_MODEL_SLUGS = ("quassi-chat", "quassi-list", "quassi-rag")

DEFAULT_CHAT_MODELS = (
    {
        "name": "google/flan-t5-small",
        "slug": "google-flan-t5-small",
        "mode": ChatModelConfig.Mode.CHAT,
        "model_name": "google/flan-t5-small",
        "is_default": True,
        "notes": "CPU-friendly default model for grounded chat against your workspace.",
    },
    {
        "name": "google/flan-t5-base",
        "slug": "google-flan-t5-base",
        "mode": ChatModelConfig.Mode.CHAT,
        "model_name": "google/flan-t5-base",
        "is_default": False,
        "notes": "A larger local model option. Better nuance, slower on CPU.",
    },
)


@dataclass
class ChatHandler:
    user: object
    model_slug: str = ""

    @classmethod
    def ensure_default_models(cls) -> None:
        ChatModelConfig.objects.filter(slug__in=LEGACY_CHAT_MODEL_SLUGS).update(enabled=False, is_default=False)
        for definition in DEFAULT_CHAT_MODELS:
            model, created = ChatModelConfig.objects.get_or_create(
                slug=definition["slug"],
                defaults=definition,
            )
            if created:
                continue
            changed_fields: list[str] = []
            for field_name, value in definition.items():
                if getattr(model, field_name) != value:
                    setattr(model, field_name, value)
                    changed_fields.append(field_name)
            if changed_fields:
                model.save(update_fields=changed_fields + ["updated_at"])

    @classmethod
    def available_models(cls) -> list[ChatModelConfig]:
        cls.ensure_default_models()
        return list(ChatModelConfig.objects.filter(enabled=True).order_by("-is_default", "name"))

    @classmethod
    def default_model(cls) -> ChatModelConfig:
        cls.ensure_default_models()
        return (
            ChatModelConfig.objects.filter(enabled=True, is_default=True).order_by("name").first()
            or ChatModelConfig.objects.filter(enabled=True).order_by("name").first()
        )

    def resolve_model(self) -> ChatModelConfig:
        self.ensure_default_models()
        if self.model_slug:
            model = ChatModelConfig.objects.filter(enabled=True, slug=self.model_slug).first()
            if model:
                return model
        default = self.default_model()
        if default is not None:
            return default
        return ChatModelConfig(
            name="google/flan-t5-small",
            slug="google-flan-t5-small",
            mode=ChatModelConfig.Mode.CHAT,
            model_name="google/flan-t5-small",
            enabled=True,
            is_default=True,
        )

    def chat(self, query: str, uploaded_document=None) -> dict:
        chat_model = self.resolve_model()
        document_name, document_text = extract_text_from_uploaded_file(uploaded_document)
        return chat_with_rag(
            self.user,
            query,
            uploaded_document_name=document_name,
            uploaded_document_text=document_text,
            chat_model=chat_model,
        )
