from __future__ import annotations

from django.contrib.auth.tokens import default_token_generator
from django.urls import reverse
from django.utils import timezone
from django.utils.encoding import force_bytes, force_str
from django.utils.http import urlsafe_base64_decode, urlsafe_base64_encode

from .models import User
from .smtp_handler import SMTPHandler


def build_verification_link(request, user: User) -> str:
    uid = urlsafe_base64_encode(force_bytes(user.pk))
    token = default_token_generator.make_token(user)
    return request.build_absolute_uri(reverse("verify_email", args=[uid, token]))


def send_verification_email(request, user: User) -> str:
    verification_link = build_verification_link(request, user)
    display_name = user.display_name or user.username
    body = (
        f"Hello {display_name},\n\n"
        "Welcome to Job Tracker. Verify your email to activate your account:\n\n"
        f"{verification_link}\n\n"
        "If you did not create this account, you can ignore this email."
    )
    with SMTPHandler.from_env() as handler:
        handler.send_email(user.email, "Verify your Job Tracker account", body)
    user.verification_email_sent_at = timezone.now()
    user.save(update_fields=["verification_email_sent_at"])
    return verification_link


def verify_user_from_token(uidb64: str, token: str) -> User | None:
    try:
        user_id = force_str(urlsafe_base64_decode(uidb64))
        user = User.objects.get(pk=user_id)
    except Exception:
        return None
    if not default_token_generator.check_token(user, token):
        return None
    if user.email_verified_at is None:
        user.email_verified_at = timezone.now()
    if not user.is_active:
        user.is_active = True
    user.save(update_fields=["email_verified_at", "is_active"])
    return user
