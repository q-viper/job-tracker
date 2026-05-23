from __future__ import annotations

import os
import smtplib
from email.message import EmailMessage
from email.utils import formatdate, make_msgid


def _bool_from_env(name: str, default: bool) -> bool:
    return os.getenv(name, str(default)).lower() in {"1", "true", "yes", "on"}


class SMTPHandler:
    def __init__(
        self,
        *,
        host: str,
        port: int,
        username: str = "",
        password: str = "",
        from_email: str = "",
        use_tls: bool = True,
        use_ssl: bool = False,
        timeout: int = 20,
    ) -> None:
        self.host = host.strip()
        self.port = int(port)
        self.username = username.strip()
        self.password = password.strip()
        self.from_email = from_email.strip() or self.username
        self.use_tls = use_tls
        self.use_ssl = use_ssl
        self.timeout = timeout
        self.server: smtplib.SMTP | smtplib.SMTP_SSL | None = None

    @classmethod
    def is_configured(cls) -> bool:
        return bool(os.getenv("SMTP_HOST", "").strip())

    @classmethod
    def from_env(cls) -> SMTPHandler:
        host = os.getenv("SMTP_HOST", "").strip()
        if not host:
            raise RuntimeError("SMTP_HOST is not configured.")
        return cls(
            host=host,
            port=int(os.getenv("SMTP_PORT", "587")),
            username=os.getenv("SMTP_USERNAME", ""),
            password=os.getenv("SMTP_PASSWORD", ""),
            from_email=os.getenv("SMTP_FROM_EMAIL", ""),
            use_tls=_bool_from_env("SMTP_USE_TLS", True),
            use_ssl=_bool_from_env("SMTP_USE_SSL", False),
        )

    def connect(self) -> SMTPHandler:
        if self.server is not None:
            return self

        if self.use_ssl:
            server = smtplib.SMTP_SSL(self.host, self.port, timeout=self.timeout)
        else:
            server = smtplib.SMTP(self.host, self.port, timeout=self.timeout)
        server.ehlo()
        if self.use_tls and not self.use_ssl:
            server.starttls()
            server.ehlo()
        if self.username and self.password:
            server.login(self.username, self.password)
        self.server = server
        return self

    def close(self) -> None:
        if self.server is None:
            return
        try:
            self.server.quit()
        finally:
            self.server = None

    def send_email(self, recipient: str, subject: str, body: str) -> str | None:
        if not self.from_email:
            raise RuntimeError("SMTP_FROM_EMAIL or SMTP_USERNAME must be configured.")

        from_domain = ""
        if "@" in self.from_email:
            from_domain = self.from_email.rsplit("@", 1)[-1].strip()

        message = EmailMessage()
        message["Subject"] = subject
        message["From"] = self.from_email
        message["To"] = recipient
        message["Date"] = formatdate(localtime=True)
        message["Message-ID"] = make_msgid(domain=from_domain or None)
        message["Auto-Submitted"] = "auto-generated"
        message.set_content(body)
        message_id = str(message.get("Message-ID", "")).strip() or None

        for attempt in range(2):
            self.connect()
            assert self.server is not None
            try:
                refused_recipients = self.server.send_message(message)
                if refused_recipients:
                    recipients = ", ".join(refused_recipients.keys())
                    raise RuntimeError(f"SMTP server refused recipient(s): {recipients}")
                return message_id
            except smtplib.SMTPServerDisconnected:
                # Retry once with a fresh connection when the server dropped a stale session.
                self.close()
                if attempt == 1:
                    raise
        return message_id

    def __enter__(self) -> SMTPHandler:
        return self.connect()

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()
