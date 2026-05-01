from __future__ import annotations

import ipaddress
import json
import os
import re
import shutil
import socket
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse
from uuid import UUID, uuid4

import requests
from bs4 import BeautifulSoup
from django.contrib.auth import get_user_model
from django.db import transaction
from django.utils import timezone
from django.utils.text import slugify
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter

from .models import DynamicColumn, DynamicData, DynamicTable, PortalJobSuggestion

DEFAULT_WORKBOOK_HEADERS = {
    "applied": [
        "Date",
        "URL",
        "Name",
        "Source",
        "Type",
        "Result",
        "Result On",
        "AI Reliance",
        "Completed Stages",
        "Remarks",
    ],
    "portals": ["URL", "Name"],
    "to-apply": ["URL", "Deadline", "Type", "English Test"],
}

DEPRECATED_COLUMN_SLUGS = {"s-n", "sn"}

DEFAULT_PORTAL_KEYWORDS = (
    "PhD",
    "Candidate",
    "Wissenschaft",
    "Wissenshaft",
    "Research",
    "Scientist",
    "Engineer",
    "Developer",
    "Assistant",
    "Postdoc",
    "Intern",
    "Internship",
    "Fellow",
    "AI",
    "ML",
    "Data",
    "Job",
)

PORTAL_NEGATIVE_KEYWORDS = (
    "cookie",
    "privacy",
    "terms",
    "imprint",
    "impressum",
    "contact",
    "about us",
    "linkedin",
    "facebook",
    "instagram",
    "sign in",
    "login",
    "register",
    "subscribe",
    "newsletter",
)

PORTAL_BROWSER_REQUIRED_MARKERS = (
    "access denied",
    "forbidden",
    "verify you are human",
    "please enable javascript",
    "javascript is required",
    "checking your browser",
    "captcha",
    "cloudflare",
)

PORTAL_BROWSER_BINARY_CANDIDATES = {
    "edge": (
        "msedge",
        "microsoft-edge",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/usr/bin/microsoft-edge",
        "/usr/bin/microsoft-edge-stable",
    ),
    "chrome": (
        "chrome",
        "google-chrome",
        "google-chrome-stable",
        "chromium",
        "chromium-browser",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
    ),
}

PORTAL_FETCH_MODES = {"auto", "requests", "browser"}
PORTAL_BROWSER_NAMES = ("edge", "chrome")

DEFAULT_HTTP_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

CHART_TYPE_CHOICES = (
    "KPI",
    "Bar",
    "Horizontal Bar",
    "Line",
    "Area",
    "Pie",
    "Doughnut",
    "Radar",
    "Polar Area",
    "Waterfall",
    "Sankey",
)

AGGREGATION_CHOICES = (
    "Count Records",
    "Count Filled Values",
    "Sum Numbers",
)

CARD_SIZE_CHOICES = (
    "Compact",
    "Standard",
    "Wide",
    "Full",
)

CHART_HEIGHT_CHOICES = (
    "Short",
    "Medium",
    "Tall",
    "XL",
)

TABLE_METADATA = {
    "dashboard": {
        "name": "Dashboard",
        "description": "Custom dashboard widgets, charts, and metrics for this user.",
        "display_order": 5,
        "show_in_navigation": False,
    },
    "to-apply": {
        "name": "To Apply",
        "description": "Opportunities you still want to apply for.",
        "display_order": 10,
        "show_in_navigation": True,
    },
    "applied": {
        "name": "Applied",
        "description": "Submitted applications, progress, and outcomes.",
        "display_order": 20,
        "show_in_navigation": True,
    },
    "portals": {
        "name": "Portals",
        "description": "Reusable portal links and sources to fetch openings from.",
        "display_order": 30,
        "show_in_navigation": True,
    },
}

COLUMN_OVERRIDES = {
    "dashboard": {
        "title": {"data_type": DynamicColumn.DataType.TEXT},
        "chart-type": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": ", ".join(CHART_TYPE_CHOICES),
        },
        "source-table": {"data_type": DynamicColumn.DataType.TEXT},
        "group-by-column": {"data_type": DynamicColumn.DataType.TEXT},
        "metric-column": {"data_type": DynamicColumn.DataType.TEXT},
        "aggregation": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": ", ".join(AGGREGATION_CHOICES),
        },
        "filters": {"data_type": DynamicColumn.DataType.TEXT},
        "limit": {"data_type": DynamicColumn.DataType.NUMBER},
        "metric": {"data_type": DynamicColumn.DataType.TEXT},
        "notes": {"data_type": DynamicColumn.DataType.LONG_TEXT},
        "accent": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": "Sky, Emerald, Amber, Slate",
        },
    },
    "to-apply": {
        "url": {
            "data_type": DynamicColumn.DataType.URL,
            "placeholder": "https://company.example/jobs/role",
        },
        "deadline": {"data_type": DynamicColumn.DataType.DATE},
        "english-test": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": "No, Planned, Submitted, Waived, Required",
        },
    },
    "applied": {
        "date": {"data_type": DynamicColumn.DataType.DATE},
        "url": {
            "data_type": DynamicColumn.DataType.URL,
            "placeholder": "https://company.example/jobs/role",
        },
        "result": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": "Pending, Rejected, Withdrawn, Offer, Accepted",
        },
        "result-on": {"data_type": DynamicColumn.DataType.DATE},
        "ai-reliance": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": "None, Low, Medium, High",
        },
        "completed-stages": {"data_type": DynamicColumn.DataType.LONG_TEXT},
        "remarks": {"data_type": DynamicColumn.DataType.LONG_TEXT},
    },
    "portals": {
        "url": {"data_type": DynamicColumn.DataType.URL},
    },
}

EXTRA_COLUMNS = {
    "dashboard": [
        {
            "name": "Template Key",
            "slug": "template-key",
            "data_type": DynamicColumn.DataType.TEXT,
            "help_text": "Optional system template key used by built-in dashboard widgets.",
        },
        {
            "name": "Chart Type",
            "slug": "chart-type",
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": ", ".join(CHART_TYPE_CHOICES),
            "insert_after": "title",
            "help_text": "Choose the visualization style for this widget.",
        },
        {
            "name": "Source Table",
            "slug": "source-table",
            "data_type": DynamicColumn.DataType.TEXT,
            "help_text": "The table slug to aggregate, such as applied or to-apply.",
        },
        {
            "name": "Group By Column",
            "slug": "group-by-column",
            "data_type": DynamicColumn.DataType.TEXT,
            "help_text": "Use a column slug like type or application-progress.",
        },
        {
            "name": "Metric Column",
            "slug": "metric-column",
            "data_type": DynamicColumn.DataType.TEXT,
            "help_text": "Needed for Sum Numbers or Count Filled Values.",
        },
        {
            "name": "Aggregation",
            "slug": "aggregation",
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": ", ".join(AGGREGATION_CHOICES),
            "help_text": "Choose how values should be aggregated.",
        },
        {
            "name": "Filters",
            "slug": "filters",
            "data_type": DynamicColumn.DataType.TEXT,
            "help_text": "Use key=value pairs separated by semicolons.",
            "placeholder": "type=PhD;result=Rejected",
        },
        {
            "name": "Limit",
            "slug": "limit",
            "data_type": DynamicColumn.DataType.NUMBER,
            "help_text": "Maximum number of groups to plot.",
            "placeholder": "8",
        },
        {
            "name": "Show On Dashboard",
            "slug": "show-on-dashboard",
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": "Show, Hide",
            "help_text": "Choose whether this widget should appear on the dashboard.",
        },
        {
            "name": "Sort Order",
            "slug": "sort-order",
            "data_type": DynamicColumn.DataType.NUMBER,
            "help_text": "Smaller numbers appear first on the dashboard.",
            "placeholder": "10",
        },
        {
            "name": "Card Size",
            "slug": "card-size",
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": ", ".join(CARD_SIZE_CHOICES),
            "help_text": "Controls how wide the widget appears on the dashboard.",
        },
        {
            "name": "Chart Height",
            "slug": "chart-height",
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": ", ".join(CHART_HEIGHT_CHOICES),
            "help_text": "Controls the chart canvas height inside the widget.",
        },
    ],
    "to-apply": [
        {
            "name": "Role / Title",
            "slug": "role-title",
            "data_type": DynamicColumn.DataType.TEXT,
            "insert_after": "url",
            "placeholder": "Research Assistant, Computer Vision",
        },
    ],
    "applied": [
        {
            "name": "Role / Title",
            "slug": "role-title",
            "data_type": DynamicColumn.DataType.TEXT,
            "insert_after": "date",
            "placeholder": "Senior ML Engineer",
        },
        {
            "name": "Application Progress",
            "slug": "application-progress",
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": (
                "Applied, Screening, Interview, Technical Interview, "
                "Assignment, Offer, Rejected, Accepted"
            ),
            "insert_after": "result-on",
        },
        {
            "name": "Cover Letter Text",
            "slug": "cover-letter-text",
            "data_type": DynamicColumn.DataType.LONG_TEXT,
        },
        {
            "name": "CV Text",
            "slug": "cv-text",
            "data_type": DynamicColumn.DataType.LONG_TEXT,
        },
    ],
    "portals": [
        {
            "name": "Fetch Keywords",
            "slug": "fetch-keywords",
            "data_type": DynamicColumn.DataType.LONG_TEXT,
            "help_text": "One keyword per line or comma separated. These terms are used while scraping the portal.",
            "placeholder": ", ".join(DEFAULT_PORTAL_KEYWORDS),
        },
    ],
}

DASHBOARD_WIDGET_TEMPLATES = {
    "applied-job-types": {
        "title": "Applied Job Types",
        "chart-type": "Pie",
        "source-table": "applied",
        "group-by-column": "type",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "8",
        "metric": "Applications",
        "notes": "A quick view of how your applied roles split across job types.",
        "accent": "Sky",
        "show-on-dashboard": "Show",
        "sort-order": "10",
        "card-size": "Wide",
        "chart-height": "Medium",
    },
    "applied-flow": {
        "title": "Applied Flow",
        "chart-type": "Sankey",
        "source-table": "applied",
        "group-by-column": "application-progress",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "12",
        "metric": "Pipeline flow",
        "notes": "From applied, through job type, into the current application stage.",
        "accent": "Amber",
        "show-on-dashboard": "Show",
        "sort-order": "20",
        "card-size": "Full",
        "chart-height": "Tall",
    },
    "applied-progress-pie": {
        "title": "Applied Progress Mix",
        "chart-type": "Doughnut",
        "source-table": "applied",
        "group-by-column": "application-progress",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "10",
        "metric": "Progress",
        "notes": "Breakdown of your current application stages.",
        "accent": "Amber",
        "show-on-dashboard": "Show",
        "sort-order": "30",
        "card-size": "Standard",
        "chart-height": "Medium",
    },
    "applied-outcomes-waterfall": {
        "title": "Applied Outcomes Waterfall",
        "chart-type": "Waterfall",
        "source-table": "applied",
        "group-by-column": "result",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "8",
        "metric": "Outcomes",
        "notes": "Positive outcomes stack upward while rejected or withdrawn outcomes step downward.",
        "accent": "Emerald",
        "show-on-dashboard": "Hide",
        "sort-order": "40",
        "card-size": "Wide",
        "chart-height": "Medium",
    },
    "deadline-count": {
        "title": "Upcoming Deadlines",
        "chart-type": "Bar",
        "source-table": "to-apply",
        "group-by-column": "deadline",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "8",
        "metric": "Deadlines",
        "notes": "Deadlines grouped by date from your To Apply list.",
        "accent": "Rose",
        "show-on-dashboard": "Hide",
        "sort-order": "50",
        "card-size": "Standard",
        "chart-height": "Short",
    },
}

DEFAULT_DASHBOARD_TEMPLATE_KEYS = ("applied-job-types", "applied-flow")

JOB_KEYWORDS = (
    "job",
    "career",
    "position",
    "opening",
    "vacancy",
    "research",
    "scientist",
    "engineer",
    "developer",
    "intern",
    "phd",
    "candidate",
    "postdoc",
    "assistant",
    "fellow",
    "wissenschaft",
    "wissenshaft",
)


@dataclass
class DisplayCell:
    column: DynamicColumn
    value: str
    display_value: str
    hyperlink: str | None


@dataclass
class GroupedRecord:
    row_key: UUID
    table: DynamicTable
    values: dict[str, str]
    metadata: dict[str, dict[str, Any]]
    cells: list[DisplayCell]
    label: str
    created_at: datetime | None
    updated_at: datetime | None


@dataclass
class DashboardWidget:
    row_key: UUID
    title: str
    template_key: str
    chart_type: str
    accent: str
    notes: str
    metric_label: str
    source_table: str
    group_by_column: str
    metric_column: str
    aggregation: str
    filters: str
    labels: list[str]
    values: list[float]
    chart_config: dict[str, Any]
    kpi_value: str
    table_columns: list[str]
    show_on_dashboard: bool
    sort_order: int
    card_size: str
    chart_height: str
    grid_span: int
    chart_height_px: int


def normalize_slug(value: str) -> str:
    cleaned = slugify((value or "").strip())
    return cleaned or "item"


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        cleaned = value.strip()
        if cleaned.lower() in {"", "none", "null", "nan"}:
            return ""
        return cleaned
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_date_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        cleaned = value.strip()
        if cleaned.lower() in {"", "none", "null"}:
            return ""
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(cleaned, fmt).date().isoformat()
            except ValueError:
                continue
        return cleaned
    return normalize_text(value)


def humanize_value(value: str, column: DynamicColumn) -> str:
    if not value:
        return "-"
    if column.data_type == DynamicColumn.DataType.DATE:
        try:
            return datetime.fromisoformat(value).strftime("%d/%m/%Y")
        except ValueError:
            return value
    return value


def looks_like_url(value: str) -> bool:
    parsed = urlparse(value)
    return bool(parsed.scheme and parsed.netloc)


def serialize_form_value(value: Any, column: DynamicColumn) -> str:
    if column.data_type == DynamicColumn.DataType.DATE:
        return normalize_date_value(value)
    return normalize_text(value)


def parse_number(value: Any) -> float | None:
    cleaned = normalize_text(value)
    if not cleaned:
        return None
    cleaned = cleaned.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def split_words(value: str) -> list[str]:
    return [item.strip() for item in re.split(r"[,;\n]+", value) if item.strip()]


def default_portal_keywords_text() -> str:
    return ", ".join(DEFAULT_PORTAL_KEYWORDS)


def normalize_search_text(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    ascii_text = unicodedata.normalize("NFKD", text)
    ascii_text = "".join(character for character in ascii_text if not unicodedata.combining(character))
    ascii_text = re.sub(r"[^0-9a-zA-Z]+", " ", ascii_text.lower())
    return re.sub(r"\s+", " ", ascii_text).strip()


def normalized_keyword_variants(keyword: str) -> list[str]:
    normalized_keyword = normalize_search_text(keyword)
    if not normalized_keyword:
        return []
    variants = [normalized_keyword]
    raw_parts = [normalize_search_text(part) for part in re.split(r"[\/|&]+", normalize_text(keyword))]
    variants.extend(part for part in raw_parts if part)
    if " " not in normalized_keyword and len(normalized_keyword) >= 3:
        if not normalized_keyword.endswith("s"):
            variants.append(f"{normalized_keyword}s")
            variants.append(f"{normalized_keyword}es")
        if normalized_keyword.endswith("y") and len(normalized_keyword) > 1:
            variants.append(f"{normalized_keyword[:-1]}ies")
    expanded_variants: list[str] = []
    for variant in variants:
        expanded_variants.append(variant)
        if " " not in variant and len(variant) >= 3:
            if not variant.endswith("s"):
                expanded_variants.append(f"{variant}s")
                expanded_variants.append(f"{variant}es")
            if variant.endswith("y") and len(variant) > 1:
                expanded_variants.append(f"{variant[:-1]}ies")
    return list(dict.fromkeys(expanded_variants))


def matches_normalized_keyword(normalized_text: str, keyword: str) -> bool:
    if not normalized_text:
        return False
    padded_text = f" {normalized_text} "
    return any(f" {variant} " in padded_text for variant in normalized_keyword_variants(keyword))


def matches_any_keyword(text: str, keywords: list[str] | tuple[str, ...]) -> bool:
    normalized_text = normalize_search_text(text)
    return any(matches_normalized_keyword(normalized_text, keyword) for keyword in keywords)


def semantic_hex(value: str, *, alpha: float | None = None) -> str:
    normalized = normalize_text(value).lower()
    color = "#64748b"
    if normalized in {"rejected", "withdrawn"}:
        color = "#f43f5e"
    elif normalized in {"offer", "accepted"}:
        color = "#10b981"
    elif "interview" in normalized:
        color = "#f59e0b"
    elif normalized in {"screening", "assignment"}:
        color = "#8b5cf6"
    elif normalized in {"applied", "pending"}:
        color = "#0ea5e9"
    elif normalized in {"phd"}:
        color = "#2563eb"
    elif normalized in {"postdoc"}:
        color = "#14b8a6"
    elif normalized in {"internship", "intern"}:
        color = "#f97316"
    elif normalized in {"job"}:
        color = "#475569"
    if alpha is None:
        return color
    color = color.lstrip("#")
    red, green, blue = tuple(int(color[index : index + 2], 16) for index in (0, 2, 4))
    return f"rgba({red}, {green}, {blue}, {alpha})"


def label_color_palette(labels: list[str]) -> list[str]:
    return [semantic_hex(label, alpha=0.78) for label in labels]


def make_unique_table_slug(user, name: str) -> str:
    base_slug = normalize_slug(name)
    candidate = base_slug
    counter = 2
    while DynamicTable.objects.filter(owner=user, slug=candidate).exists():
        candidate = f"{base_slug}-{counter}"
        counter += 1
    return candidate


def make_unique_column_slug(table: DynamicTable, name: str) -> str:
    base_slug = normalize_slug(name)
    candidate = base_slug
    counter = 2
    while DynamicColumn.objects.filter(table=table, slug=candidate).exists():
        candidate = f"{base_slug}-{counter}"
        counter += 1
    return candidate


def validate_safe_external_url(url: str) -> None:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError("Only http and https URLs are allowed.")
    host = (parsed.hostname or "").strip().lower()
    if not host:
        raise ValueError("The URL is missing a host.")
    if host in {"localhost", "127.0.0.1", "::1"} or host.endswith(".local"):
        raise ValueError("Local or loopback URLs are not allowed.")
    try:
        addresses = {info[-1][0] for info in socket.getaddrinfo(host, None)}
    except socket.gaierror as exc:
        raise ValueError("The URL host could not be resolved.") from exc
    for address in addresses:
        ip = ipaddress.ip_address(address)
        if (
            ip.is_private
            or ip.is_loopback
            or ip.is_link_local
            or ip.is_multicast
            or ip.is_reserved
            or ip.is_unspecified
        ):
            raise ValueError("Private or unsafe network targets are not allowed.")


def safe_get_url(url: str, timeout: int = 15) -> requests.Response:
    validate_safe_external_url(url)
    response = requests.get(
        url,
        timeout=timeout,
        headers={
            "User-Agent": DEFAULT_HTTP_USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Upgrade-Insecure-Requests": "1",
        },
    )
    response.raise_for_status()
    return response


def env_flag(name: str, default: bool) -> bool:
    raw_value = os.getenv(name)
    if raw_value is None:
        return default
    return raw_value.strip().lower() in {"1", "true", "yes", "on"}


def env_int(name: str, default: int, minimum: int = 0) -> int:
    try:
        return max(int(os.getenv(name, str(default))), minimum)
    except (TypeError, ValueError):
        return default


def portal_fetch_mode() -> str:
    configured = (os.getenv("PORTAL_FETCH_MODE", "auto") or "auto").strip().lower()
    return configured if configured in PORTAL_FETCH_MODES else "auto"


def portal_browser_name() -> str:
    configured = (os.getenv("PORTAL_BROWSER", "auto") or "auto").strip().lower()
    return configured if configured == "auto" or configured in PORTAL_BROWSER_NAMES else "auto"


def portal_browser_headless() -> bool:
    return env_flag("PORTAL_BROWSER_HEADLESS", True)


def portal_browser_verbose_logs() -> bool:
    return env_flag("PORTAL_BROWSER_VERBOSE_LOGS", False)


def portal_browser_timeout_seconds() -> int:
    return env_int("PORTAL_BROWSER_TIMEOUT_SECONDS", 30, minimum=5)


def find_portal_browser_binary(browser_name: str) -> str | None:
    override = (os.getenv("PORTAL_BROWSER_BINARY", "") or "").strip()
    if override:
        return override
    for candidate in PORTAL_BROWSER_BINARY_CANDIDATES.get(browser_name, ()):
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
        if Path(candidate).exists():
            return str(Path(candidate))
    return None


def portal_browser_candidates() -> list[str]:
    configured = portal_browser_name()
    if configured in PORTAL_BROWSER_NAMES:
        return [configured]

    installed = [name for name in PORTAL_BROWSER_NAMES if find_portal_browser_binary(name)]
    if installed:
        return installed
    return list(PORTAL_BROWSER_NAMES)


def portal_html_looks_blocked(html: str) -> bool:
    lowered = normalize_text(BeautifulSoup(html, "html.parser").get_text(" ", strip=True)).lower()
    return any(marker in lowered for marker in PORTAL_BROWSER_REQUIRED_MARKERS)


def should_retry_portal_with_browser(exc: Exception) -> bool:
    if isinstance(exc, requests.exceptions.SSLError):
        return True
    if not isinstance(exc, requests.HTTPError):
        return False
    status_code = getattr(getattr(exc, "response", None), "status_code", None)
    return status_code in {401, 403, 429}


def fetch_portal_html_in_browser(url: str, timeout: int | None = None) -> str:
    validate_safe_external_url(url)
    active_timeout = timeout or portal_browser_timeout_seconds()
    last_error: Exception | None = None

    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as expected
        from selenium.webdriver.support.ui import WebDriverWait
    except ImportError as exc:
        raise RuntimeError(
            "Selenium is not installed, so browser-backed portal scraping is unavailable."
        ) from exc

    for browser_name in portal_browser_candidates():
        driver = None
        try:
            service = None
            if browser_name == "edge":
                from selenium.webdriver.edge.options import Options as EdgeOptions
                from selenium.webdriver.edge.service import Service as EdgeService

                options = EdgeOptions()
                browser_binary = find_portal_browser_binary("edge")
                if browser_binary:
                    options.binary_location = browser_binary
                service = EdgeService(log_output=None if portal_browser_verbose_logs() else os.devnull)
                driver_factory = webdriver.Edge
            else:
                from selenium.webdriver.chrome.options import Options as ChromeOptions
                from selenium.webdriver.chrome.service import Service as ChromeService

                options = ChromeOptions()
                browser_binary = find_portal_browser_binary("chrome")
                if browser_binary:
                    options.binary_location = browser_binary
                service = ChromeService(log_output=None if portal_browser_verbose_logs() else os.devnull)
                driver_factory = webdriver.Chrome

            if portal_browser_headless():
                options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-logging")
            options.add_argument("--log-level=3")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--no-sandbox")
            options.add_argument("--window-size=1600,1200")
            options.add_argument(f"--user-agent={DEFAULT_HTTP_USER_AGENT}")
            options.add_experimental_option("excludeSwitches", ["enable-logging"])
            options.page_load_strategy = "eager"

            driver = driver_factory(options=options, service=service)
            driver.set_page_load_timeout(active_timeout)
            driver.get(url)
            WebDriverWait(driver, active_timeout).until(expected.presence_of_element_located((By.TAG_NAME, "body")))
            WebDriverWait(driver, min(active_timeout, 5)).until(
                lambda current: bool(current.find_elements(By.CSS_SELECTOR, "a[href]"))
                or current.execute_script("return document.readyState") == "complete"
            )
            return driver.page_source
        except Exception as exc:
            last_error = exc
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass

    raise RuntimeError(
        f"Could not start a Selenium browser session for portal scraping: {last_error}"
    )


def guess_title_from_url(url: str) -> str:
    parsed = urlparse(url)
    path_part = parsed.path.rstrip("/").split("/")[-1] if parsed.path else ""
    cleaned = re.sub(r"[-_]+", " ", path_part).strip()
    return cleaned.title() if cleaned else ""


def extract_host_name(url: str) -> str:
    host = urlparse(url).hostname or ""
    host = host.replace("www.", "")
    return host.split(".")[0].replace("-", " ").title() if host else ""


def infer_fields_from_url(url: str) -> dict[str, str]:
    inferred = {
        "source": extract_host_name(url),
        "name": extract_host_name(url),
    }
    try:
        response = safe_get_url(url, timeout=12)
        soup = BeautifulSoup(response.text, "html.parser")
    except Exception:
        fallback_title = guess_title_from_url(url)
        if fallback_title:
            inferred["role-title"] = fallback_title
        return inferred

    page_title = ""
    employment_type = ""
    organization_name = ""
    long_description = ""
    for script_tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            script_payload = json.loads(script_tag.string or "")
        except Exception:
            continue
        payload_items = script_payload if isinstance(script_payload, list) else [script_payload]
        for item in payload_items:
            if not isinstance(item, dict):
                continue
            item_type = normalize_text(item.get("@type") or item.get("type")).lower()
            if item_type != "jobposting":
                continue
            page_title = normalize_text(item.get("title")) or page_title
            employment_type = normalize_text(item.get("employmentType")) or employment_type
            organization = item.get("hiringOrganization")
            if isinstance(organization, dict):
                organization_name = normalize_text(organization.get("name")) or organization_name
            long_description = BeautifulSoup(
                normalize_text(item.get("description")),
                "html.parser",
            ).get_text(" ", strip=True)[:600]
            break
        if page_title:
            break

    for candidate in (
        soup.find("meta", property="og:title"),
        soup.find("meta", attrs={"name": "twitter:title"}),
    ):
        if candidate and candidate.get("content"):
            page_title = normalize_text(candidate["content"])
            break
    if not page_title:
        heading = soup.find("h1")
        if heading:
            page_title = normalize_text(heading.get_text(" ", strip=True))
    if not page_title and soup.title and soup.title.string:
        page_title = normalize_text(soup.title.string)

    site_name = ""
    site_tag = soup.find("meta", property="og:site_name")
    if site_tag and site_tag.get("content"):
        site_name = normalize_text(site_tag["content"])

    if page_title:
        role_title = re.split(r"\s[\|\-]\s", page_title)[0].strip()
        if role_title:
            inferred["role-title"] = role_title
        if not site_name:
            parts = re.split(r"\s[\|\-]\s", page_title)
            if len(parts) > 1:
                site_name = parts[-1].strip()

    if organization_name:
        site_name = organization_name
    if site_name:
        inferred["name"] = site_name
        inferred["source"] = site_name

    description_tag = soup.find("meta", attrs={"name": "description"})
    if description_tag and description_tag.get("content"):
        inferred["remarks"] = normalize_text(description_tag["content"])[:400]
    elif long_description:
        inferred["remarks"] = long_description[:400]

    if employment_type and "type" not in inferred:
        inferred["type"] = employment_type

    inferred_type_source = normalize_text(inferred.get("role-title")) or normalize_text(page_title)
    if inferred_type_source and not normalize_text(inferred.get("type")):
        inferred["type"] = infer_application_type(inferred_type_source)

    return inferred


def fallback_specs() -> dict[str, dict[str, Any]]:
    return {
        "dashboard": {
            "sheet_name": None,
            "source_headers": [],
            "columns": [
                {"name": "Title", "slug": "title"},
                {"name": "Chart Type", "slug": "chart-type"},
                {"name": "Source Table", "slug": "source-table"},
                {"name": "Group By Column", "slug": "group-by-column"},
                {"name": "Metric Column", "slug": "metric-column"},
                {"name": "Aggregation", "slug": "aggregation"},
                {"name": "Filters", "slug": "filters"},
                {"name": "Limit", "slug": "limit"},
                {"name": "Metric", "slug": "metric"},
                {"name": "Notes", "slug": "notes"},
                {"name": "Accent", "slug": "accent"},
            ],
        },
        **{
            slug: {
                "sheet_name": TABLE_METADATA[slug]["name"],
                "source_headers": headers,
                "columns": [{"name": header, "slug": normalize_slug(header)} for header in headers],
            }
            for slug, headers in DEFAULT_WORKBOOK_HEADERS.items()
        },
    }


def build_specs(excel_path: str | Path | None = None) -> dict[str, dict[str, Any]]:
    specs = fallback_specs()
    workbook_path = Path(excel_path) if excel_path else None
    if workbook_path and workbook_path.exists():
        workbook = load_workbook(workbook_path)
        for sheet in workbook.worksheets:
            table_slug = normalize_slug(sheet.title)
            headers: list[str] = []
            seen: set[str] = set()
            for cell in sheet[1]:
                header = normalize_text(cell.value)
                if not header:
                    continue
                header_slug = normalize_slug(header)
                if header_slug in seen or header_slug in DEPRECATED_COLUMN_SLUGS:
                    continue
                headers.append(header)
                seen.add(header_slug)
            specs[table_slug] = {
                "sheet_name": sheet.title,
                "source_headers": headers,
                "columns": [{"name": header, "slug": normalize_slug(header)} for header in headers],
            }

    for table_slug, spec in specs.items():
        meta = TABLE_METADATA.get(
            table_slug,
            {
                "name": spec.get("sheet_name") or table_slug.replace("-", " ").title(),
                "description": f"{table_slug.replace('-', ' ').title()} data",
                "display_order": 100,
                "show_in_navigation": True,
            },
        )
        columns: list[dict[str, Any]] = []
        for column in spec["columns"]:
            overrides = COLUMN_OVERRIDES.get(table_slug, {}).get(column["slug"], {})
            columns.append(
                {
                    "name": column["name"],
                    "slug": column["slug"],
                    "data_type": overrides.get("data_type", DynamicColumn.DataType.TEXT),
                    "placeholder": overrides.get("placeholder", ""),
                    "help_text": overrides.get("help_text", ""),
                    "choices_text": overrides.get("choices_text", ""),
                    "is_required": False,
                    "is_visible": overrides.get("is_visible", True),
                }
            )
        for extra in EXTRA_COLUMNS.get(table_slug, []):
            if any(existing["slug"] == extra["slug"] for existing in columns):
                continue
            insert_after = extra.get("insert_after")
            if insert_after:
                insert_at = next(
                    (index + 1 for index, existing in enumerate(columns) if existing["slug"] == insert_after),
                    len(columns),
                )
                columns.insert(insert_at, dict(extra))
            else:
                columns.append(dict(extra))
        for position, column in enumerate(columns, start=1):
            column["position"] = position
            column.setdefault("data_type", DynamicColumn.DataType.TEXT)
            column.setdefault("placeholder", "")
            column.setdefault("help_text", "")
            column.setdefault("choices_text", "")
            column.setdefault("is_required", False)
            column.setdefault("is_visible", True)
        spec.update(
            {
                "name": meta["name"],
                "description": meta["description"],
                "display_order": meta["display_order"],
                "show_in_navigation": meta["show_in_navigation"],
                "columns": columns,
            }
        )
    return specs


@transaction.atomic
def ensure_user_workspace(user, excel_path: str | Path | None = None) -> dict[str, DynamicTable]:
    specs = build_specs(excel_path)
    tables: dict[str, DynamicTable] = {}
    for table_slug, spec in specs.items():
        table, _ = DynamicTable.objects.get_or_create(
            owner=user,
            slug=table_slug,
            defaults={
                "name": spec["name"],
                "description": spec["description"],
                "display_order": spec["display_order"],
                "show_in_navigation": spec["show_in_navigation"],
            },
        )
        changed_fields: list[str] = []
        for field_name in ("name", "description", "display_order", "show_in_navigation"):
            new_value = spec[field_name]
            if getattr(table, field_name) != new_value:
                setattr(table, field_name, new_value)
                changed_fields.append(field_name)
        if changed_fields:
            table.save(update_fields=changed_fields + ["updated_at"])

        for column_spec in spec["columns"]:
            column, _ = DynamicColumn.objects.get_or_create(
                table=table,
                slug=column_spec["slug"],
                defaults={
                    "name": column_spec["name"],
                    "data_type": column_spec["data_type"],
                    "position": column_spec["position"],
                    "is_required": column_spec["is_required"],
                    "is_visible": column_spec["is_visible"],
                    "placeholder": column_spec["placeholder"],
                    "help_text": column_spec["help_text"],
                    "choices_text": column_spec["choices_text"],
                },
            )
            column_changes: list[str] = []
            for field_name in (
                "name",
                "data_type",
                "position",
                "is_required",
                "is_visible",
                "placeholder",
                "help_text",
                "choices_text",
            ):
                new_value = column_spec.get(field_name, getattr(column, field_name))
                if getattr(column, field_name) != new_value:
                    setattr(column, field_name, new_value)
                    column_changes.append(field_name)
            if column_changes:
                column.save(update_fields=column_changes + ["updated_at"])

        deprecated_columns = DynamicColumn.objects.filter(table=table, slug__in=DEPRECATED_COLUMN_SLUGS)
        if deprecated_columns.exists():
            deprecated_columns.delete()

        tables[table_slug] = table
    return tables


def get_user_table(user, slug: str) -> DynamicTable:
    return DynamicTable.objects.prefetch_related("columns").get(owner=user, slug=slug)


def count_records(table: DynamicTable) -> int:
    return DynamicData.objects.filter(table=table, user=table.owner).values("row_key").distinct().count()


def build_label(values: dict[str, str]) -> str:
    for key in ("title", "role-title", "name", "url"):
        value = values.get(key)
        if value:
            return value
    return "Untitled record"


def default_sort_field_for_table(table: DynamicTable) -> str:
    columns = list(table.columns.filter(is_visible=True))
    by_slug = {column.slug: column for column in columns}
    for slug in ("deadline", "date"):
        column = by_slug.get(slug)
        if column and column.data_type == DynamicColumn.DataType.DATE:
            return slug
    first_date_column = next(
        (column.slug for column in columns if column.data_type == DynamicColumn.DataType.DATE),
        None,
    )
    return first_date_column or "updated_at"


def default_sort_direction_for_field(table: DynamicTable, sort_field: str) -> str:
    if sort_field == "deadline":
        return "asc"
    if sort_field in {"created_at", "updated_at"}:
        return "desc"
    column = table.columns.filter(is_visible=True, slug=sort_field).first()
    if column and column.data_type == DynamicColumn.DataType.DATE:
        return "desc"
    return "desc"


def record_sort_value(
    record: GroupedRecord,
    sort_field: str,
    *,
    column_lookup: dict[str, DynamicColumn] | None = None,
) -> tuple[int, int, Any]:
    if sort_field == "label":
        return (0, 2, normalize_text(record.label).lower())
    if sort_field in {"created_at", "updated_at"}:
        timestamp = getattr(record, sort_field)
        return (0, 0, timestamp.timestamp()) if timestamp else (1, 99, "")

    raw_value = normalize_text(record.values.get(sort_field))
    if not raw_value:
        return (1, 99, "")

    column = (column_lookup or {}).get(sort_field)
    if column and column.data_type == DynamicColumn.DataType.DATE:
        try:
            return (0, 0, datetime.fromisoformat(raw_value).timestamp())
        except ValueError:
            return (1, 99, "")
    if column and column.data_type == DynamicColumn.DataType.NUMBER:
        numeric_value = parse_number(raw_value)
        if numeric_value is not None:
            return (0, 1, numeric_value)

    numeric_value = parse_number(raw_value)
    if numeric_value is not None:
        return (0, 1, numeric_value)
    try:
        return (0, 0, datetime.fromisoformat(raw_value).timestamp())
    except ValueError:
        pass
    return (0, 2, raw_value.lower())


def sort_records_with_field(
    records: list[GroupedRecord],
    table: DynamicTable,
    *,
    sort_field: str,
    direction: str,
) -> list[GroupedRecord]:
    column_lookup = {
        column.slug: column
        for column in table.columns.filter(is_visible=True)
    }
    sortable_records: list[tuple[tuple[int, int, Any], GroupedRecord]] = []
    missing_records: list[GroupedRecord] = []

    for record in records:
        sort_key = record_sort_value(record, sort_field, column_lookup=column_lookup)
        if sort_key[0]:
            missing_records.append(record)
        else:
            sortable_records.append((sort_key, record))

    buckets: dict[int, list[tuple[Any, GroupedRecord]]] = {}
    for sort_key, record in sortable_records:
        buckets.setdefault(sort_key[1], []).append((sort_key[2], record))

    ordered_records: list[GroupedRecord] = []
    for type_rank in sorted(buckets):
        bucket = sorted(
            buckets[type_rank],
            key=lambda item: item[0],
            reverse=direction == "desc",
        )
        ordered_records.extend(record for _, record in bucket)
    return ordered_records + missing_records


def sort_records(table: DynamicTable, records: list[GroupedRecord]) -> list[GroupedRecord]:
    sort_field = default_sort_field_for_table(table)
    direction = default_sort_direction_for_field(table, sort_field)
    return sort_records_with_field(
        records,
        table,
        sort_field=sort_field,
        direction=direction,
    )


def get_records(table: DynamicTable, user) -> list[GroupedRecord]:
    entries = list(
        DynamicData.objects.filter(table=table, user=user)
        .select_related("column")
        .order_by("row_key", "column__position")
    )
    columns = list(table.columns.filter(is_visible=True))
    grouped_values: dict[UUID, dict[str, str]] = defaultdict(dict)
    grouped_metadata: dict[UUID, dict[str, dict[str, Any]]] = defaultdict(dict)
    grouped_times: dict[UUID, dict[str, datetime | None]] = defaultdict(
        lambda: {"created_at": None, "updated_at": None}
    )

    for entry in entries:
        grouped_values[entry.row_key][entry.column.slug] = entry.value
        grouped_metadata[entry.row_key][entry.column.slug] = entry.metadata
        current = grouped_times[entry.row_key]
        current["created_at"] = (
            entry.created_at
            if current["created_at"] is None
            else min(current["created_at"], entry.created_at)
        )
        current["updated_at"] = (
            entry.updated_at
            if current["updated_at"] is None
            else max(current["updated_at"], entry.updated_at)
        )

    records: list[GroupedRecord] = []
    for row_key, values in grouped_values.items():
        metadata = grouped_metadata[row_key]
        cells: list[DisplayCell] = []
        for column in columns:
            raw_value = values.get(column.slug, "")
            cell_metadata = metadata.get(column.slug, {})
            hyperlink = cell_metadata.get("hyperlink")
            if not hyperlink and column.data_type == DynamicColumn.DataType.URL and looks_like_url(raw_value):
                hyperlink = raw_value
            cells.append(
                DisplayCell(
                    column=column,
                    value=raw_value,
                    display_value=humanize_value(raw_value, column),
                    hyperlink=hyperlink,
                )
            )
        records.append(
            GroupedRecord(
                row_key=row_key,
                table=table,
                values=values,
                metadata=metadata,
                cells=cells,
                label=build_label(values),
                created_at=grouped_times[row_key]["created_at"],
                updated_at=grouped_times[row_key]["updated_at"],
            )
        )
    return sort_records(table, records)


def get_record(table: DynamicTable, user, row_key: UUID | str) -> GroupedRecord | None:
    target = UUID(str(row_key))
    for record in get_records(table, user):
        if record.row_key == target:
            return record
    return None


def infer_application_type(title: str) -> str:
    lowered = title.lower()
    if "phd" in lowered:
        return "PhD"
    if "postdoc" in lowered:
        return "Postdoc"
    if "intern" in lowered:
        return "Internship"
    return "Job"


def apply_inferred_record_values(
    table: DynamicTable,
    values: dict[str, Any],
    existing_values: dict[str, str] | None = None,
) -> dict[str, Any]:
    merged_values = dict(existing_values or {})
    merged_values.update(values)
    url_value = normalize_text(merged_values.get("url"))
    if not url_value:
        return merged_values

    fields_to_fill = [
        field_name
        for field_name in ("role-title", "name", "source", "remarks")
        if not normalize_text(merged_values.get(field_name))
    ]
    if fields_to_fill:
        inferred = infer_fields_from_url(url_value)
        for field_name in fields_to_fill:
            if field_name in inferred and not normalize_text(merged_values.get(field_name)):
                merged_values[field_name] = inferred[field_name]

    if table.slug in {"to-apply", "applied"} and not normalize_text(merged_values.get("type")):
        title = normalize_text(merged_values.get("role-title")) or normalize_text(merged_values.get("name"))
        if title:
            merged_values["type"] = infer_application_type(title)

    if table.slug == "portals" and not normalize_text(merged_values.get("fetch-keywords")):
        merged_values["fetch-keywords"] = default_portal_keywords_text()

    return merged_values


def save_record(
    table: DynamicTable,
    user,
    values: dict[str, Any],
    row_key: UUID | str | None = None,
    metadata_map: dict[str, dict[str, Any]] | None = None,
) -> UUID:
    if table.owner_id != user.id:
        raise PermissionError("You can only edit your own tables.")

    row_uuid = UUID(str(row_key)) if row_key else uuid4()
    metadata_map = metadata_map or {}
    existing_record = get_record(table, user, row_uuid) if row_key else None
    prepared_values = apply_inferred_record_values(
        table,
        values,
        existing_values=existing_record.values if existing_record else None,
    )

    columns = list(table.columns.all())
    by_slug = {column.slug: column for column in columns}
    existing_entries = {
        entry.column.slug: entry
        for entry in DynamicData.objects.filter(table=table, user=user, row_key=row_uuid).select_related("column")
    }

    with transaction.atomic():
        for slug, column in by_slug.items():
            raw_value = prepared_values.get(slug, "")
            serialized = serialize_form_value(raw_value, column)
            existing = existing_entries.get(slug)
            cell_metadata = metadata_map.get(slug, existing.metadata if existing else {})
            if not serialized:
                if existing:
                    existing.delete()
                continue
            if existing:
                existing.value = serialized
                existing.metadata = cell_metadata
                existing.save(update_fields=["value", "metadata", "updated_at"])
            else:
                DynamicData.objects.create(
                    table=table,
                    column=column,
                    user=user,
                    row_key=row_uuid,
                    value=serialized,
                    metadata=cell_metadata,
                )
    return row_uuid


def save_partial_record(table: DynamicTable, user, row_key: UUID | str, values: dict[str, Any]) -> UUID:
    record = get_record(table, user, row_key)
    if not record:
        raise ValueError("Record not found.")
    merged = dict(record.values)
    merged.update(values)
    return save_record(table, user, merged, row_key=row_key, metadata_map=record.metadata)


def delete_record(table: DynamicTable, user, row_key: UUID | str) -> None:
    if table.owner_id != user.id:
        raise PermissionError("You can only delete records from your own tables.")
    DynamicData.objects.filter(table=table, user=user, row_key=row_key).delete()


def extract_cell_payload(
    column: DynamicColumn,
    cell: Cell,
) -> tuple[str, dict[str, Any], dict[str, str]]:
    metadata: dict[str, Any] = {}
    companion_values: dict[str, str] = {}
    hyperlink = cell.hyperlink.target if cell.hyperlink else None
    raw_value = cell.value

    if column.data_type == DynamicColumn.DataType.URL:
        display_value = normalize_text(raw_value)
        if hyperlink:
            metadata["hyperlink"] = hyperlink
            value = hyperlink
            if display_value and display_value != hyperlink and not looks_like_url(display_value):
                companion_values["role-title"] = display_value
        else:
            value = display_value
    elif column.data_type == DynamicColumn.DataType.DATE:
        value = normalize_date_value(raw_value)
    else:
        value = normalize_text(raw_value)

    return value, metadata, companion_values


def import_workbook_for_user(
    excel_path: str | Path,
    user,
    replace: bool = False,
) -> dict[str, int]:
    workbook_path = Path(excel_path)
    specs = build_specs(workbook_path)
    tables = ensure_user_workspace(user, workbook_path)
    workbook = load_workbook(workbook_path)
    imported_counts: dict[str, int] = {}

    with transaction.atomic():
        if replace:
            DynamicData.objects.filter(table__owner=user, user=user).delete()
            PortalJobSuggestion.objects.filter(user=user).delete()

        for sheet in workbook.worksheets:
            table_slug = normalize_slug(sheet.title)
            table = tables[table_slug]
            sheet_column_map: list[DynamicColumn | None] = []
            seen_headers: set[str] = set()
            for cell in sheet[1]:
                header_slug = normalize_slug(normalize_text(cell.value))
                if not header_slug or header_slug in seen_headers or header_slug in DEPRECATED_COLUMN_SLUGS:
                    sheet_column_map.append(None)
                    continue
                seen_headers.add(header_slug)
                sheet_column_map.append(table.columns.filter(slug=header_slug).first())
            extra_columns = {column.slug: column for column in table.columns.all()}
            imported_rows = 0

            for row in sheet.iter_rows(min_row=2):
                row_key = uuid4()
                row_payload: dict[str, str] = {}
                row_metadata: dict[str, dict[str, Any]] = {}

                for index, cell in enumerate(row):
                    if index >= len(sheet_column_map):
                        break
                    column = sheet_column_map[index]
                    if column is None:
                        continue
                    value, metadata, companion_values = extract_cell_payload(column, cell)
                    if value:
                        row_payload[column.slug] = value
                    if metadata:
                        row_metadata[column.slug] = metadata
                    for companion_slug, companion_value in companion_values.items():
                        if companion_value and companion_slug in extra_columns:
                            row_payload[companion_slug] = companion_value

                if table.slug == "applied" and row_payload:
                    row_payload.setdefault("application-progress", "Applied")
                if table.slug == "portals" and row_payload:
                    row_payload.setdefault("fetch-keywords", default_portal_keywords_text())
                if not any(row_payload.values()):
                    continue

                for slug, value in row_payload.items():
                    column = extra_columns.get(slug)
                    if column:
                        DynamicData.objects.create(
                            table=table,
                            column=column,
                            user=user,
                            row_key=row_key,
                            value=value,
                            metadata=row_metadata.get(slug, {}),
                        )
                imported_rows += 1

            imported_counts[table.slug] = imported_rows
    return imported_counts


def workbook_template_sheets(template_path: str | Path | None = None) -> list[dict[str, Any]]:
    workbook_path = Path(template_path) if template_path else None
    if not workbook_path or not workbook_path.exists():
        return []

    workbook = load_workbook(workbook_path, read_only=True)
    template_sheets: list[dict[str, Any]] = []
    try:
        for sheet in workbook.worksheets:
            try:
                header_row = next(sheet.iter_rows(min_row=1, max_row=1))
            except StopIteration:
                header_row = []
            headers = [normalize_text(cell.value) for cell in header_row if normalize_text(cell.value)]
            template_sheets.append(
                {
                    "slug": normalize_slug(sheet.title),
                    "title": sheet.title,
                    "headers": headers,
                }
            )
    finally:
        workbook.close()

    return template_sheets


def make_excel_sheet_title(title: str, used_titles: set[str]) -> str:
    cleaned = re.sub(r"[\[\]\*:/\\?]", " ", normalize_text(title) or "Sheet").strip()
    base_title = cleaned[:31] or "Sheet"
    candidate = base_title
    counter = 2
    while candidate in used_titles:
        suffix = f" {counter}"
        candidate = f"{base_title[: 31 - len(suffix)]}{suffix}"
        counter += 1
    used_titles.add(candidate)
    return candidate


def export_header_items(table: DynamicTable, template_headers: list[str]) -> list[dict[str, str]]:
    columns = list(table.columns.all())
    columns_by_slug = {column.slug: column for column in columns}
    header_items: list[dict[str, str]] = []
    used_slugs: set[str] = set()

    for header in template_headers:
        header_slug = normalize_slug(header)
        if header_slug in DEPRECATED_COLUMN_SLUGS:
            header_items.append({"label": header, "slug": ""})
            continue
        column = columns_by_slug.get(header_slug)
        if not column:
            continue
        header_items.append({"label": header, "slug": column.slug})
        used_slugs.add(column.slug)

    for column in columns:
        if column.slug in used_slugs:
            continue
        header_items.append({"label": column.name, "slug": column.slug})

    return header_items


def export_cell_display_value(table: DynamicTable, record: GroupedRecord, column: DynamicColumn) -> str:
    raw_value = normalize_text(record.values.get(column.slug))
    if not raw_value:
        return ""
    if table.slug == "applied" and column.slug == "url":
        return normalize_text(record.values.get("role-title")) or raw_value
    return raw_value


def build_export_workbook(
    user,
    template_path: str | Path | None = None,
) -> Workbook:
    template_sheets = workbook_template_sheets(template_path)
    template_by_slug = {item["slug"]: item for item in template_sheets}
    tables = list(
        DynamicTable.objects.filter(owner=user)
        .exclude(slug="dashboard")
        .prefetch_related("columns")
        .order_by("display_order", "name")
    )
    tables_by_slug = {table.slug: table for table in tables}
    ordered_tables: list[DynamicTable] = []

    for template_sheet in template_sheets:
        table = tables_by_slug.pop(template_sheet["slug"], None)
        if table:
            ordered_tables.append(table)

    ordered_tables.extend(
        sorted(
            tables_by_slug.values(),
            key=lambda table: (table.display_order, table.name.lower()),
        )
    )

    workbook = Workbook()
    workbook.remove(workbook.active)
    used_titles: set[str] = set()

    for table in ordered_tables:
        template_sheet = template_by_slug.get(table.slug)
        sheet_title = template_sheet["title"] if template_sheet else table.name
        worksheet = workbook.create_sheet(make_excel_sheet_title(sheet_title, used_titles))
        header_items = export_header_items(table, template_sheet["headers"] if template_sheet else [])

        if not header_items:
            continue

        column_widths: dict[int, int] = {}
        column_map = {column.slug: column for column in table.columns.all()}

        for column_index, header_item in enumerate(header_items, start=1):
            worksheet.cell(row=1, column=column_index, value=header_item["label"])
            column_widths[column_index] = len(header_item["label"])

        records = get_records(table, user)
        for row_index, record in enumerate(records, start=2):
            for column_index, header_item in enumerate(header_items, start=1):
                column_slug = header_item["slug"]
                cell = worksheet.cell(row=row_index, column=column_index)

                if not column_slug:
                    cell.value = row_index - 1
                    column_widths[column_index] = max(column_widths[column_index], len(str(cell.value)))
                    continue

                column = column_map[column_slug]
                raw_value = normalize_text(record.values.get(column_slug))
                if not raw_value:
                    continue

                display_value = export_cell_display_value(table, record, column)
                display_text = display_value or raw_value

                if column.data_type == DynamicColumn.DataType.DATE:
                    try:
                        cell.value = datetime.fromisoformat(raw_value).date()
                        cell.number_format = "DD/MM/YYYY"
                        display_text = cell.value.strftime("%d/%m/%Y")
                    except ValueError:
                        cell.value = raw_value
                        display_text = raw_value
                elif column.data_type == DynamicColumn.DataType.URL:
                    hyperlink = normalize_text(record.metadata.get(column_slug, {}).get("hyperlink")) or raw_value
                    cell.value = display_text
                    if looks_like_url(hyperlink):
                        cell.hyperlink = hyperlink
                        cell.style = "Hyperlink"
                else:
                    cell.value = raw_value
                    display_text = raw_value

                column_widths[column_index] = max(column_widths[column_index], len(display_text))

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = f"A1:{get_column_letter(len(header_items))}1"
        for column_index, width in column_widths.items():
            worksheet.column_dimensions[get_column_letter(column_index)].width = min(max(width + 2, 10), 60)

    return workbook


def infer_progress(record: GroupedRecord) -> str:
    result = (record.values.get("result") or "").strip()
    progress = (record.values.get("application-progress") or "").strip()
    stages = (record.values.get("completed-stages") or "").lower()
    if result:
        normalized = result.lower()
        if normalized == "rejected":
            return "Rejected"
        if normalized in {"offer", "accepted"}:
            return result.title()
    if progress:
        return progress
    if "interview" in stages:
        return "Interview"
    return "Applied"


def dashboard_template_initial_values(
    template_key: str,
    *,
    sort_order: int | None = None,
) -> dict[str, str]:
    template = DASHBOARD_WIDGET_TEMPLATES.get(template_key)
    if not template:
        raise ValueError("Unknown dashboard template.")
    values = {key: normalize_text(value) for key, value in template.items()}
    values["template-key"] = template_key
    if sort_order is not None:
        values["sort-order"] = str(sort_order)
    return values


def ensure_default_dashboard_widgets(user, dashboard_table: DynamicTable) -> None:
    existing_records = get_records(dashboard_table, user)
    existing_by_template = {
        normalize_text(record.values.get("template-key")): record
        for record in existing_records
        if normalize_text(record.values.get("template-key"))
    }
    for position, template_key in enumerate(DEFAULT_DASHBOARD_TEMPLATE_KEYS, start=1):
        existing_record = existing_by_template.get(template_key)
        if existing_record:
            missing_values = {}
            defaults = dashboard_template_initial_values(template_key, sort_order=position * 10)
            for field_name in ("show-on-dashboard", "sort-order", "card-size", "chart-height"):
                if not normalize_text(existing_record.values.get(field_name)):
                    missing_values[field_name] = defaults[field_name]
            if missing_values:
                save_partial_record(dashboard_table, user, existing_record.row_key, missing_values)
            continue
        save_record(
            dashboard_table,
            user,
            dashboard_template_initial_values(template_key, sort_order=position * 10),
        )


def build_dashboard_template_catalog(dashboard_records: list[GroupedRecord]) -> list[dict[str, Any]]:
    records_by_template = {
        normalize_text(record.values.get("template-key")): record
        for record in dashboard_records
        if normalize_text(record.values.get("template-key"))
    }
    catalog: list[dict[str, Any]] = []
    for template_key, definition in DASHBOARD_WIDGET_TEMPLATES.items():
        record = records_by_template.get(template_key)
        catalog.append(
            {
                "key": template_key,
                "title": definition["title"],
                "chart_type": definition["chart-type"],
                "notes": definition["notes"],
                "is_seeded": template_key in DEFAULT_DASHBOARD_TEMPLATE_KEYS,
                "has_record": record is not None,
                "row_key": record.row_key if record else None,
                "show_on_dashboard": normalize_text(record.values.get("show-on-dashboard")) == "Show" if record else False,
            }
        )
    return catalog


def resolve_dashboard_grid_span(card_size: str) -> int:
    normalized = normalize_text(card_size).lower()
    if normalized == "compact":
        return 4
    if normalized == "wide":
        return 8
    if normalized == "full":
        return 12
    return 6


def resolve_dashboard_chart_height(chart_height: str) -> int:
    normalized = normalize_text(chart_height).lower()
    if normalized == "short":
        return 280
    if normalized == "tall":
        return 430
    if normalized == "xl":
        return 520
    return 360


def normalize_dashboard_widget_dimensions(
    chart_type: str,
    *,
    grid_span: int,
    chart_height_px: int,
) -> tuple[int, int]:
    normalized_chart_type = normalize_text(chart_type).lower()
    if normalized_chart_type == "sankey":
        return max(grid_span, 8), max(chart_height_px, 430)
    return grid_span, chart_height_px


def parse_filter_text(filter_text: str) -> dict[str, list[str]]:
    parsed: dict[str, list[str]] = {}
    for item in split_words(filter_text):
        if "=" not in item:
            continue
        key, raw_values = item.split("=", 1)
        values = [value.strip() for value in raw_values.split("|") if value.strip()]
        if key.strip() and values:
            parsed[normalize_slug(key)] = values
    return parsed


def record_matches_filters(record: GroupedRecord, filter_map: dict[str, list[str]]) -> bool:
    for key, expected_values in filter_map.items():
        actual_value = normalize_text(record.values.get(key))
        if not actual_value:
            return False
        actual_normalized = actual_value.lower()
        allowed = {value.lower() for value in expected_values}
        if actual_normalized not in allowed:
            return False
    return True


def aggregate_records(
    records: list[GroupedRecord],
    group_by_column: str,
    aggregation: str,
    metric_column: str,
    limit: int,
) -> tuple[list[str], list[float]]:
    if not group_by_column:
        if aggregation == "sum-numbers":
            total = sum(parse_number(record.values.get(metric_column)) or 0 for record in records)
        elif aggregation == "count-filled-values":
            total = sum(1 for record in records if normalize_text(record.values.get(metric_column)))
        else:
            total = len(records)
        return ["Total"], [float(total)]

    grouped: dict[str, list[GroupedRecord]] = defaultdict(list)
    for record in records:
        key = normalize_text(record.values.get(group_by_column)) or "Unspecified"
        grouped[key].append(record)

    labels: list[str] = []
    values: list[float] = []
    for label, group in grouped.items():
        if aggregation == "sum-numbers":
            total = sum(parse_number(record.values.get(metric_column)) or 0 for record in group)
        elif aggregation == "count-filled-values":
            total = sum(1 for record in group if normalize_text(record.values.get(metric_column)))
        else:
            total = len(group)
        labels.append(label)
        values.append(float(total))

    group_by_definition = records[0].table.columns.filter(slug=group_by_column).first() if records else None
    if group_by_definition and group_by_definition.data_type == DynamicColumn.DataType.DATE:
        def date_group_key(item: tuple[str, float]) -> tuple[int, Any]:
            label = normalize_text(item[0])
            if not label or label == "Unspecified":
                return (1, "")
            try:
                return (0, datetime.fromisoformat(label).date())
            except ValueError:
                return (1, label.lower())

        ordered = sorted(zip(labels, values), key=date_group_key)
    else:
        ordered = sorted(zip(labels, values), key=lambda item: (-item[1], item[0].lower()))
    if limit > 0:
        ordered = ordered[:limit]
    if not ordered:
        return [], []
    sorted_labels, sorted_values = zip(*ordered)
    return list(sorted_labels), list(sorted_values)


def chart_type_slug(value: str) -> str:
    return normalize_slug(value or "kpi")


def build_waterfall_payload(labels: list[str], values: list[float]) -> dict[str, Any]:
    running_total = 0.0
    floating_values: list[list[float]] = []
    colors: list[str] = []
    raw_values: list[float] = []
    for label, value in zip(labels, values):
        delta = -value if label.lower() in {"rejected", "withdrawn"} else value
        raw_values.append(delta)
        start = running_total
        end = running_total + delta
        floating_values.append([min(start, end), max(start, end)])
        colors.append("#f43f5e" if delta < 0 else "#0ea5e9")
        running_total = end
    return {
        "labels": labels,
        "datasets": [
            {
                "type": "bar",
                "label": "Waterfall",
                "data": floating_values,
                "rawValues": raw_values,
                "backgroundColor": colors,
                "borderRadius": 8,
            }
        ],
    }


def build_applied_job_type_chart(applied_records: list[GroupedRecord]) -> dict[str, Any]:
    type_counts = Counter(normalize_text(record.values.get("type")) or "Unspecified" for record in applied_records)
    labels = list(type_counts.keys())
    values = [type_counts[label] for label in labels]
    return {
        "title": "Applied Job Types",
        "description": "A quick view of how your applied roles split across job types.",
        "kind": "chartjs",
        "chart_type": "pie",
        "chart_config": {
            "labels": labels,
            "datasets": [
                {
                    "label": "Applied job types",
                    "data": values,
                    "backgroundColor": label_color_palette(labels),
                    "borderColor": [semantic_hex(label) for label in labels],
                    "borderWidth": 1,
                }
            ],
        },
    }


def build_applied_flow_chart(applied_records: list[GroupedRecord]) -> dict[str, Any]:
    if not applied_records:
        return {
            "title": "Applied Flow",
            "description": "Applied to job type to progress status flow.",
            "kind": "plotly",
            "chart_type": "sankey",
            "chart_config": {"data": [], "layout": {}},
        }

    root_label = "Applied"
    job_types = sorted({normalize_text(record.values.get("type")) or "Unspecified" for record in applied_records})
    stages = sorted({infer_progress(record) for record in applied_records})
    counts_by_type = Counter(normalize_text(record.values.get("type")) or "Unspecified" for record in applied_records)
    stage_counts = Counter(infer_progress(record) for record in applied_records)
    node_keys = [root_label, *job_types, *stages]
    node_labels = [
        f"{root_label} ({len(applied_records)})",
        *[f"{job_type} ({counts_by_type[job_type]})" for job_type in job_types],
        *[f"{stage} ({stage_counts[stage]})" for stage in stages],
    ]
    node_indices = {label: index for index, label in enumerate(node_keys)}
    job_type_colors = [semantic_hex(job_type, alpha=0.72) for job_type in job_types]
    stage_colors = [semantic_hex(stage, alpha=0.82) for stage in stages]

    source: list[int] = []
    target: list[int] = []
    value: list[int] = []
    link_colors: list[str] = []

    for job_type, count in counts_by_type.items():
        source.append(node_indices[root_label])
        target.append(node_indices[job_type])
        value.append(count)
        link_colors.append(semantic_hex(job_type, alpha=0.24))

    type_stage_counts = Counter(
        (
            normalize_text(record.values.get("type")) or "Unspecified",
            infer_progress(record),
        )
        for record in applied_records
    )
    for (job_type, stage), count in type_stage_counts.items():
        source.append(node_indices[job_type])
        target.append(node_indices[stage])
        value.append(count)
        link_colors.append(semantic_hex(stage, alpha=0.24))

    return {
        "title": "Applied Flow",
        "description": "From applied, through job type, into the current application stage.",
        "kind": "plotly",
        "chart_type": "sankey",
        "chart_config": {
            "data": [
                {
                    "type": "sankey",
                    "orientation": "h",
                    "arrangement": "snap",
                    "node": {
                        "pad": 22,
                        "thickness": 20,
                        "line": {"color": "rgba(15, 23, 42, 0.18)", "width": 1},
                        "label": node_labels,
                        "color": ["rgba(14, 165, 233, 0.85)", *job_type_colors, *stage_colors],
                    },
                    "link": {
                        "source": source,
                        "target": target,
                        "value": value,
                        "color": link_colors,
                    },
                }
            ],
            "layout": {
                "paper_bgcolor": "rgba(0,0,0,0)",
                "plot_bgcolor": "rgba(0,0,0,0)",
                "font": {"size": 12, "color": "#0f172a"},
                "margin": {"l": 40, "r": 40, "t": 18, "b": 24},
            },
        },
    }


def build_dashboard_widget(record: GroupedRecord, user) -> DashboardWidget:
    chart_type = chart_type_slug(record.values.get("chart-type") or "kpi")
    source_table_slug = normalize_slug(record.values.get("source-table") or "applied")
    aggregation = chart_type_slug(record.values.get("aggregation") or "count-records")
    group_by_column = normalize_slug(record.values.get("group-by-column"))
    metric_column = normalize_slug(record.values.get("metric-column"))
    filters = normalize_text(record.values.get("filters"))
    limit = int(parse_number(record.values.get("limit")) or 8)
    accent = normalize_text(record.values.get("accent")) or "Sky"
    notes = normalize_text(record.values.get("notes"))
    metric_label = normalize_text(record.values.get("metric"))
    template_key = normalize_text(record.values.get("template-key"))
    show_on_dashboard = normalize_text(record.values.get("show-on-dashboard")).lower() != "hide"
    sort_order = int(parse_number(record.values.get("sort-order")) or 999)
    card_size = normalize_text(record.values.get("card-size")) or "Standard"
    chart_height = normalize_text(record.values.get("chart-height")) or "Medium"
    grid_span = resolve_dashboard_grid_span(card_size)
    chart_height_px = resolve_dashboard_chart_height(chart_height)
    grid_span, chart_height_px = normalize_dashboard_widget_dimensions(
        chart_type,
        grid_span=grid_span,
        chart_height_px=chart_height_px,
    )
    labels: list[str] = []
    values: list[float] = []
    chart_config: dict[str, Any] = {}
    kpi_value = "-"
    table_columns: list[str] = []

    try:
        source_table = get_user_table(user, source_table_slug)
        table_columns = [column.slug for column in source_table.columns.filter(is_visible=True)]
        source_records = get_records(source_table, user)
        filter_map = parse_filter_text(filters)
        if filter_map:
            source_records = [item for item in source_records if record_matches_filters(item, filter_map)]
        labels, values = aggregate_records(
            source_records,
            group_by_column=group_by_column,
            aggregation=aggregation,
            metric_column=metric_column,
            limit=limit,
        )
        if chart_type == "kpi":
            if values:
                integer_value = int(values[0])
                kpi_value = str(integer_value) if values[0] == integer_value else f"{values[0]:.2f}"
            else:
                kpi_value = "0"
        elif chart_type == "sankey":
            kpi_value = str(len(source_records))
            if not metric_label:
                metric_label = "Records in flow"
            if source_table_slug == "applied":
                chart_config = build_applied_flow_chart(source_records)["chart_config"]
            else:
                chart_config = {"data": [], "layout": {}}
        elif chart_type == "waterfall":
            chart_config = build_waterfall_payload(labels, values)
        else:
            chart_config = {
                "labels": labels,
                "datasets": [
                    {
                        "label": record.label,
                        "data": values,
                        "backgroundColor": label_color_palette(labels),
                        "borderColor": [semantic_hex(label) for label in labels],
                        "borderWidth": 1,
                    }
                ],
            }
    except Exception:
        chart_config = {"data": [], "layout": {}} if chart_type == "sankey" else {"labels": [], "datasets": [{"label": record.label, "data": []}]}

    return DashboardWidget(
        row_key=record.row_key,
        title=normalize_text(record.values.get("title")) or record.label,
        template_key=template_key,
        chart_type=chart_type,
        accent=accent,
        notes=notes,
        metric_label=metric_label,
        source_table=source_table_slug,
        group_by_column=group_by_column,
        metric_column=metric_column,
        aggregation=aggregation,
        filters=filters,
        labels=labels,
        values=values,
        chart_config=chart_config,
        kpi_value=kpi_value,
        table_columns=table_columns,
        show_on_dashboard=show_on_dashboard,
        sort_order=sort_order,
        card_size=card_size,
        chart_height=chart_height,
        grid_span=grid_span,
        chart_height_px=chart_height_px,
    )


def build_dashboard(user) -> dict[str, Any]:
    workbook_path = Path("Applications.xlsx") if Path("Applications.xlsx").exists() else None
    tables = ensure_user_workspace(user, workbook_path)
    to_apply_records = get_records(tables["to-apply"], user)
    applied_records = get_records(tables["applied"], user)
    ensure_default_dashboard_widgets(user, tables["dashboard"])
    dashboard_records = get_records(tables["dashboard"], user)
    portal_records = get_records(tables["portals"], user)
    all_dashboard_widgets = [build_dashboard_widget(record, user) for record in dashboard_records]
    dashboard_widgets = sorted(
        [widget for widget in all_dashboard_widgets if widget.show_on_dashboard],
        key=lambda widget: (widget.sort_order, widget.title.lower()),
    )
    hidden_dashboard_widgets = sorted(
        [widget for widget in all_dashboard_widgets if not widget.show_on_dashboard],
        key=lambda widget: (widget.sort_order, widget.title.lower()),
    )

    progress_counts = Counter(infer_progress(record) for record in applied_records)
    max_stage_count = max(progress_counts.values(), default=1)
    upcoming_deadlines = [record for record in to_apply_records if record.values.get("deadline")]

    metrics = [
        {
            "label": "To Apply",
            "value": len(to_apply_records),
            "hint": "Open opportunities waiting on you",
        },
        {
            "label": "Applied",
            "value": len(applied_records),
            "hint": "Applications already submitted",
        },
        {
            "label": "Interviews",
            "value": sum(count for label, count in progress_counts.items() if "interview" in label.lower()),
            "hint": "Applications currently in interview stages",
        },
        {
            "label": "Portals",
            "value": len(portal_records),
            "hint": "Saved portals you can scrape for new openings",
        },
    ]
    stage_breakdown = [
        {
            "label": label,
            "count": count,
            "percentage": round((count / max_stage_count) * 100),
            "color": semantic_hex(label),
        }
        for label, count in progress_counts.most_common()
    ]
    return {
        "metrics": metrics,
        "stage_breakdown": stage_breakdown,
        "upcoming_deadlines": upcoming_deadlines[:6],
        "recent_applied": applied_records[:6],
        "custom_cards": dashboard_records,
        "dashboard_widgets": dashboard_widgets,
        "hidden_dashboard_widgets": hidden_dashboard_widgets,
        "dashboard_templates": build_dashboard_template_catalog(dashboard_records),
        "dashboard_table": tables["dashboard"],
    }


def move_to_applied(
    *,
    user,
    applied_table: DynamicTable,
    to_apply_table: DynamicTable,
    values: dict[str, Any],
    source_row_key: UUID | str | None = None,
) -> UUID:
    source_values: dict[str, str] = {}
    source_metadata: dict[str, dict[str, Any]] = {}
    if source_row_key:
        record = get_record(to_apply_table, user, source_row_key)
        if not record:
            raise ValueError("The selected To Apply entry no longer exists.")
        source_values = dict(record.values)
        source_metadata = dict(record.metadata)

    merged_values = dict(source_values)
    merged_values.update({key: value for key, value in values.items() if value not in {"", None}})
    merged_values.setdefault("date", timezone.localdate().isoformat())
    merged_values.setdefault("application-progress", "Applied")

    row_uuid = save_record(
        applied_table,
        user,
        merged_values,
        metadata_map=source_metadata,
    )
    if source_row_key:
        delete_record(to_apply_table, user, source_row_key)
    return row_uuid


def resolve_portal_keywords(raw_keywords: str | None) -> list[str]:
    keywords = split_words(raw_keywords or "")
    if not keywords:
        keywords = list(DEFAULT_PORTAL_KEYWORDS)
    return keywords


def extract_jobs_from_portal_html(url: str, html: str, keywords: list[str] | None = None) -> list[dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    seen: set[str] = set()
    jobs: list[dict[str, str]] = []
    active_keywords = keywords or list(DEFAULT_PORTAL_KEYWORDS)

    for anchor in soup.find_all("a", href=True):
        title = normalize_text(anchor.get_text(" ", strip=True))
        href = urljoin(url, anchor["href"])
        summary = normalize_text(anchor.parent.get_text(" ", strip=True))
        haystack = f"{title} {summary} {href}"
        if len(title) < 8:
            continue
        if matches_any_keyword(haystack, PORTAL_NEGATIVE_KEYWORDS):
            continue
        if not matches_any_keyword(haystack, JOB_KEYWORDS):
            continue
        if active_keywords and not matches_any_keyword(haystack, active_keywords):
            continue
        if href in seen:
            continue
        try:
            validate_safe_external_url(href)
        except ValueError:
            continue
        seen.add(href)
        jobs.append(
            {
                "title": title[:255],
                "url": href[:1000],
                "summary": summary[:500],
            }
        )
        if len(jobs) >= 40:
            break
    return jobs


def scrape_jobs_from_portal(url: str, keywords: list[str] | None = None) -> list[dict[str, str]]:
    mode = portal_fetch_mode()
    browser_required_error: Exception | None = None
    static_jobs: list[dict[str, str]] = []
    static_html = ""

    if mode != "browser":
        try:
            response = safe_get_url(url, timeout=20)
            static_html = response.text
        except Exception as exc:
            if mode == "requests" or not should_retry_portal_with_browser(exc):
                raise
            browser_required_error = exc
        else:
            static_jobs = extract_jobs_from_portal_html(url, static_html, keywords)
            if mode == "requests":
                return static_jobs
            if static_jobs and not portal_html_looks_blocked(static_html):
                return static_jobs
            if portal_html_looks_blocked(static_html):
                browser_required_error = RuntimeError("Portal page requires browser rendering.")
            else:
                try:
                    browser_html = fetch_portal_html_in_browser(url)
                except Exception:
                    return static_jobs
                browser_jobs = extract_jobs_from_portal_html(url, browser_html, keywords)
                return browser_jobs or static_jobs

    browser_html = fetch_portal_html_in_browser(url)
    browser_jobs = extract_jobs_from_portal_html(url, browser_html, keywords)
    if browser_jobs or browser_required_error is None:
        return browser_jobs

    raise RuntimeError(
        "Portal fetch was blocked over regular HTTP "
        f"({browser_required_error}). Selenium browser fallback also found no matching jobs."
    )


def refresh_portal_jobs(
    user,
    portal_table: DynamicTable,
    portal_row_key: UUID | str,
    raw_keywords: str | None = None,
) -> list[PortalJobSuggestion]:
    portal_record = get_record(portal_table, user, portal_row_key)
    if not portal_record:
        raise ValueError("Portal entry not found.")

    portal_url = portal_record.values.get("url")
    if not portal_url:
        raise ValueError("This portal entry does not have a URL.")

    portal_keywords_text = raw_keywords or portal_record.values.get("fetch-keywords") or default_portal_keywords_text()
    save_partial_record(
        portal_table,
        user,
        portal_record.row_key,
        {"fetch-keywords": portal_keywords_text},
    )

    portal_name = portal_record.values.get("name") or portal_record.label
    jobs = scrape_jobs_from_portal(portal_url, resolve_portal_keywords(portal_keywords_text))

    with transaction.atomic():
        PortalJobSuggestion.objects.filter(
            user=user,
            portal_table=portal_table,
            portal_row_key=portal_record.row_key,
        ).delete()
        created_items = [
            PortalJobSuggestion(
                user=user,
                portal_table=portal_table,
                portal_row_key=portal_record.row_key,
                portal_name=portal_name[:255],
                portal_url=portal_url,
                title=job["title"],
                url=job["url"],
                summary=job["summary"],
            )
            for job in jobs
        ]
        PortalJobSuggestion.objects.bulk_create(created_items)

    return list(
        PortalJobSuggestion.objects.filter(
            user=user,
            portal_table=portal_table,
            portal_row_key=portal_record.row_key,
        ).order_by("title")
    )


def refresh_all_portal_jobs_for_user(user) -> dict[str, Any]:
    portal_table = get_user_table(user, "portals")
    portal_records = get_records(portal_table, user)
    refreshed = 0
    total_jobs = 0
    errors: list[str] = []

    for portal_record in portal_records:
        try:
            jobs = refresh_portal_jobs(
                user,
                portal_table,
                portal_record.row_key,
                raw_keywords=portal_record.values.get("fetch-keywords"),
            )
            refreshed += 1
            total_jobs += len(jobs)
        except Exception as exc:
            errors.append(f"{portal_record.label}: {exc}")

    return {
        "refreshed": refreshed,
        "total_jobs": total_jobs,
        "errors": errors,
        "portal_records": portal_records,
    }


def add_portal_job_to_to_apply(user, suggestion: PortalJobSuggestion) -> UUID:
    to_apply_table = get_user_table(user, "to-apply")
    if DynamicData.objects.filter(
        table=to_apply_table,
        user=user,
        column__slug="url",
        value=suggestion.url,
    ).exists():
        suggestion.added_to_apply = True
        suggestion.save(update_fields=["added_to_apply"])
        existing = DynamicData.objects.get(
            table=to_apply_table,
            user=user,
            column__slug="url",
            value=suggestion.url,
        )
        return existing.row_key

    row_key = save_record(
        to_apply_table,
        user,
        {
            "role-title": suggestion.title,
            "url": suggestion.url,
            "type": infer_application_type(suggestion.title),
        },
    )
    suggestion.added_to_apply = True
    suggestion.save(update_fields=["added_to_apply"])
    return row_key


def ensure_named_user(
    *,
    username: str,
    email: str,
    password: str,
    is_staff: bool = False,
    is_superuser: bool = False,
):
    User = get_user_model()
    now = timezone.now()
    user, _ = User.objects.get_or_create(
        username=username,
        defaults={
            "email": email,
            "is_staff": is_staff,
            "is_superuser": is_superuser,
            "is_active": True,
            "email_verified_at": now,
            "chat_history_retention_days": 1,
        },
    )
    changed = False
    if email and user.email != email:
        user.email = email
        changed = True
    if email and user.notification_email != email:
        user.notification_email = email
        changed = True
    if user.is_staff != is_staff:
        user.is_staff = is_staff
        changed = True
    if user.is_superuser != is_superuser:
        user.is_superuser = is_superuser
        changed = True
    if not user.is_active:
        user.is_active = True
        changed = True
    if not user.email_verified_at:
        user.email_verified_at = now
        changed = True
    current_retention = int(getattr(user, "chat_history_retention_days", 1) or 1)
    if current_retention < 1 or current_retention == 2:
        user.chat_history_retention_days = 1
        changed = True
    user.set_password(password)
    changed = True
    if changed:
        user.save()
    return user


def bootstrap_local_users(excel_path: str | Path | None = None, replace: bool = True) -> dict[str, Any]:
    from .chat_handler import ChatHandler

    admin = ensure_named_user(
        username=os.getenv("ADMIN_USERNAME", "admin"),
        email=os.getenv("ADMIN_EMAIL", "admin@example.com"),
        password=os.getenv("ADMIN_PASSWORD", "Admin12345!"),
        is_staff=True,
        is_superuser=True,
    )
    test_user = ensure_named_user(
        username=os.getenv("TEST_USERNAME", "testuser"),
        email=os.getenv("TEST_EMAIL", "testuser@example.com"),
        password=os.getenv("TEST_PASSWORD", "TestUser12345!"),
        is_staff=False,
        is_superuser=False,
    )

    ensure_user_workspace(admin, excel_path)
    ensure_user_workspace(test_user, excel_path)
    ChatHandler.ensure_default_models()
    imported_counts = {}
    if excel_path and Path(excel_path).exists():
        imported_counts = import_workbook_for_user(excel_path, test_user, replace=replace)
    return {
        "admin": admin,
        "test_user": test_user,
        "imported_counts": imported_counts,
    }
