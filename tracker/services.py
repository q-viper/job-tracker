from __future__ import annotations

import ipaddress
import json
import os
import re
import shutil
import socket
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime, time, timedelta
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

from .models import (
    ApplicationHistoryEvent,
    ApplicationTag,
    ApplicationTagAssignment,
    DashboardBoard,
    DashboardBoardItem,
    DashboardPanelPreference,
    DynamicColumn,
    DynamicData,
    DynamicTable,
    PortalJobSuggestion,
)

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
        "number-of-positions": {
            "data_type": DynamicColumn.DataType.NUMBER,
            "placeholder": "1",
            "help_text": "Optional. Record how many openings the posting mentions.",
        },
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
        "deadline": {
            "data_type": DynamicColumn.DataType.DATE,
            "help_text": "Optional. Track the posting deadline so you can compare when you applied.",
        },
        "result": {
            "data_type": DynamicColumn.DataType.STATUS,
            "choices_text": "Pending, Rejected, Withdrawn, Offer, Accepted",
        },
        "result-on": {"data_type": DynamicColumn.DataType.DATE},
        "number-of-positions": {
            "data_type": DynamicColumn.DataType.NUMBER,
            "placeholder": "1",
            "help_text": "Optional. Record how many openings the posting mentions.",
        },
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
        {
            "name": "Number of Positions",
            "slug": "number-of-positions",
            "data_type": DynamicColumn.DataType.NUMBER,
            "insert_after": "type",
            "placeholder": "1",
            "help_text": "Optional. Record how many openings the posting mentions.",
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
            "name": "Deadline",
            "slug": "deadline",
            "data_type": DynamicColumn.DataType.DATE,
            "insert_after": "type",
            "help_text": "Optional. Track the posting deadline so you can compare when you applied.",
        },
        {
            "name": "Number of Positions",
            "slug": "number-of-positions",
            "data_type": DynamicColumn.DataType.NUMBER,
            "insert_after": "type",
            "placeholder": "1",
            "help_text": "Optional. Record how many openings the posting mentions.",
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
        "notes": "Stage-to-stage application flow built from tracked history.",
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
    "applied-active-pipeline": {
        "title": "Active Applied Pipeline",
        "chart-type": "Bar",
        "source-table": "applied",
        "group-by-column": "application-progress",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "10",
        "metric": "Active applications",
        "notes": "Only includes applied jobs that are not rejected or withdrawn.",
        "accent": "Sky",
        "show-on-dashboard": "Show",
        "sort-order": "35",
        "card-size": "Wide",
        "chart-height": "Medium",
    },
    "applied-stage-timing": {
        "title": "Average Days To Each Stage",
        "chart-type": "Horizontal Bar",
        "source-table": "applied",
        "group-by-column": "",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "12",
        "metric": "Average days",
        "notes": "Average days from the first tracked moment for each job, using To Apply history when available.",
        "accent": "Emerald",
        "show-on-dashboard": "Show",
        "sort-order": "40",
        "card-size": "Wide",
        "chart-height": "Tall",
    },
    "applied-deadline-lead": {
        "title": "Applied Before Deadline",
        "chart-type": "Bar",
        "source-table": "applied",
        "group-by-column": "",
        "metric-column": "",
        "aggregation": "Count Records",
        "filters": "",
        "limit": "12",
        "metric": "Applications",
        "notes": "How many days before the deadline you usually submit. Records without both applied date and deadline are skipped.",
        "accent": "Rose",
        "show-on-dashboard": "Show",
        "sort-order": "50",
        "card-size": "Wide",
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

DEFAULT_DASHBOARD_TEMPLATE_KEYS = (
    "applied-job-types",
    "applied-flow",
    "applied-progress-pie",
    "applied-active-pipeline",
    "applied-stage-timing",
    "applied-deadline-lead",
)

DASHBOARD_PANEL_LIBRARY = (
    {
        "key": "stage-breakdown",
        "label": "Application Stages",
        "description": "Progress bars that summarize your current stage mix.",
        "default_zone": "left",
        "default_order": 10,
    },
    {
        "key": "quick-actions",
        "label": "Quick Actions",
        "description": "Handy shortcuts for exporting, editing widgets, and managing layout.",
        "default_zone": "left",
        "default_order": 20,
    },
    {
        "key": "plot-templates",
        "label": "Plot Templates",
        "description": "Starter plot definitions you can add or edit.",
        "default_zone": "left",
        "default_order": 30,
    },
    {
        "key": "metrics-strip",
        "label": "Metrics Strip",
        "description": "Top-level counters for to-apply, applied, interviews, and portals.",
        "default_zone": "main",
        "default_order": 10,
    },
    {
        "key": "widget-canvas",
        "label": "Widget Canvas",
        "description": "Your draggable chart area with resizable plot cards.",
        "default_zone": "main",
        "default_order": 20,
    },
    {
        "key": "hidden-widgets",
        "label": "Hidden Widgets",
        "description": "Saved plots that are not currently shown on the dashboard.",
        "default_zone": "right",
        "default_order": 10,
    },
    {
        "key": "upcoming-deadlines",
        "label": "Upcoming Deadlines",
        "description": "To Apply jobs with deadlines coming up.",
        "default_zone": "right",
        "default_order": 20,
    },
    {
        "key": "recent-submissions",
        "label": "Recent Submissions",
        "description": "Your most recent applied records and their current stage.",
        "default_zone": "right",
        "default_order": 30,
    },
)

DEFAULT_DASHBOARD_BOARD_PRESETS = (
    {
        "name": "Home Lens",
        "description": "A clean view of your momentum, progress mix, and stage flow.",
        "show_on_homepage": True,
        "autoplay_enabled": True,
        "autoplay_interval_seconds": 8,
        "sort_order": 10,
        "widgets": (
            {"template_key": "applied-flow", "card_size": "Full", "chart_height": "Tall"},
            {"template_key": "applied-progress-pie", "card_size": "Standard", "chart_height": "Medium"},
            {"template_key": "applied-stage-timing", "card_size": "Wide", "chart_height": "Tall"},
            {"template_key": "applied-deadline-lead", "card_size": "Wide", "chart_height": "Medium"},
        ),
    },
    {
        "name": "Pipeline Lab",
        "description": "Focus on how applications move through the funnel and where outcomes stack up.",
        "show_on_homepage": True,
        "autoplay_enabled": False,
        "autoplay_interval_seconds": 8,
        "sort_order": 20,
        "widgets": (
            {"template_key": "applied-flow", "card_size": "Full", "chart_height": "Tall"},
            {"template_key": "applied-stage-timing", "card_size": "Wide", "chart_height": "Tall"},
            {"template_key": "applied-job-types", "card_size": "Wide", "chart_height": "Medium"},
            {"template_key": "applied-progress-pie", "card_size": "Standard", "chart_height": "Medium"},
        ),
    },
    {
        "name": "Deadline Radar",
        "description": "Keep an eye on submission lead time and upcoming date pressure.",
        "show_on_homepage": False,
        "autoplay_enabled": True,
        "autoplay_interval_seconds": 10,
        "sort_order": 30,
        "widgets": (
            {"template_key": "applied-deadline-lead", "card_size": "Wide", "chart_height": "Medium"},
            {"template_key": "deadline-count", "card_size": "Wide", "chart_height": "Medium"},
            {"template_key": "applied-progress-pie", "card_size": "Standard", "chart_height": "Medium"},
        ),
    },
)

JOB_TABLE_SLUGS = {"to-apply", "applied"}

DEFAULT_APPLICATION_TAGS = ("Dummy", "PhD", "Job", "Priority")

APPLICATION_STAGE_ORDER = [
    "To Apply",
    "Applied",
    "Screening",
    "Interview",
    "Technical Interview",
    "Assignment",
    "Offer",
    "Accepted",
    "Rejected",
    "Withdrawn",
]

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
    layout_item_id: int | None = None


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


KEYWORD_EQUIVALENTS = {
    "ai": ("artificial intelligence",),
    "artificial intelligence": ("ai",),
    "ml": ("machine learning",),
    "machine learning": ("ml",),
    "cv": ("computer vision",),
    "computer vision": ("cv",),
    "nlp": ("natural language processing",),
    "natural language processing": ("nlp",),
}


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
    variants.extend(KEYWORD_EQUIVALENTS.get(normalized_keyword, ()))
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
    for variant in normalized_keyword_variants(keyword):
        if f" {variant} " in padded_text:
            return True
        variant_tokens = variant.split()
        if len(variant_tokens) > 1 and all(f" {token} " in padded_text for token in variant_tokens):
            return True
    return False


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
    import cloudscraper
    scraper = cloudscraper.create_scraper(
        browser={
            'browser': 'chrome',
            'platform': 'windows',
            'mobile': False,
        }
    )
    response = scraper.get(
        url,
        timeout=timeout,
        headers={
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
    ensure_default_application_tags(user)
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


def ensure_default_application_tags(user) -> None:
    for tag_name in DEFAULT_APPLICATION_TAGS:
        ApplicationTag.objects.get_or_create(
            user=user,
            slug=normalize_slug(tag_name),
            defaults={"name": tag_name},
        )


def parse_tag_names(raw_value: str | None) -> list[str]:
    seen: set[str] = set()
    tag_names: list[str] = []
    for item in re.split(r"[,\n;]+", raw_value or ""):
        cleaned = normalize_text(item)
        slug = normalize_slug(cleaned) if cleaned else ""
        if not cleaned or slug in seen:
            continue
        seen.add(slug)
        tag_names.append(cleaned[:80])
    return tag_names


def get_application_tags_map(user, table_slug: str, row_keys: list[UUID]) -> dict[UUID, list[ApplicationTag]]:
    assignments = list(
        ApplicationTagAssignment.objects.filter(
            user=user,
            table_slug=table_slug,
            record_row_key__in=row_keys,
        )
        .select_related("tag")
        .order_by("tag__name", "id")
    )
    tags_by_row: dict[UUID, list[ApplicationTag]] = defaultdict(list)
    for assignment in assignments:
        tags_by_row[assignment.record_row_key].append(assignment.tag)
    return tags_by_row


def get_application_tags(user, table_slug: str, row_key: UUID | str) -> list[ApplicationTag]:
    return get_application_tags_map(user, table_slug, [UUID(str(row_key))]).get(UUID(str(row_key)), [])


def sync_application_tags(user, table_slug: str, row_key: UUID | str, tag_names: list[str]) -> list[ApplicationTag]:
    ensure_default_application_tags(user)
    row_uuid = UUID(str(row_key))
    normalized_names = parse_tag_names(", ".join(tag_names))
    desired_slugs = {normalize_slug(name): name for name in normalized_names}
    existing_assignments = list(
        ApplicationTagAssignment.objects.filter(
            user=user,
            table_slug=table_slug,
            record_row_key=row_uuid,
        ).select_related("tag")
    )
    existing_by_slug = {assignment.tag.slug: assignment for assignment in existing_assignments}

    with transaction.atomic():
        for slug, assignment in existing_by_slug.items():
            if slug not in desired_slugs:
                assignment.delete()

        for slug, name in desired_slugs.items():
            tag, _ = ApplicationTag.objects.get_or_create(
                user=user,
                slug=slug,
                defaults={"name": name},
            )
            if tag.name != name:
                tag.name = name
                tag.save(update_fields=["name", "updated_at"])
            ApplicationTagAssignment.objects.get_or_create(
                user=user,
                tag=tag,
                table_slug=table_slug,
                record_row_key=row_uuid,
            )

    return get_application_tags(user, table_slug, row_uuid)


def add_application_tags(user, table_slug: str, row_key: UUID | str, tag_names: list[str]) -> list[ApplicationTag]:
    existing = [tag.name for tag in get_application_tags(user, table_slug, row_key)]
    merged = existing + list(tag_names)
    return sync_application_tags(user, table_slug, row_key, merged)


def remove_application_tags(user, table_slug: str, row_key: UUID | str, tag_names: list[str]) -> list[ApplicationTag]:
    removal_slugs = {normalize_slug(name) for name in parse_tag_names(", ".join(tag_names))}
    remaining = [
        tag.name
        for tag in get_application_tags(user, table_slug, row_key)
        if tag.slug not in removal_slugs
    ]
    return sync_application_tags(user, table_slug, row_key, remaining)


def move_application_related_data(
    user,
    source_table_slug: str,
    target_table_slug: str,
    row_key: UUID | str,
) -> None:
    row_uuid = UUID(str(row_key))
    ApplicationHistoryEvent.objects.filter(
        user=user,
        table_slug=source_table_slug,
        record_row_key=row_uuid,
    ).update(table_slug=target_table_slug)
    ApplicationTagAssignment.objects.filter(
        user=user,
        table_slug=source_table_slug,
        record_row_key=row_uuid,
    ).update(table_slug=target_table_slug)


def delete_application_related_data(user, table_slug: str, row_key: UUID | str) -> None:
    row_uuid = UUID(str(row_key))
    ApplicationHistoryEvent.objects.filter(
        user=user,
        table_slug=table_slug,
        record_row_key=row_uuid,
    ).delete()
    ApplicationTagAssignment.objects.filter(
        user=user,
        table_slug=table_slug,
        record_row_key=row_uuid,
    ).delete()


def application_stage_sort_key(stage: str) -> tuple[int, str]:
    normalized = normalize_text(stage)
    if normalized in APPLICATION_STAGE_ORDER:
        return (APPLICATION_STAGE_ORDER.index(normalized), normalized.lower())
    return (len(APPLICATION_STAGE_ORDER), normalized.lower())


def history_datetime_from_date_value(raw_value: Any) -> datetime | None:
    normalized = normalize_date_value(raw_value)
    if not normalized:
        return None
    try:
        parsed_date = datetime.fromisoformat(normalized).date()
    except ValueError:
        return None
    naive = datetime.combine(parsed_date, time(hour=12, minute=0))
    return timezone.make_aware(naive, timezone.get_current_timezone())


def default_sort_field_for_table(table: DynamicTable) -> str:
    columns = list(table.columns.filter(is_visible=True))
    by_slug = {column.slug: column for column in columns}
    if table.slug == "to-apply":
        deadline_column = by_slug.get("deadline")
        if deadline_column and deadline_column.data_type == DynamicColumn.DataType.DATE:
            return "deadline"
    date_column = by_slug.get("date")
    if date_column and date_column.data_type == DynamicColumn.DataType.DATE:
        return "date"
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
    skip_history: bool = False,
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
    if not skip_history:
        sync_job_history_for_saved_record(table, user, row_uuid, existing_record, prepared_values)
    return row_uuid


def save_partial_record(table: DynamicTable, user, row_key: UUID | str, values: dict[str, Any]) -> UUID:
    record = get_record(table, user, row_key)
    if not record:
        raise ValueError("Record not found.")
    merged = dict(record.values)
    merged.update(values)
    return save_record(table, user, merged, row_key=row_key, metadata_map=record.metadata)


def delete_record(table: DynamicTable, user, row_key: UUID | str, *, cleanup_related: bool = True) -> None:
    if table.owner_id != user.id:
        raise PermissionError("You can only delete records from your own tables.")
    DynamicData.objects.filter(table=table, user=user, row_key=row_key).delete()
    if table.slug == "dashboard":
        DashboardBoardItem.objects.filter(board__user=user, widget_row_key=UUID(str(row_key))).delete()
    if cleanup_related and table.slug in JOB_TABLE_SLUGS:
        delete_application_related_data(user, table.slug, row_key)


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
            ApplicationHistoryEvent.objects.filter(user=user).delete()
            ApplicationTagAssignment.objects.filter(user=user).delete()
            DashboardBoardItem.objects.filter(board__user=user).delete()
            DashboardBoard.objects.filter(user=user).delete()

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

                save_record(
                    table,
                    user,
                    row_payload,
                    row_key=row_key,
                    metadata_map=row_metadata,
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


def infer_progress_from_values(values: dict[str, Any]) -> str:
    result = normalize_text(values.get("result"))
    progress = normalize_text(values.get("application-progress"))
    stages = normalize_text(values.get("completed-stages")).lower()
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


def infer_progress(record: GroupedRecord) -> str:
    return infer_progress_from_values(record.values)


def record_application_history_event(
    user,
    table_slug: str,
    row_key: UUID | str,
    event_type: str,
    *,
    from_stage: str = "",
    to_stage: str = "",
    occurred_at: datetime | None = None,
    notes: str = "",
) -> ApplicationHistoryEvent:
    return ApplicationHistoryEvent.objects.create(
        user=user,
        table_slug=table_slug,
        record_row_key=UUID(str(row_key)),
        event_type=event_type,
        from_stage=normalize_text(from_stage),
        to_stage=normalize_text(to_stage),
        occurred_at=occurred_at or timezone.now(),
        notes=normalize_text(notes),
    )


def sync_job_history_for_saved_record(
    table: DynamicTable,
    user,
    row_key: UUID | str,
    existing_record: GroupedRecord | None,
    prepared_values: dict[str, Any],
) -> None:
    if table.slug not in JOB_TABLE_SLUGS:
        return

    if not existing_record:
        if table.slug == "to-apply":
            record_application_history_event(
                user,
                table.slug,
                row_key,
                ApplicationHistoryEvent.EventType.ADDED_TO_APPLY,
                to_stage="To Apply",
                notes="Added to To Apply.",
            )
            return

        applied_at = history_datetime_from_date_value(prepared_values.get("date")) or timezone.now()
        prior_history_time = (
            ApplicationHistoryEvent.objects.filter(
                user=user,
                record_row_key=UUID(str(row_key)),
            )
            .order_by("-occurred_at", "-id")
            .values_list("occurred_at", flat=True)
            .first()
        )
        if prior_history_time and applied_at <= prior_history_time:
            applied_at = prior_history_time + timedelta(minutes=1)
        record_application_history_event(
            user,
            table.slug,
            row_key,
            ApplicationHistoryEvent.EventType.APPLIED,
            to_stage="Applied",
            occurred_at=applied_at,
            notes="Application submitted.",
        )
        current_stage = infer_progress_from_values(prepared_values)
        if current_stage != "Applied":
            record_application_history_event(
                user,
                table.slug,
                row_key,
                ApplicationHistoryEvent.EventType.STAGE_CHANGE,
                from_stage="Applied",
                to_stage=current_stage,
                notes=f"Progress updated to {current_stage}.",
            )
        return

    if table.slug != "applied":
        return

    previous_stage = infer_progress(existing_record)
    current_stage = infer_progress_from_values(prepared_values)
    if current_stage != previous_stage:
        record_application_history_event(
            user,
            table.slug,
            row_key,
            ApplicationHistoryEvent.EventType.STAGE_CHANGE,
            from_stage=previous_stage,
            to_stage=current_stage,
            notes=f"Progress changed from {previous_stage} to {current_stage}.",
        )


def get_application_history_events(user, table_slug: str, row_key: UUID | str) -> list[ApplicationHistoryEvent]:
    return list(
        ApplicationHistoryEvent.objects.filter(
            user=user,
            table_slug=table_slug,
            record_row_key=UUID(str(row_key)),
        ).order_by("occurred_at", "id")
    )


def replace_applied_history_events(user, row_key: UUID | str, stage_dates: dict[str, Any]) -> None:
    row_uuid = UUID(str(row_key))
    timeline_items: list[tuple[str, datetime]] = []
    for stage_name, raw_value in (stage_dates or {}).items():
        stage = normalize_text(stage_name)
        if not stage or stage == "To Apply":
            continue
        occurred_at = history_datetime_from_date_value(raw_value)
        if not occurred_at:
            continue
        timeline_items.append((stage, occurred_at))

    if not timeline_items:
        return

    by_stage: dict[str, datetime] = {}
    for stage, occurred_at in sorted(timeline_items, key=lambda item: (item[1], application_stage_sort_key(item[0]))):
        by_stage.setdefault(stage, occurred_at)

    applied_at = by_stage.get("Applied")
    if not applied_at:
        earliest = min(by_stage.values())
        applied_at = earliest
        by_stage["Applied"] = earliest

    ordered_items = sorted(by_stage.items(), key=lambda item: (item[1], application_stage_sort_key(item[0])))
    if ordered_items[0][0] != "Applied":
        ordered_items = [("Applied", applied_at)] + [item for item in ordered_items if item[0] != "Applied"]

    ApplicationHistoryEvent.objects.filter(
        user=user,
        table_slug="applied",
        record_row_key=row_uuid,
    ).delete()

    previous_stage = "Applied"
    previous_time = applied_at
    record_application_history_event(
        user,
        "applied",
        row_uuid,
        ApplicationHistoryEvent.EventType.APPLIED,
        to_stage="Applied",
        occurred_at=applied_at,
        notes="Application submitted.",
    )

    for stage, occurred_at in ordered_items:
        if stage == "Applied":
            continue
        if occurred_at <= previous_time:
            occurred_at = previous_time + timedelta(minutes=1)
        record_application_history_event(
            user,
            "applied",
            row_uuid,
            ApplicationHistoryEvent.EventType.STAGE_CHANGE,
            from_stage=previous_stage,
            to_stage=stage,
            occurred_at=occurred_at,
            notes=f"Progress updated to {stage}.",
        )
        previous_stage = stage
        previous_time = occurred_at


def get_application_history_events_map(
    user,
    table_slug: str,
    row_keys: list[UUID],
) -> dict[UUID, list[ApplicationHistoryEvent]]:
    if not row_keys:
        return {}
    history_entries = list(
        ApplicationHistoryEvent.objects.filter(
            user=user,
            table_slug=table_slug,
            record_row_key__in=row_keys,
        ).order_by("record_row_key", "occurred_at", "id")
    )
    history_by_row: dict[UUID, list[ApplicationHistoryEvent]] = defaultdict(list)
    for entry in history_entries:
        history_by_row[entry.record_row_key].append(entry)
    return history_by_row


def earliest_tracked_job_timestamp(
    record: GroupedRecord,
    events: list[ApplicationHistoryEvent],
) -> datetime | None:
    if events:
        return events[0].occurred_at
    return (
        history_datetime_from_date_value(record.values.get("date"))
        or record.created_at
        or record.updated_at
    )


def first_stage_timestamps(
    record: GroupedRecord,
    events: list[ApplicationHistoryEvent],
) -> dict[str, datetime]:
    first_seen: dict[str, datetime] = {}
    for event in events:
        stage = normalize_text(event.to_stage)
        if stage and stage not in first_seen:
            first_seen[stage] = event.occurred_at

    if first_seen:
        return first_seen

    fallback_stage = infer_progress(record)
    fallback_time = (
        history_datetime_from_date_value(record.values.get("date"))
        or record.created_at
        or record.updated_at
    )
    if fallback_stage and fallback_time:
        first_seen[fallback_stage] = fallback_time
    return first_seen


def build_application_history_timeline(
    user,
    table_slug: str,
    record: GroupedRecord,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    events = get_application_history_events(user, table_slug, record.row_key)
    if not events:
        fallback_stage = "To Apply" if table_slug == "to-apply" else infer_progress(record)
        fallback_time = record.created_at or timezone.now()
        events = [
            ApplicationHistoryEvent(
                user=user,
                table_slug=table_slug,
                record_row_key=record.row_key,
                event_type=(
                    ApplicationHistoryEvent.EventType.ADDED_TO_APPLY
                    if table_slug == "to-apply"
                    else ApplicationHistoryEvent.EventType.APPLIED
                ),
                to_stage=fallback_stage,
                occurred_at=fallback_time,
                notes="Initial record snapshot.",
            )
        ]

    timeline_items: list[dict[str, Any]] = []
    plot_points: list[dict[str, str]] = []
    for event in events:
        if event.event_type == ApplicationHistoryEvent.EventType.ADDED_TO_APPLY:
            title = "Added to To Apply"
        elif event.event_type == ApplicationHistoryEvent.EventType.APPLIED:
            title = "Applied"
        elif event.from_stage and event.to_stage:
            title = f"{event.from_stage} -> {event.to_stage}"
        else:
            title = event.to_stage or "Stage updated"

        description = event.notes or title
        timeline_items.append(
            {
                "title": title,
                "description": description,
                "stage": event.to_stage or event.from_stage or "",
                "timestamp": event.occurred_at,
                "timestamp_display": timezone.localtime(event.occurred_at).strftime("%d/%m/%Y %H:%M"),
                "event_type": event.event_type,
            }
        )
        plot_points.append(
            {
                "x": event.occurred_at.isoformat(),
                "y": event.to_stage or event.from_stage or title,
                "text": description,
            }
        )
    stage_labels = [
        label
        for label in sorted(
            {point["y"] for point in plot_points},
            key=application_stage_sort_key,
        )
    ]

    plot_config = {
        "data": [
            {
                "type": "scatter",
                "mode": "lines+markers+text",
                "x": [point["x"] for point in plot_points],
                "y": [point["y"] for point in plot_points],
                "text": [point["text"] for point in plot_points],
                "textposition": "top center",
                "hovertemplate": "%{y}<br>%{x}<br>%{text}<extra></extra>",
                "line": {"color": "#0ea5e9", "width": 3, "shape": "hv"},
                "marker": {
                    "size": 12,
                    "color": [semantic_hex(point["y"]) for point in plot_points],
                    "line": {"color": "rgba(15, 23, 42, 0.18)", "width": 1},
                },
            }
        ],
        "layout": {
            "paper_bgcolor": "rgba(0,0,0,0)",
            "plot_bgcolor": "rgba(0,0,0,0)",
            "font": {"size": 12, "color": "#0f172a"},
            "margin": {"l": 40, "r": 20, "t": 18, "b": 36},
            "xaxis": {
                "title": "Timeline",
                "type": "date",
                "showgrid": True,
                "gridcolor": "rgba(148, 163, 184, 0.18)",
            },
            "yaxis": {
                "title": "Stage",
                "categoryorder": "array",
                "categoryarray": stage_labels,
            },
        },
    }
    return timeline_items, plot_config


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


def dashboard_panel_catalog() -> dict[str, dict[str, Any]]:
    return {item["key"]: dict(item) for item in DASHBOARD_PANEL_LIBRARY}


def ensure_dashboard_panel_preferences(user) -> None:
    existing = {
        preference.panel_key: preference
        for preference in DashboardPanelPreference.objects.filter(user=user)
    }
    for panel in DASHBOARD_PANEL_LIBRARY:
        preference = existing.get(panel["key"])
        if preference is None:
            DashboardPanelPreference.objects.create(
                user=user,
                panel_key=panel["key"],
                zone=panel["default_zone"],
                sort_order=panel["default_order"],
            )
            continue
        changed_fields: list[str] = []
        if not normalize_text(preference.zone):
            preference.zone = panel["default_zone"]
            changed_fields.append("zone")
        if preference.sort_order <= 0:
            preference.sort_order = panel["default_order"]
            changed_fields.append("sort_order")
        if changed_fields:
            preference.save(update_fields=changed_fields + ["updated_at"])


def build_dashboard_panel_layout(user) -> dict[str, list[dict[str, Any]]]:
    ensure_dashboard_panel_preferences(user)
    catalog = dashboard_panel_catalog()
    layout = {"left": [], "main": [], "right": [], "hidden": []}
    preferences = DashboardPanelPreference.objects.filter(user=user).order_by("zone", "sort_order", "panel_key")
    for preference in preferences:
        panel = dict(catalog.get(preference.panel_key, {"key": preference.panel_key, "label": preference.panel_key.title()}))
        panel.update(
            {
                "zone": preference.zone,
                "sort_order": preference.sort_order,
            }
        )
        layout.setdefault(preference.zone, []).append(panel)
    return layout


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


def clone_dashboard_widget_with_layout(
    widget: DashboardWidget,
    *,
    card_size: str,
    chart_height: str,
    sort_order: int,
    layout_item_id: int | None = None,
) -> DashboardWidget:
    grid_span = resolve_dashboard_grid_span(card_size)
    chart_height_px = resolve_dashboard_chart_height(chart_height)
    grid_span, chart_height_px = normalize_dashboard_widget_dimensions(
        widget.chart_type,
        grid_span=grid_span,
        chart_height_px=chart_height_px,
    )
    return replace(
        widget,
        card_size=card_size,
        chart_height=chart_height,
        sort_order=sort_order,
        grid_span=grid_span,
        chart_height_px=chart_height_px,
        layout_item_id=layout_item_id,
    )


def ensure_dashboard_widget_from_template(
    user,
    dashboard_table: DynamicTable,
    template_key: str,
    *,
    existing_records: list[GroupedRecord] | None = None,
) -> GroupedRecord:
    dashboard_records = existing_records if existing_records is not None else get_records(dashboard_table, user)
    for record in dashboard_records:
        if normalize_text(record.values.get("template-key")) == template_key:
            return record
    row_key = save_record(
        dashboard_table,
        user,
        dashboard_template_initial_values(template_key),
    )
    return get_record(dashboard_table, user, row_key)


def add_dashboard_widget_to_board(
    board: DashboardBoard,
    widget_row_key: UUID | str,
    *,
    card_size: str = DashboardBoardItem.CardSize.STANDARD,
    chart_height: str = DashboardBoardItem.ChartHeight.MEDIUM,
) -> DashboardBoardItem:
    allowed_card_sizes = {choice[0] for choice in DashboardBoardItem.CardSize.choices}
    allowed_chart_heights = {choice[0] for choice in DashboardBoardItem.ChartHeight.choices}
    if card_size not in allowed_card_sizes:
        card_size = DashboardBoardItem.CardSize.STANDARD
    if chart_height not in allowed_chart_heights:
        chart_height = DashboardBoardItem.ChartHeight.MEDIUM
    existing = DashboardBoardItem.objects.filter(board=board, widget_row_key=UUID(str(widget_row_key))).first()
    if existing:
        return existing
    next_sort_order = (
        DashboardBoardItem.objects.filter(board=board).order_by("-sort_order").values_list("sort_order", flat=True).first()
        or 0
    ) + 10
    return DashboardBoardItem.objects.create(
        board=board,
        widget_row_key=UUID(str(widget_row_key)),
        sort_order=next_sort_order,
        card_size=card_size,
        chart_height=chart_height,
    )


def ensure_default_dashboard_boards(
    user,
    dashboard_table: DynamicTable,
    widget_records: list[GroupedRecord] | None = None,
) -> list[DashboardBoard]:
    existing_boards = list(DashboardBoard.objects.filter(user=user).prefetch_related("items"))
    if existing_boards:
        return existing_boards

    dashboard_records = widget_records if widget_records is not None else get_records(dashboard_table, user)
    widget_map = {
        normalize_text(record.values.get("template-key")): record
        for record in dashboard_records
        if normalize_text(record.values.get("template-key"))
    }
    visible_widgets = sorted(
        [
            build_dashboard_widget(record, user)
            for record in dashboard_records
            if normalize_text(record.values.get("show-on-dashboard")).lower() != "hide"
        ],
        key=lambda widget: (widget.sort_order, widget.title.lower()),
    )

    created_boards: list[DashboardBoard] = []
    for index, preset in enumerate(DEFAULT_DASHBOARD_BOARD_PRESETS, start=1):
        board = DashboardBoard.objects.create(
            user=user,
            name=preset["name"],
            description=preset["description"],
            show_on_homepage=preset["show_on_homepage"],
            autoplay_enabled=preset["autoplay_enabled"],
            autoplay_interval_seconds=preset["autoplay_interval_seconds"],
            sort_order=preset["sort_order"],
        )
        created_boards.append(board)
        if index == 1 and visible_widgets:
            for widget in visible_widgets:
                DashboardBoardItem.objects.create(
                    board=board,
                    widget_row_key=widget.row_key,
                    sort_order=widget.sort_order,
                    card_size=widget.card_size,
                    chart_height=widget.chart_height,
                )
            continue
        for item_index, widget_spec in enumerate(preset["widgets"], start=1):
            template_key = widget_spec["template_key"]
            widget_record = widget_map.get(template_key)
            if widget_record is None:
                widget_record = ensure_dashboard_widget_from_template(
                    user,
                    dashboard_table,
                    template_key,
                    existing_records=dashboard_records,
                )
                widget_map[template_key] = widget_record
                dashboard_records.append(widget_record)
            DashboardBoardItem.objects.create(
                board=board,
                widget_row_key=widget_record.row_key,
                sort_order=item_index * 10,
                card_size=widget_spec["card_size"],
                chart_height=widget_spec["chart_height"],
            )
    return created_boards


def build_dashboard_board_widgets(
    board: DashboardBoard | None,
    widget_map: dict[UUID, DashboardWidget],
) -> list[DashboardWidget]:
    if board is None:
        return []
    board_items = list(board.items.order_by("sort_order", "id"))
    widgets: list[DashboardWidget] = []
    for item in board_items:
        widget = widget_map.get(item.widget_row_key)
        if widget is None:
            continue
        widgets.append(
            clone_dashboard_widget_with_layout(
                widget,
                card_size=item.card_size,
                chart_height=item.chart_height,
                sort_order=item.sort_order,
                layout_item_id=item.id,
            )
        )
    return widgets


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


def build_applied_flow_sequences(user, applied_records: list[GroupedRecord]) -> list[list[str]]:
    history_by_row = get_application_history_events_map(
        user,
        "applied",
        [record.row_key for record in applied_records],
    )
    sequences: list[list[str]] = []
    for record in applied_records:
        sequence: list[str] = []
        for event in history_by_row.get(record.row_key, []):
            stage = normalize_text(event.to_stage)
            if not stage:
                continue
            if not sequence or sequence[-1] != stage:
                sequence.append(stage)

        if not sequence:
            current_stage = infer_progress(record)
            if current_stage == "Applied":
                sequence = ["Applied"]
            else:
                sequence = ["Applied", current_stage]
        sequences.append(sequence)
    return sequences


def build_applied_flow_chart(applied_records: list[GroupedRecord], user) -> dict[str, Any]:
    if not applied_records:
        return {
            "title": "Applied Flow",
            "description": "Stage-to-stage application flow built from tracked history.",
            "kind": "plotly",
            "chart_type": "sankey",
            "chart_config": {"data": [], "layout": {}},
        }

    sequences = build_applied_flow_sequences(user, applied_records)
    stage_counts = Counter()
    transition_counts = Counter()
    for sequence in sequences:
        for stage in sequence:
            stage_counts[stage] += 1
        for source_stage, target_stage in zip(sequence, sequence[1:]):
            transition_counts[(source_stage, target_stage)] += 1

    if not transition_counts:
        tracked_label = "Tracked"
        current_stage_counts = Counter(infer_progress(record) for record in applied_records)
        node_keys = [tracked_label, *sorted(current_stage_counts, key=application_stage_sort_key)]
        node_labels = [
            f"{tracked_label} ({len(applied_records)})",
            *[
                f"{stage} ({current_stage_counts[stage]})"
                for stage in sorted(current_stage_counts, key=application_stage_sort_key)
            ],
        ]
        node_indices = {label: index for index, label in enumerate(node_keys)}
        source = [node_indices[tracked_label] for _ in current_stage_counts]
        target = [node_indices[stage] for stage in sorted(current_stage_counts, key=application_stage_sort_key)]
        value = [current_stage_counts[stage] for stage in sorted(current_stage_counts, key=application_stage_sort_key)]
        link_colors = [semantic_hex(stage, alpha=0.24) for stage in sorted(current_stage_counts, key=application_stage_sort_key)]
        node_colors = [
            "rgba(14, 165, 233, 0.85)",
            *[semantic_hex(stage, alpha=0.82) for stage in sorted(current_stage_counts, key=application_stage_sort_key)],
        ]
    else:
        ordered_stages = sorted(stage_counts, key=application_stage_sort_key)
        node_keys = ordered_stages
        node_labels = [f"{stage} ({stage_counts[stage]})" for stage in ordered_stages]
        node_indices = {label: index for index, label in enumerate(node_keys)}
        source = []
        target = []
        value = []
        link_colors = []
        for (source_stage, target_stage), count in sorted(
            transition_counts.items(),
            key=lambda item: (
                application_stage_sort_key(item[0][0]),
                application_stage_sort_key(item[0][1]),
            ),
        ):
            source.append(node_indices[source_stage])
            target.append(node_indices[target_stage])
            value.append(count)
            link_colors.append(semantic_hex(target_stage, alpha=0.24))
        node_colors = [semantic_hex(stage, alpha=0.82) for stage in ordered_stages]

    return {
        "title": "Applied Flow",
        "description": "Stage-to-stage application flow built from tracked history.",
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
                        "color": node_colors,
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


def build_average_stage_timing_chart(applied_records: list[GroupedRecord], user) -> dict[str, Any]:
    history_by_row = get_application_history_events_map(
        user,
        "applied",
        [record.row_key for record in applied_records],
    )
    durations_by_stage: dict[str, list[float]] = defaultdict(list)

    for record in applied_records:
        events = history_by_row.get(record.row_key, [])
        baseline = earliest_tracked_job_timestamp(record, events)
        if baseline is None:
            continue
        for stage, occurred_at in first_stage_timestamps(record, events).items():
            if stage == "To Apply" or occurred_at < baseline:
                continue
            durations_by_stage[stage].append((occurred_at - baseline).total_seconds() / 86400)

    ordered_stages = [
        stage
        for stage in sorted(durations_by_stage, key=application_stage_sort_key)
        if durations_by_stage[stage]
    ]
    labels = ordered_stages
    values = [round(sum(durations_by_stage[stage]) / len(durations_by_stage[stage]), 1) for stage in ordered_stages]
    return {
        "labels": labels,
        "values": values,
        "chart_config": {
            "labels": labels,
            "datasets": [
                {
                    "label": "Average days from first tracked step",
                    "data": values,
                    "backgroundColor": [semantic_hex(stage, alpha=0.72) for stage in labels],
                    "borderColor": [semantic_hex(stage) for stage in labels],
                    "borderWidth": 1,
                    "borderRadius": 8,
                }
            ],
        },
    }


def is_active_applied_record(record: GroupedRecord) -> bool:
    inactive_values = {"rejected", "withdrawn"}
    result = normalize_text(record.values.get("result")).lower()
    progress = infer_progress(record).lower()
    return result not in inactive_values and progress not in inactive_values


def build_active_applied_pipeline_chart(applied_records: list[GroupedRecord]) -> dict[str, Any]:
    active_records = [record for record in applied_records if is_active_applied_record(record)]
    progress_counts = Counter(infer_progress(record) for record in active_records)
    labels = [
        stage
        for stage in sorted(progress_counts, key=application_stage_sort_key)
        if progress_counts[stage]
    ]
    values = [float(progress_counts[label]) for label in labels]
    colors = [semantic_hex(label, alpha=0.72) for label in labels]
    return {
        "labels": labels,
        "values": values,
        "chart_config": {
            "labels": labels,
            "datasets": [
                {
                    "label": "Active applications",
                    "data": values,
                    "backgroundColor": colors,
                    "borderColor": [semantic_hex(label) for label in labels],
                    "borderWidth": 1,
                    "borderRadius": 8,
                }
            ],
        },
    }


def deadline_lead_bucket(lead_days: int) -> str:
    if lead_days < 0:
        return "Applied late"
    if lead_days == 0:
        return "On deadline"
    if lead_days <= 3:
        return "1-3 days before"
    if lead_days <= 7:
        return "4-7 days before"
    if lead_days <= 14:
        return "8-14 days before"
    return "15+ days before"


def build_deadline_lead_time_chart(applied_records: list[GroupedRecord]) -> dict[str, Any]:
    bucket_labels = [
        "Applied late",
        "On deadline",
        "1-3 days before",
        "4-7 days before",
        "8-14 days before",
        "15+ days before",
    ]
    bucket_colors = {
        "Applied late": "rgba(244, 63, 94, 0.72)",
        "On deadline": "rgba(245, 158, 11, 0.72)",
        "1-3 days before": "rgba(249, 115, 22, 0.72)",
        "4-7 days before": "rgba(14, 165, 233, 0.72)",
        "8-14 days before": "rgba(16, 185, 129, 0.72)",
        "15+ days before": "rgba(34, 197, 94, 0.72)",
    }
    bucket_counts = {label: 0 for label in bucket_labels}

    for record in applied_records:
        applied_date = normalize_date_value(record.values.get("date"))
        deadline_date = normalize_date_value(record.values.get("deadline"))
        if not applied_date or not deadline_date:
            continue
        try:
            applied_on = datetime.fromisoformat(applied_date).date()
            deadline_on = datetime.fromisoformat(deadline_date).date()
        except ValueError:
            continue
        bucket_counts[deadline_lead_bucket((deadline_on - applied_on).days)] += 1

    labels = bucket_labels
    values = [bucket_counts[label] for label in labels]
    return {
        "labels": labels,
        "values": values,
        "chart_config": {
            "labels": labels,
            "datasets": [
                {
                    "label": "Applications",
                    "data": values,
                    "backgroundColor": [bucket_colors[label] for label in labels],
                    "borderColor": [bucket_colors[label].replace("0.72", "1") for label in labels],
                    "borderWidth": 1,
                    "borderRadius": 8,
                }
            ],
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
        if template_key == "applied-active-pipeline" and source_table_slug == "applied":
            filters = ""
            special_chart = build_active_applied_pipeline_chart(source_records)
            labels = special_chart["labels"]
            values = special_chart["values"]
            chart_config = special_chart["chart_config"]
            source_records = [record for record in source_records if is_active_applied_record(record)]
        filter_map = parse_filter_text(filters)
        if filter_map:
            source_records = [item for item in source_records if record_matches_filters(item, filter_map)]
        if chart_config:
            pass
        elif template_key == "applied-stage-timing" and source_table_slug == "applied":
            special_chart = build_average_stage_timing_chart(source_records, user)
            labels = special_chart["labels"]
            values = special_chart["values"]
            chart_config = special_chart["chart_config"]
        elif template_key == "applied-deadline-lead" and source_table_slug == "applied":
            special_chart = build_deadline_lead_time_chart(source_records)
            labels = special_chart["labels"]
            values = special_chart["values"]
            chart_config = special_chart["chart_config"]
        else:
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
                chart_config = build_applied_flow_chart(source_records, user)["chart_config"]
            else:
                chart_config = {"data": [], "layout": {}}
        elif chart_type == "waterfall":
            chart_config = build_waterfall_payload(labels, values)
        elif not chart_config:
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


def build_dashboard(user, selected_board_id: int | None = None) -> dict[str, Any]:
    workbook_path = Path("Applications.xlsx") if Path("Applications.xlsx").exists() else None
    tables = ensure_user_workspace(user, workbook_path)
    to_apply_records = get_records(tables["to-apply"], user)
    applied_records = get_records(tables["applied"], user)
    ensure_default_dashboard_widgets(user, tables["dashboard"])
    dashboard_records = get_records(tables["dashboard"], user)
    portal_records = get_records(tables["portals"], user)
    all_dashboard_widgets = [build_dashboard_widget(record, user) for record in dashboard_records]
    widget_map = {widget.row_key: widget for widget in all_dashboard_widgets}
    dashboard_widgets = sorted(
        [widget for widget in all_dashboard_widgets if widget.show_on_dashboard],
        key=lambda widget: (widget.sort_order, widget.title.lower()),
    )
    hidden_dashboard_widgets = sorted(
        [widget for widget in all_dashboard_widgets if not widget.show_on_dashboard],
        key=lambda widget: (widget.sort_order, widget.title.lower()),
    )
    ensure_default_dashboard_boards(user, tables["dashboard"], dashboard_records)
    dashboard_boards = list(DashboardBoard.objects.filter(user=user).prefetch_related("items").order_by("sort_order", "name", "id"))
    homepage_boards = [board for board in dashboard_boards if board.show_on_homepage] or dashboard_boards[:1]
    selected_homepage_board = next(
        (board for board in dashboard_boards if selected_board_id and board.id == selected_board_id),
        homepage_boards[0] if homepage_boards else None,
    )
    selected_homepage_widgets = build_dashboard_board_widgets(selected_homepage_board, widget_map)
    if selected_homepage_board is None and not selected_homepage_widgets:
        selected_homepage_widgets = dashboard_widgets
    dashboard_panel_layout = build_dashboard_panel_layout(user)

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
        "all_dashboard_widgets": sorted(all_dashboard_widgets, key=lambda widget: (widget.sort_order, widget.title.lower())),
        "dashboard_widgets": dashboard_widgets,
        "hidden_dashboard_widgets": hidden_dashboard_widgets,
        "dashboard_templates": build_dashboard_template_catalog(dashboard_records),
        "dashboard_boards": dashboard_boards,
        "homepage_boards": homepage_boards,
        "selected_homepage_board": selected_homepage_board,
        "selected_homepage_widgets": selected_homepage_widgets,
        "dashboard_panel_layout": dashboard_panel_layout,
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
        row_key=source_row_key or None,
        metadata_map=source_metadata,
    )
    if source_row_key:
        move_application_related_data(user, to_apply_table.slug, applied_table.slug, source_row_key)
        delete_record(to_apply_table, user, source_row_key, cleanup_related=False)
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


def seed_sample_jobs_for_user(user) -> None:
    tables = ensure_user_workspace(user)
    applied_table = tables["applied"]
    sample_prefix = "https://example.com/sample-job-tracker/"

    existing_records = [
        *get_records(applied_table, user),
        *get_records(tables["to-apply"], user),
    ]
    for record in existing_records:
        if normalize_text(record.values.get("url")).startswith(sample_prefix):
            delete_record(record.table, user, record.row_key)

    now = timezone.now()
    sample_jobs = [
        {
            "table": applied_table,
            "values": {
                "date": (timezone.localdate() - timedelta(days=18)).isoformat(),
                "deadline": (timezone.localdate() + timedelta(days=4)).isoformat(),
                "role-title": "Dummy Vision PhD",
                "url": f"{sample_prefix}vision-phd",
                "type": "PhD",
                "number-of-positions": 2,
                "application-progress": "Screening",
                "remarks": "Demo sample with early-stage screening progress.",
            },
            "tags": ["Dummy", "PhD", "Priority"],
            "history": [
                (ApplicationHistoryEvent.EventType.ADDED_TO_APPLY, "", "To Apply", now - timedelta(days=21)),
                (ApplicationHistoryEvent.EventType.APPLIED, "", "Applied", now - timedelta(days=18)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Applied", "Screening", now - timedelta(days=14)),
            ],
        },
        {
            "table": applied_table,
            "values": {
                "date": (timezone.localdate() - timedelta(days=16)).isoformat(),
                "deadline": (timezone.localdate() - timedelta(days=13)).isoformat(),
                "role-title": "Dummy Research Engineer",
                "url": f"{sample_prefix}research-engineer",
                "type": "Job",
                "number-of-positions": 1,
                "application-progress": "Interview",
                "remarks": "Demo sample currently in the interview loop.",
            },
            "tags": ["Dummy", "Job"],
            "history": [
                (ApplicationHistoryEvent.EventType.ADDED_TO_APPLY, "", "To Apply", now - timedelta(days=20)),
                (ApplicationHistoryEvent.EventType.APPLIED, "", "Applied", now - timedelta(days=16)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Applied", "Screening", now - timedelta(days=12)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Screening", "Interview", now - timedelta(days=8)),
            ],
        },
        {
            "table": applied_table,
            "values": {
                "date": (timezone.localdate() - timedelta(days=13)).isoformat(),
                "deadline": (timezone.localdate() - timedelta(days=13)).isoformat(),
                "role-title": "Dummy Applied Scientist",
                "url": f"{sample_prefix}applied-scientist",
                "type": "Job",
                "number-of-positions": 1,
                "application-progress": "Technical Interview",
                "remarks": "Demo sample with a technical round in progress.",
            },
            "tags": ["Dummy", "Job", "Priority"],
            "history": [
                (ApplicationHistoryEvent.EventType.ADDED_TO_APPLY, "", "To Apply", now - timedelta(days=15)),
                (ApplicationHistoryEvent.EventType.APPLIED, "", "Applied", now - timedelta(days=13)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Applied", "Screening", now - timedelta(days=11)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Screening", "Interview", now - timedelta(days=7)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Interview", "Technical Interview", now - timedelta(days=4)),
            ],
        },
        {
            "table": applied_table,
            "values": {
                "date": (timezone.localdate() - timedelta(days=14)).isoformat(),
                "deadline": (timezone.localdate() - timedelta(days=16)).isoformat(),
                "role-title": "Dummy Robotics Role",
                "url": f"{sample_prefix}robotics-role",
                "type": "Job",
                "number-of-positions": 3,
                "application-progress": "Offer",
                "result": "Offer",
                "result-on": (timezone.localdate() - timedelta(days=2)).isoformat(),
                "remarks": "Demo sample with a positive outcome.",
            },
            "tags": ["Dummy", "Job"],
            "history": [
                (ApplicationHistoryEvent.EventType.ADDED_TO_APPLY, "", "To Apply", now - timedelta(days=17)),
                (ApplicationHistoryEvent.EventType.APPLIED, "", "Applied", now - timedelta(days=14)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Applied", "Screening", now - timedelta(days=10)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Screening", "Interview", now - timedelta(days=6)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Interview", "Offer", now - timedelta(days=2)),
            ],
        },
        {
            "table": applied_table,
            "values": {
                "date": (timezone.localdate() - timedelta(days=11)).isoformat(),
                "deadline": (timezone.localdate() - timedelta(days=4)).isoformat(),
                "role-title": "Dummy Platform Engineer",
                "url": f"{sample_prefix}platform-engineer",
                "type": "Job",
                "number-of-positions": 2,
                "application-progress": "Assignment",
                "remarks": "Demo sample waiting on an assignment submission result.",
            },
            "tags": ["Dummy", "Job", "Priority"],
            "history": [
                (ApplicationHistoryEvent.EventType.ADDED_TO_APPLY, "", "To Apply", now - timedelta(days=13)),
                (ApplicationHistoryEvent.EventType.APPLIED, "", "Applied", now - timedelta(days=11)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Applied", "Screening", now - timedelta(days=8)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Screening", "Interview", now - timedelta(days=5)),
                (ApplicationHistoryEvent.EventType.STAGE_CHANGE, "Interview", "Assignment", now - timedelta(days=1)),
            ],
        },
    ]

    for sample in sample_jobs:
        row_key = save_record(
            sample["table"],
            user,
            sample["values"],
            skip_history=True,
        )
        sync_application_tags(user, sample["table"].slug, row_key, sample["tags"])
        for event_type, from_stage, to_stage, occurred_at in sample["history"]:
            record_application_history_event(
                user,
                sample["table"].slug,
                row_key,
                event_type,
                from_stage=from_stage,
                to_stage=to_stage,
                occurred_at=occurred_at,
            )


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
    seed_sample_jobs_for_user(test_user)
    return {
        "admin": admin,
        "test_user": test_user,
        "imported_counts": imported_counts,
    }
