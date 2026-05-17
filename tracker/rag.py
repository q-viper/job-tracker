from __future__ import annotations

import math
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import timedelta
from functools import lru_cache
from typing import Any

from django.db import transaction
from django.utils import timezone

from .models import ChatMessage, ChatModelConfig, DynamicTable, PortalJobSuggestion
from .services import (
    add_portal_job_to_to_apply,
    ensure_user_workspace,
    get_records,
    get_user_table,
    infer_application_type,
    refresh_portal_jobs,
)

RAG_DEFAULT_PROMPTS = [
    "Fetch the latest portal jobs and filter by computer vision.",
    "What are the PhD positions available with computer vision?",
    "Show me applied roles that are currently in interview stages.",
    "Match my uploaded CV with the latest relevant jobs.",
    "Which deadlines are coming up this week?",
]

SHORT_QUERY_TOKENS = {
    "ai",
    "ml",
    "cv",
    "nlp",
    "llm",
    "ui",
    "ux",
    "qa",
}

STOPWORDS = {
    "the",
    "and",
    "for",
    "with",
    "that",
    "this",
    "from",
    "into",
    "what",
    "which",
    "show",
    "me",
    "are",
    "all",
    "jobs",
    "job",
    "role",
    "roles",
    "available",
    "latest",
    "your",
    "please",
    "then",
    "than",
    "them",
}

PROFILE_STOPWORDS = STOPWORDS | {
    "curriculum",
    "vitae",
    "resume",
    "email",
    "phone",
    "mobile",
    "linkedin",
    "github",
    "experience",
    "education",
    "skills",
    "projects",
    "project",
    "research",
    "work",
    "employment",
    "profile",
    "summary",
    "objective",
    "address",
    "contact",
}

COVER_LETTER_TERMS = (
    "cover letter",
    "motivation letter",
    "draft a letter",
    "write a letter",
)

DOCUMENT_SUMMARIZE_TERMS = (
    "summarize",
    "summary",
    "summarise",
    "profile me",
    "bio",
    "elevator pitch",
)

DOCUMENT_JOB_MATCH_TERMS = (
    "match my cv",
    "match my resume",
    "match my profile",
    "matching jobs",
    "matching roles",
    "relevant jobs",
    "relevant roles",
    "suitable jobs",
    "suitable roles",
    "find jobs",
    "find roles",
    "what jobs",
    "which jobs",
    "what roles",
    "which roles",
)

GREETING_TERMS = (
    "hi",
    "hello",
    "hey",
    "good morning",
    "good afternoon",
    "good evening",
)

HELP_TERMS = (
    "help",
    "what can you do",
    "how can you help",
    "who are you",
)

THANKS_TERMS = (
    "thanks",
    "thank you",
)

COUNT_FILTER_IGNORE_TOKENS = {
    "count",
    "counts",
    "did",
    "do",
    "does",
    "done",
    "how",
    "have",
    "has",
    "had",
    "use",
    "used",
    "using",
    "time",
    "times",
    "often",
    "number",
    "numbers",
    "many",
    "much",
    "total",
    "totals",
    "current",
    "currently",
    "apply",
    "applied",
    "application",
    "applications",
    "submitted",
    "submit",
    "submission",
    "submissions",
    "entry",
    "entries",
    "item",
    "items",
    "record",
    "records",
    "list",
    "my",
}


@dataclass
class RagDocument:
    source_label: str
    source_type: str
    title: str
    text: str
    url: str
    metadata: dict[str, Any] = field(default_factory=dict)


def tokenize(text: str) -> list[str]:
    return [
        token
        for token in re.findall(r"[a-zA-Z0-9\-\+]+", text.lower())
        if (
            token not in STOPWORDS
            and (len(token) > 2 or token.isdigit() or token in SHORT_QUERY_TOKENS)
        )
    ]


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def contains_any_term(text: str, terms: tuple[str, ...]) -> bool:
    lowered = text.lower()
    return any(term in lowered for term in terms)


def normalized_intent_query(query: str) -> str:
    return re.sub(r"\s+", " ", query.lower()).strip(" .,!?:;\n\t")


def build_document_text(title: str, values: dict[str, str]) -> str:
    lines = [f"Title: {title}"]
    for key, value in values.items():
        if value:
            lines.append(f"{key}: {value}")
    return "\n".join(lines)


def build_record_search_text(record) -> str:
    return build_document_text(record.label, record.values).lower()


def purge_expired_chat_history(user) -> int:
    retention_days = max(int(getattr(user, "chat_history_retention_days", 1) or 1), 1)
    cutoff = timezone.now() - timedelta(days=retention_days)
    deleted, _ = ChatMessage.objects.filter(user=user, created_at__lt=cutoff).delete()
    return deleted


def build_user_documents(user) -> list[RagDocument]:
    ensure_user_workspace(user)
    documents: list[RagDocument] = []

    tables = DynamicTable.objects.filter(owner=user).exclude(slug="dashboard").prefetch_related("columns")
    for table in tables:
        for record in get_records(table, user):
            documents.append(
                RagDocument(
                    source_label=table.name,
                    source_type=table.slug,
                    title=record.label,
                    text=build_document_text(record.label, record.values),
                    url=record.values.get("url", ""),
                    metadata={
                        "row_key": str(record.row_key),
                        "values": record.values,
                    },
                )
            )

    jobs = PortalJobSuggestion.objects.filter(user=user).order_by("-fetched_at", "title")
    for job in jobs:
        text = "\n".join(
            [
                f"Title: {job.title}",
                f"Portal: {job.portal_name or job.portal_url}",
                f"Summary: {job.summary}",
                f"URL: {job.url}",
            ]
        )
        documents.append(
            RagDocument(
                source_label="Fetched Portal Jobs",
                source_type="portal-jobs",
                title=job.title,
                text=text,
                url=job.url,
                metadata={
                    "portal_job_id": job.id,
                    "added_to_apply": job.added_to_apply,
                    "portal_name": job.portal_name,
                },
            )
        )
    return documents


def score_documents(query: str, documents: list[RagDocument], limit: int = 8) -> list[tuple[RagDocument, float]]:
    query_tokens = tokenize(query)
    if not query_tokens:
        return []

    doc_tokens = [tokenize(document.text) for document in documents]
    document_frequency = Counter()
    for tokens in doc_tokens:
        document_frequency.update(set(tokens))

    total_docs = max(len(documents), 1)
    scored: list[tuple[RagDocument, float]] = []
    lowered_query = query.lower()

    for document, tokens in zip(documents, doc_tokens):
        token_counts = Counter(tokens)
        score = 0.0
        for token in query_tokens:
            if token not in token_counts:
                continue
            idf = math.log((1 + total_docs) / (1 + document_frequency[token])) + 1
            score += (1 + math.log(token_counts[token])) * idf
        title_lower = document.title.lower()
        if any(token in title_lower for token in query_tokens):
            score += 1.5
        if lowered_query and lowered_query in document.text.lower():
            score += 2.0
        if document.source_type == "portal-jobs" and ("latest" in lowered_query or "portal" in lowered_query):
            score += 0.8
        if score > 0:
            scored.append((document, score))

    scored.sort(key=lambda item: item[1], reverse=True)
    return scored[:limit]


def should_refresh_portals(query: str) -> bool:
    lowered = query.lower()
    refresh_terms = ("fetch", "refresh", "latest")
    target_terms = ("job", "jobs", "portal", "portals")
    return any(term in lowered for term in refresh_terms) and any(term in lowered for term in target_terms)


def refresh_all_portals(user) -> int:
    try:
        portals_table = get_user_table(user, "portals")
    except DynamicTable.DoesNotExist:
        return 0

    refreshed = 0
    for record in get_records(portals_table, user):
        try:
            refresh_portal_jobs(user, portals_table, record.row_key)
            refreshed += 1
        except Exception:
            continue
    return refreshed


def extract_text_from_uploaded_file(uploaded_file) -> tuple[str, str]:
    if uploaded_file is None:
        return "", ""

    file_name = normalize_text(getattr(uploaded_file, "name", "document"))
    suffix = os.path.splitext(file_name.lower())[1]
    raw_bytes = uploaded_file.read()

    if suffix in {".txt", ".md", ".csv"}:
        return file_name, raw_bytes.decode("utf-8", errors="ignore")[:24000]

    if suffix == ".pdf":
        try:
            from io import BytesIO
            from pypdf import PdfReader

            reader = PdfReader(BytesIO(raw_bytes))
            text = "\n".join((page.extract_text() or "") for page in reader.pages)
            return file_name, text[:24000]
        except Exception:
            return file_name, raw_bytes.decode("utf-8", errors="ignore")[:24000]

    if suffix == ".docx":
        try:
            from io import BytesIO
            from docx import Document

            document = Document(BytesIO(raw_bytes))
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            return file_name, text[:24000]
        except Exception:
            return file_name, raw_bytes.decode("utf-8", errors="ignore")[:24000]

    return file_name, raw_bytes.decode("utf-8", errors="ignore")[:24000]


def top_keywords_from_text(text: str, limit: int = 18) -> list[str]:
    counts = Counter(tokenize(text))
    return [token for token, _ in counts.most_common(limit)]


def profile_keywords_from_text(text: str, limit: int = 6) -> list[str]:
    return [
        token
        for token in top_keywords_from_text(text, limit=limit * 3)
        if token not in PROFILE_STOPWORDS
    ][:limit]


def readable_keyword(token: str) -> str:
    return token.replace("-", " ").replace("+", " + ").strip().title()


def extract_candidate_name(document_text: str) -> str:
    for raw_line in document_text.splitlines()[:6]:
        line = normalize_text(raw_line)
        if not line or "@" in line or "http" in line or any(char.isdigit() for char in line):
            continue
        words = line.split()
        if 1 < len(words) <= 5:
            return line
    return "[Your Name]"


def extract_requested_role(query: str) -> str:
    match = re.search(
        r"(?:cover letter|motivation letter|letter)\s+(?:for|to apply for)\s+(?P<role>.+)",
        query,
        re.IGNORECASE,
    )
    if not match:
        return ""
    role = re.split(r"[.?!]", match.group("role"), maxsplit=1)[0]
    role = re.sub(r"^(?:a|an|the)\s+", "", role.strip(), flags=re.IGNORECASE)
    return role.strip()


def is_cover_letter_query(query: str) -> bool:
    lowered = query.lower()
    if contains_any_term(lowered, COVER_LETTER_TERMS):
        return True
    return "cover" in lowered and "letter" in lowered


def is_document_summary_query(query: str, *, document_text: str = "") -> bool:
    return bool(document_text) and contains_any_term(query, DOCUMENT_SUMMARIZE_TERMS)


def is_document_matching_query(query: str, *, document_text: str = "") -> bool:
    if not document_text or is_cover_letter_query(query):
        return False
    lowered = query.lower()
    if contains_any_term(lowered, DOCUMENT_JOB_MATCH_TERMS):
        return True
    has_job_target = any(term in lowered for term in ("job", "jobs", "role", "roles", "position", "positions", "opening", "openings"))
    has_matching_language = any(term in lowered for term in ("match", "matching", "fit", "relevant", "suitable", "aligned"))
    return has_job_target and has_matching_language


def is_greeting_query(query: str) -> bool:
    normalized = normalized_intent_query(query)
    if not normalized:
        return False
    if normalized in GREETING_TERMS:
        return True
    return any(
        normalized.startswith(f"{term} ") or normalized == term
        for term in GREETING_TERMS
    ) and len(normalized.split()) <= 4


def is_help_query(query: str) -> bool:
    normalized = normalized_intent_query(query)
    return bool(normalized) and contains_any_term(normalized, HELP_TERMS)


def is_thanks_query(query: str) -> bool:
    normalized = normalized_intent_query(query)
    return bool(normalized) and contains_any_term(normalized, THANKS_TERMS)


def smalltalk_answer_for_query(query: str) -> str | None:
    if is_greeting_query(query):
        return (
            "Hi. I can help you count applications, filter PhD roles, check deadlines, "
            "review portal jobs, and use an uploaded CV for matching or drafting."
        )
    if is_help_query(query):
        return (
            "I can count your applied or To Apply records, filter by things like PhD or interview stage, "
            "fetch portal jobs, summarize deadlines, and use an uploaded CV for job matching or cover-letter drafts."
        )
    if is_thanks_query(query):
        return "You're welcome. If you want, I can also summarize your pipeline or count a filtered set of roles."
    return None


def build_cover_letter_from_document(query: str, document_text: str) -> str:
    role = extract_requested_role(query) or "the role"
    candidate_name = extract_candidate_name(document_text)
    keywords = profile_keywords_from_text(document_text, limit=5)
    focus = ", ".join(readable_keyword(token) for token in keywords[:3]) or "research, collaboration, and technical problem solving"
    secondary_focus = ", ".join(readable_keyword(token) for token in keywords[3:5]) or "hands-on execution and clear communication"
    return "\n".join(
        [
            "Dear Hiring Team,",
            "",
            f"I am writing to express my interest in {role}. My background reflects strong experience in {focus}, and I am excited about the opportunity to contribute that foundation in a thoughtful, practical way.",
            "",
            f"Across my work and academic projects, I have built a profile that combines analytical depth with execution. I am especially motivated by opportunities that value {secondary_focus}, continuous learning, and the ability to turn ideas into reliable outcomes.",
            "",
            "I would welcome the chance to bring this background, curiosity, and commitment to your team. Thank you for your time and consideration, and I would be glad to discuss my fit in more detail.",
            "",
            "Sincerely,",
            candidate_name,
        ]
    )


def build_document_summary(document_name: str, document_text: str) -> str:
    keywords = profile_keywords_from_text(document_text, limit=6)
    if not keywords:
        return f"I parsed {document_name}, but I could not extract enough detail for a useful summary."
    readable = ", ".join(readable_keyword(token) for token in keywords[:4])
    return (
        f"I parsed {document_name}. The strongest themes I found are {readable}. "
        "If you want, I can turn that into a cover letter, short bio, or job-match query."
    )


def shared_keywords(document_text: str, target_text: str, limit: int = 5) -> list[str]:
    left = set(top_keywords_from_text(document_text, limit=24))
    right = set(top_keywords_from_text(target_text, limit=24))
    shared = sorted(left & right)
    return shared[:limit]


def augment_query_with_document(query: str, document_text: str) -> str:
    if not document_text:
        return query
    keywords = top_keywords_from_text(document_text, limit=24)
    if not keywords:
        return query
    return f"{query} {' '.join(keywords)}"


def relevant_matches_for_intent(
    query: str,
    matches: list[tuple[RagDocument, float]],
    *,
    document_text: str = "",
) -> list[tuple[RagDocument, float]]:
    lowered = query.lower()
    if is_document_matching_query(query, document_text=document_text):
        filtered = [
            item for item in matches if item[0].source_type in {"portal-jobs", "to-apply", "portals"}
        ]
        return filtered or matches
    if "deadline" in lowered:
        filtered = [item for item in matches if item[0].source_type == "to-apply"]
        return filtered or matches
    if "interview" in lowered or "progress" in lowered or "stage" in lowered:
        filtered = [item for item in matches if item[0].source_type == "applied"]
        return filtered or matches
    if "applied" in lowered:
        filtered = [item for item in matches if item[0].source_type == "applied"]
        return filtered or matches
    if "to apply" in lowered:
        filtered = [item for item in matches if item[0].source_type == "to-apply"]
        return filtered or matches
    if "portal" in lowered or "latest" in lowered:
        filtered = [item for item in matches if item[0].source_type == "portal-jobs"]
        return filtered or matches
    return matches


def is_count_query(query: str) -> bool:
    lowered = normalized_intent_query(query)
    return (
        "how many" in lowered
        or "count" in lowered
        or "counts" in lowered
        or "number of" in lowered
        or "total" in lowered
    )


def count_scope_for_query(query: str) -> str:
    lowered = normalized_intent_query(query)
    if "portal" in lowered or "fetched jobs" in lowered or "latest jobs" in lowered:
        return "portal-jobs"
    if "to apply" in lowered:
        return "to-apply"
    if any(term in lowered for term in ("applied", "apply", "submitted", "submission")):
        return "applied"
    return "all"


def count_filter_tokens(query: str) -> list[str]:
    return [
        token
        for token in tokenize(query)
        if token not in COUNT_FILTER_IGNORE_TOKENS
    ]


def format_count_filter(tokens: list[str]) -> str:
    if not tokens:
        return ""
    return " ".join(readable_keyword(token) for token in tokens)


def records_matching_tokens(records, tokens: list[str]) -> list[Any]:
    if not tokens:
        return list(records)
    matched = []
    for record in records:
        haystack = build_record_search_text(record)
        if all(token in haystack for token in tokens):
            matched.append(record)
    return matched


def records_matching_ai_reliance(records, numeric_tokens: list[str]) -> list[Any]:
    matched = []
    for record in records:
        ai_reliance = normalize_text(record.values.get("ai-reliance")).lower()
        if not ai_reliance:
            continue
        if numeric_tokens and not all(token in ai_reliance for token in numeric_tokens):
            continue
        matched.append(record)
    return matched


def portal_jobs_matching_tokens(jobs, tokens: list[str]) -> list[Any]:
    if not tokens:
        return list(jobs)
    matched = []
    for job in jobs:
        haystack = " ".join(
            [
                normalize_text(getattr(job, "title", "")),
                normalize_text(getattr(job, "summary", "")),
                normalize_text(getattr(job, "portal_name", "")),
            ]
        ).lower()
        if all(token in haystack for token in tokens):
            matched.append(job)
    return matched


def count_answer_for_query(user, query: str) -> str | None:
    lowered = query.lower()
    if not is_count_query(query):
        return None

    filter_tokens = count_filter_tokens(query)
    filter_label = format_count_filter(filter_tokens)
    scope = count_scope_for_query(query)
    numeric_filter_tokens = [token for token in filter_tokens if token.isdigit()]
    mentions_ai_reliance = any(token in {"ai", "ai-reliance", "reliance"} for token in filter_tokens)

    if mentions_ai_reliance:
        applied_records = get_records(get_user_table(user, "applied"), user)
        records = records_matching_ai_reliance(applied_records, numeric_filter_tokens)
        remaining_tokens = [
            token
            for token in filter_tokens
            if token not in {"ai", "ai-reliance", "reliance"} and token not in numeric_filter_tokens
        ]
        if remaining_tokens:
            records = records_matching_tokens(records, remaining_tokens)
        ai_label = (
            f"AI reliance {' / '.join(numeric_filter_tokens)}"
            if numeric_filter_tokens
            else "a saved AI reliance value"
        )
        extra_label = format_count_filter(remaining_tokens)
        return (
            f"You have {len(records)} applied roles with {ai_label}"
            f"{' matching ' + extra_label if extra_label else ''}."
        )

    if "interview" in lowered:
        records = get_records(get_user_table(user, "applied"), user)
        if filter_tokens:
            records = records_matching_tokens(records, filter_tokens)
        count = sum(
            1
            for record in records
            if "interview" in (record.values.get("application-progress") or "").lower()
            or "interview" in (record.values.get("completed-stages") or "").lower()
        )
        return (
            f"You currently have {count} applied roles in interview-related stages"
            f"{' matching ' + filter_label if filter_label else ''}."
        )
    if "deadline" in lowered:
        records = get_records(get_user_table(user, "to-apply"), user)
        if filter_tokens:
            records = records_matching_tokens(records, filter_tokens)
        count = sum(1 for record in records if normalize_text(record.values.get("deadline")))
        return (
            f"You have {count} To Apply entries with a saved deadline"
            f"{' matching ' + filter_label if filter_label else ''}."
        )
    if scope == "portal-jobs":
        jobs = portal_jobs_matching_tokens(PortalJobSuggestion.objects.filter(user=user), filter_tokens)
        return (
            f"I can see {len(jobs)} fetched portal job suggestions"
            f"{' matching ' + filter_label if filter_label else ''} right now."
        )
    if scope == "to-apply":
        records = records_matching_tokens(get_records(get_user_table(user, "to-apply"), user), filter_tokens)
        return (
            f"You currently have {len(records)} roles in your To Apply list"
            f"{' matching ' + filter_label if filter_label else ''}."
        )
    if scope == "applied":
        records = records_matching_tokens(get_records(get_user_table(user, "applied"), user), filter_tokens)
        return (
            f"You have applied to {len(records)} roles so far"
            f"{' matching ' + filter_label if filter_label else ''}."
        )
    applied_records = records_matching_tokens(get_records(get_user_table(user, "applied"), user), filter_tokens)
    to_apply_records = records_matching_tokens(get_records(get_user_table(user, "to-apply"), user), filter_tokens)
    total = len(applied_records) + len(to_apply_records)
    return (
        f"I found {total} tracked application records across Applied and To Apply"
        f"{' matching ' + filter_label if filter_label else ''}."
    )


def build_heuristic_answer(
    user,
    query: str,
    matches: list[tuple[RagDocument, float]],
    *,
    document_name: str = "",
    document_text: str = "",
) -> str | None:
    smalltalk_answer = smalltalk_answer_for_query(query)
    if smalltalk_answer:
        return smalltalk_answer
    count_answer = count_answer_for_query(user, query)
    if count_answer:
        return count_answer

    lowered = query.lower()
    if document_text and is_cover_letter_query(query):
        return build_cover_letter_from_document(query, document_text)
    if document_text and is_document_summary_query(query, document_text=document_text):
        return build_document_summary(document_name or "your uploaded document", document_text)
    if not matches:
        if is_document_matching_query(query, document_text=document_text):
            return "I parsed your uploaded document, but I could not find matching jobs in your current workspace."
        return "I could not find matching records in your current tables or fetched portal jobs."
    filtered_matches = relevant_matches_for_intent(query, matches, document_text=document_text)
    if not filtered_matches:
        filtered_matches = matches

    if document_text and is_document_matching_query(query, document_text=document_text):
        lines = [f"Based on {document_name or 'your uploaded document'}, these roles look strongest:"]
        for document, _ in filtered_matches[:5]:
            overlap = shared_keywords(document_text, document.text)
            reason = f"Matched keywords: {', '.join(overlap)}" if overlap else "Matched against your document themes."
            line = f"- {document.title} [{document.source_label}]"
            if document.url:
                line += f" - {document.url}"
            lines.append(line)
            lines.append(f"  {reason}")
        return "\n".join(lines)

    if "deadline" in lowered:
        lines = ["These are the closest upcoming deadlines I found:"]
        for document, _ in filtered_matches[:5]:
            deadline = document.metadata.get("values", {}).get("deadline", "No deadline saved")
            line = f"- {document.title} [{deadline}]"
            if document.url:
                line += f" - {document.url}"
            lines.append(line)
        return "\n".join(lines)

    if "interview" in lowered or "progress" in lowered or "stage" in lowered:
        interview_matches = [
            item
            for item in filtered_matches
            if "interview" in item[0].text.lower() or "screening" in item[0].text.lower()
        ]
        active_matches = interview_matches or filtered_matches
        lines = ["Here are the applied roles that look active in your pipeline:"]
        for document, _ in active_matches[:5]:
            progress = document.metadata.get("values", {}).get("application-progress", "Applied")
            lines.append(f"- {document.title} [{progress}]")
        return "\n".join(lines)

    if any(term in lowered for term in {"phd", "computer vision", "filter", "match", "find", "what are", "which"}):
        lines = ["These are the strongest matches I found:"]
        for document, _ in filtered_matches[:5]:
            line = f"- {document.title} [{document.source_label}]"
            if document.url:
                line += f" - {document.url}"
            lines.append(line)
        return "\n".join(lines)

    return None


def answer_with_ai(
    query: str,
    matches: list[tuple[RagDocument, float]],
    *,
    model_name: str = "",
    document_name: str = "",
    document_text: str = "",
) -> str | None:
    if not matches and not document_text:
        return None

    context_parts: list[str] = []
    if document_text:
        context_parts.append(f"[Uploaded document: {document_name or 'document'}]\n{document_text[:3500]}")
    if matches:
        context_parts.extend(
            f"[{index + 1}] {document.text[:260]}"
            for index, (document, _) in enumerate(matches[:3])
        )
    context = "\n\n".join(context_parts)
    if document_text:
        prompt = (
            "You are Quassi, a job search assistant. Use the uploaded document and any provided workspace context "
            "to answer the request. If the user asks for drafting help, write the draft directly. "
            "If the answer is not supported by the context, say that clearly.\n\n"
            f"Context:\n{context}\n\nRequest: {query}\nAnswer:"
        )
    else:
        prompt = (
            "You are Quassi, a job search assistant. Answer using only the provided context. "
            "If the answer is not in the context, say that clearly.\n\n"
            f"Context:\n{context}\n\nQuestion: {query}\nAnswer:"
        )
    try:
        from g4f.client import Client
        client = Client()
        response = client.chat.completions.create(
            model=model_name or "gpt-4o",
            messages=[{"role": "user", "content": prompt}],
        )
        generated = response.choices[0].message.content
        if generated:
            generated = generated.strip()
            if len(generated) < 12 or re.fullmatch(r"[\[\]\d\W]+", generated):
                return None
            return generated
    except Exception:
        return None
    return None


def fallback_answer(
    query: str,
    matches: list[tuple[RagDocument, float]],
    *,
    document_name: str = "",
    document_text: str = "",
) -> str:
    if document_text:
        if is_cover_letter_query(query):
            return build_cover_letter_from_document(query, document_text)
        if is_document_summary_query(query, document_text=document_text):
            return build_document_summary(document_name or "your uploaded document", document_text)
        if not is_document_matching_query(query, document_text=document_text):
            return (
                f"I parsed {document_name or 'your uploaded document'}. "
                "Ask me to draft a cover letter, summarize it, or match it against relevant jobs."
            )
    if not matches:
        return "I could not find matching records in your current tables or fetched portal jobs."

    lines = ["Here are the strongest matches I found:"]
    for document, _ in matches[:5]:
        line = f"- {document.title} [{document.source_label}]"
        if document.url:
            line += f" - {document.url}"
        lines.append(line)
    return "\n".join(lines)


def should_include_sources(query: str, *, document_text: str = "") -> bool:
    if is_count_query(query):
        return False
    if smalltalk_answer_for_query(query):
        return False
    if document_text and not is_document_matching_query(query, document_text=document_text):
        return False
    return True


def save_chat_exchange(
    user,
    *,
    query: str,
    answer: str,
    sources: list[dict[str, Any]],
    model_slug: str = "",
    attachment_name: str = "",
    attachment_text: str = "",
) -> None:
    with transaction.atomic():
        ChatMessage.objects.create(
            user=user,
            role=ChatMessage.Role.USER,
            content=query,
            attachment_name=attachment_name,
            attachment_text=attachment_text[:24000],
            metadata={"model_slug": model_slug},
        )
        ChatMessage.objects.create(
            user=user,
            role=ChatMessage.Role.ASSISTANT,
            content=answer,
            metadata={"sources": sources, "model_slug": model_slug},
        )


def add_source_to_to_apply(user, portal_job_id: int) -> str:
    suggestion = PortalJobSuggestion.objects.get(id=portal_job_id, user=user)
    row_key = add_portal_job_to_to_apply(user, suggestion)
    return str(row_key)


def chat_with_rag(
    user,
    query: str,
    *,
    uploaded_document_name: str = "",
    uploaded_document_text: str = "",
    chat_model: ChatModelConfig | None = None,
) -> dict[str, Any]:
    purge_expired_chat_history(user)
    refreshed_portals = refresh_all_portals(user) if should_refresh_portals(query) else 0
    documents = build_user_documents(user)
    search_query = (
        augment_query_with_document(query, uploaded_document_text)
        if is_document_matching_query(query, document_text=uploaded_document_text)
        else query
    )
    matches = score_documents(search_query, documents)
    heuristic_answer = build_heuristic_answer(
        user,
        query,
        matches,
        document_name=uploaded_document_name,
        document_text=uploaded_document_text,
    )
    model_name = normalize_text(chat_model.model_name) if chat_model else os.getenv("RAG_MODEL_NAME", "gpt-4o")
    model_answer = (
        answer_with_ai(
            query,
            matches,
            model_name=model_name,
            document_name=uploaded_document_name,
            document_text=uploaded_document_text,
        )
        if not smalltalk_answer_for_query(query)
        else None
    )
    answer = heuristic_answer or model_answer or fallback_answer(
        query,
        matches,
        document_name=uploaded_document_name,
        document_text=uploaded_document_text,
    )

    source_matches = relevant_matches_for_intent(query, matches, document_text=uploaded_document_text) if should_include_sources(
        query,
        document_text=uploaded_document_text,
    ) else []
    sources = [
        {
            "title": document.title,
            "source_label": document.source_label,
            "source_type": document.source_type,
            "url": document.url,
            "preview": document.text[:280],
            "portal_job_id": document.metadata.get("portal_job_id"),
            "added_to_apply": bool(document.metadata.get("added_to_apply")),
            "application_type": document.metadata.get("values", {}).get("type") or infer_application_type(document.title),
            "can_add_to_apply": document.source_type == "portal-jobs" and not bool(document.metadata.get("added_to_apply")),
        }
        for document, _ in source_matches[:5]
    ]
    save_chat_exchange(
        user,
        query=query,
        answer=answer,
        sources=sources,
        model_slug=normalize_text(chat_model.slug) if chat_model else "",
        attachment_name=uploaded_document_name,
        attachment_text=uploaded_document_text,
    )

    return {
        "answer": answer,
        "assistant_name": "Quassi",
        "used_local_model": model_answer is not None,
        "model_slug": normalize_text(chat_model.slug) if chat_model else "",
        "model_name": normalize_text(chat_model.name) if chat_model else "",
        "refreshed_portals": refreshed_portals,
        "uploaded_document_name": uploaded_document_name,
        "sources": sources,
        "default_prompts": RAG_DEFAULT_PROMPTS,
    }
