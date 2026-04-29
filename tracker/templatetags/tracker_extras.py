from django import template

register = template.Library()


@register.filter
def status_chip_classes(value: str) -> str:
    normalized = (value or "").strip().lower()
    if normalized in {"rejected", "withdrawn"}:
        return "bg-rose-100 text-rose-700 ring-1 ring-inset ring-rose-200"
    if normalized in {"offer", "accepted"}:
        return "bg-emerald-100 text-emerald-700 ring-1 ring-inset ring-emerald-200"
    if "interview" in normalized:
        return "bg-amber-100 text-amber-700 ring-1 ring-inset ring-amber-200"
    if normalized in {"screening"}:
        return "bg-violet-100 text-violet-700 ring-1 ring-inset ring-violet-200"
    if normalized in {"assignment"}:
        return "bg-fuchsia-100 text-fuchsia-700 ring-1 ring-inset ring-fuchsia-200"
    if normalized in {"applied", "pending"}:
        return "bg-sky-100 text-sky-700 ring-1 ring-inset ring-sky-200"
    if normalized in {"phd"}:
        return "bg-blue-100 text-blue-700 ring-1 ring-inset ring-blue-200"
    if normalized in {"postdoc"}:
        return "bg-teal-100 text-teal-700 ring-1 ring-inset ring-teal-200"
    if normalized in {"internship", "intern"}:
        return "bg-orange-100 text-orange-700 ring-1 ring-inset ring-orange-200"
    if normalized in {"job"}:
        return "bg-slate-100 text-slate-700 ring-1 ring-inset ring-slate-200"
    if normalized in {"sky", "emerald", "amber", "slate"}:
        return {
            "sky": "bg-sky-100 text-sky-700 ring-1 ring-inset ring-sky-200",
            "emerald": "bg-emerald-100 text-emerald-700 ring-1 ring-inset ring-emerald-200",
            "amber": "bg-amber-100 text-amber-700 ring-1 ring-inset ring-amber-200",
            "slate": "bg-slate-100 text-slate-700 ring-1 ring-inset ring-slate-200",
        }[normalized]
    return "bg-slate-100 text-slate-700 ring-1 ring-inset ring-slate-200"
