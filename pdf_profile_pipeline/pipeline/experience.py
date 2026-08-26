"""Decide a single years_experience value for a profile.

A resume PDF usually mentions experience in several ways at once: a summary
line ("12+ years of experience"), a per-role duration ("3 years" at one clinic),
and a list of employment date ranges. Taking the first match is wrong -- it
often picks up one job's length. This module gathers every candidate, labels it
by kind, and applies a priority order:

    1. an explicit TOTAL/OVERALL statement       ("15 years of total experience")
    2. a summary-section experience statement    (top-of-resume "12+ years ...")
    3. the model's own total_years_experience
    4. the span covered by the employment date ranges, de-overlapped
    5. the largest single stated duration        (last resort)

Ties inside a kind resolve to the largest value, since a resume that says both
"5 years" and "10 years" of total experience usually means the longer career.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

MAX_REASONABLE_YEARS = 60

# Rank order; lower number wins.
KIND_PRIORITY = {
    "explicit_total": 1,
    "summary_statement": 2,
    "model_total": 3,
    "date_range_span": 4,
    "single_duration": 5,
}

# Deliberately narrow: a phrase like "years of experience in cardiology" is NOT
# a total claim, so generic words such as "experience in" are not hints here.
TOTAL_HINTS = (
    "total", "overall", "combined", "cumulative", "collectively",
    "altogether", "career", "in aggregate", "across my career",
)

SUMMARY_HEADINGS = (
    "summary", "professional summary", "profile", "objective", "about",
    "overview", "career summary", "professional profile", "highlights",
)

# Where the per-job section starts. A "3 years" inside here is one job's
# duration, never the career total, so the summary region must stop before it.
EXPERIENCE_HEADINGS = (
    "experience", "work experience", "professional experience", "employment",
    "employment history", "work history", "clinical experience", "career history",
    "positions", "education",
)

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}

# (?<![\d.]) keeps "400 years" from being read as "00 years".
_YEARS_PATTERN = re.compile(
    r"(?<![\d.])(?P<value>\d{1,2}(?:\.\d)?)\s*\+?\s*(?:-|to|–)?\s*(?:\d{1,2})?\s*"
    r"(?:yrs?|years?)\b",
    re.IGNORECASE,
)
_WORD_YEARS_PATTERN = re.compile(
    r"\b(?P<word>one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|"
    r"fifteen|twenty|twenty-five|thirty)\s+(?:\+\s*)?(?:yrs?|years?)\b",
    re.IGNORECASE,
)
_WORD_VALUES = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
    "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "fifteen": 15, "twenty": 20, "twenty-five": 25, "thirty": 30,
}

_MONTH_YEAR = r"(?:(?P<m{i}>[A-Za-z]{{3,9}})\.?\s+)?(?P<y{i}>(?:19|20)\d{{2}})"
_DATE_RANGE_PATTERN = re.compile(
    _MONTH_YEAR.format(i=1)
    + r"\s*(?:-|to|through|until|–|—)\s*"
    + r"(?:(?P<present>present|current|now|to\s*date|ongoing)|"
    + _MONTH_YEAR.format(i=2) + r")",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ExperienceCandidate:
    years: float
    kind: str
    evidence: str

    @property
    def priority(self) -> int:
        return KIND_PRIORITY.get(self.kind, 99)


def _plausible(years: float | int | None) -> bool:
    return years is not None and 0 < float(years) <= MAX_REASONABLE_YEARS


def _first_heading_index(text: str, headings: tuple[str, ...]) -> int | None:
    # The leading "\n" lets a heading on the very first line match too.
    lowered = "\n" + text.lower()
    best = None
    for heading in headings:
        for probe in (f"\n{heading}\n", f"\n{heading} ", f"\n{heading}:"):
            index = lowered.find(probe)
            if index != -1 and (best is None or index < best):
                best = index
    return None if best is None else max(0, best - 1)


def _summary_bounds(text: str) -> tuple[int, int]:
    """Character range of the summary/profile section, as (start, end).

    The region always stops at the experience section. Without that cut, a
    resume that opens straight into job entries would have a single job's
    "3 years" misread as a summary-level career figure.
    """
    experience_start = _first_heading_index(text, EXPERIENCE_HEADINGS)
    summary_start = _first_heading_index(text, SUMMARY_HEADINGS)

    if summary_start is None:
        # No summary heading: the opening of the resume acts as the summary,
        # but only up to where the per-job section begins.
        end = min(len(text), 900)
        if experience_start is not None:
            end = min(end, experience_start)
        return 0, max(0, end)

    end = min(len(text), summary_start + 1200)
    if experience_start is not None and experience_start > summary_start:
        end = min(end, experience_start)
    return summary_start, end


def _statement_around(text: str, position: int) -> str:
    """The sentence/line containing `position`.

    Scoped this tightly on purpose: a wide character window lets a "total"
    from one sentence leak into the classification of the next one.
    """
    start = max(
        text.rfind("\n", 0, position),
        text.rfind(". ", 0, position),
        text.rfind(";", 0, position),
    )
    start = 0 if start == -1 else start + 1
    candidates = [index for index in (
        text.find("\n", position),
        text.find(". ", position),
        text.find(";", position),
    ) if index != -1]
    end = min(candidates) if candidates else len(text)
    return text[start:end]


def _mentions_total(fragment: str) -> bool:
    lowered = fragment.lower()
    return any(hint in lowered for hint in TOTAL_HINTS)


def scan_text_candidates(text: str) -> list[ExperienceCandidate]:
    """Every 'N years' phrase in the PDF text, labelled by how it was stated."""
    if not text:
        return []

    summary_start, summary_end = _summary_bounds(text)
    candidates: list[ExperienceCandidate] = []

    def _add(value: float, start: int, evidence: str) -> None:
        if not _plausible(value):
            return
        statement = _statement_around(text, start)
        if _mentions_total(statement):
            kind = "explicit_total"
        elif summary_start <= start < summary_end:
            kind = "summary_statement"
        else:
            kind = "single_duration"
        candidates.append(ExperienceCandidate(float(value), kind, " ".join(statement.split())[:160]))

    for match in _YEARS_PATTERN.finditer(text):
        try:
            value = float(match.group("value"))
        except (TypeError, ValueError):
            continue
        _add(value, match.start(), match.group(0))

    for match in _WORD_YEARS_PATTERN.finditer(text):
        value = _WORD_VALUES.get(match.group("word").lower())
        if value:
            _add(float(value), match.start(), match.group(0))

    return candidates


def _month_index(name: str | None) -> int:
    if not name:
        return 1
    return MONTHS.get(name[:4].lower().rstrip("."), MONTHS.get(name[:3].lower(), 1))


def parse_date_ranges(text: str, *, today: date | None = None) -> list[tuple[float, float]]:
    """Employment date ranges as (start, end) decimal years."""
    today = today or date.today()
    now_decimal = today.year + (today.month - 1) / 12
    ranges: list[tuple[float, float]] = []

    for match in _DATE_RANGE_PATTERN.finditer(text or ""):
        start_year = int(match.group("y1"))
        start = start_year + (_month_index(match.group("m1")) - 1) / 12
        if match.group("present"):
            end = now_decimal
        else:
            end_year_raw = match.group("y2")
            if not end_year_raw:
                continue
            end = int(end_year_raw) + (_month_index(match.group("m2")) - 1) / 12
        if end < start:
            start, end = end, start
        if start < 1950 or end > now_decimal + 1:
            continue
        ranges.append((start, min(end, now_decimal)))

    return ranges


def merged_span_years(ranges: list[tuple[float, float]]) -> float:
    """Total years covered by the ranges, counting overlapping jobs once."""
    if not ranges:
        return 0.0
    merged: list[list[float]] = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return round(sum(end - start for start, end in merged), 2)


def collect_candidates(
    text: str,
    *,
    model_total: float | int | None = None,
    model_positions: list[dict] | None = None,
    today: date | None = None,
) -> list[ExperienceCandidate]:
    candidates = scan_text_candidates(text)

    if _plausible(model_total):
        candidates.append(ExperienceCandidate(
            float(model_total), "model_total", "model total_years_experience",
        ))

    ranges = parse_date_ranges(text, today=today)
    ranges += _ranges_from_model_positions(model_positions or [], today=today)
    span = merged_span_years(ranges)
    if _plausible(span):
        candidates.append(ExperienceCandidate(
            span, "date_range_span", f"{len(ranges)} employment date range(s)",
        ))

    return candidates


def _ranges_from_model_positions(positions: list[dict], *, today: date | None) -> list[tuple[float, float]]:
    today = today or date.today()
    now_decimal = today.year + (today.month - 1) / 12
    ranges: list[tuple[float, float]] = []
    for position in positions:
        if not isinstance(position, dict):
            continue
        start = _year_value(position.get("start_year") or position.get("start_date"))
        raw_end = position.get("end_year") or position.get("end_date")
        if start is None:
            continue
        if isinstance(raw_end, str) and raw_end.strip().lower() in {"present", "current", "now", ""}:
            end = now_decimal
        else:
            end = _year_value(raw_end)
        if end is None:
            end = now_decimal
        if end < start:
            start, end = end, start
        if start >= 1950 and end <= now_decimal + 1:
            ranges.append((start, min(end, now_decimal)))
    return ranges


def _year_value(raw) -> float | None:
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        year = float(raw)
        return year if 1950 <= year <= 2100 else None
    match = re.search(r"(19|20)\d{2}", str(raw))
    return float(match.group(0)) if match else None


def resolve_years_experience(
    text: str,
    *,
    model_total: float | int | None = None,
    model_positions: list[dict] | None = None,
    today: date | None = None,
) -> tuple[int, dict]:
    """Return (years, decision detail) using the documented priority order."""
    candidates = collect_candidates(
        text, model_total=model_total, model_positions=model_positions, today=today,
    )
    detail = {
        "candidates": [
            {"years": c.years, "kind": c.kind, "evidence": c.evidence} for c in candidates
        ],
    }
    if not candidates:
        detail.update({"chosen_kind": None, "chosen_years": 0, "reason": "no experience signal found"})
        return 0, detail

    best = sorted(candidates, key=lambda c: (c.priority, -c.years))[0]
    years = int(round(best.years))
    years = max(0, min(years, MAX_REASONABLE_YEARS))
    detail.update({
        "chosen_kind": best.kind,
        "chosen_years": years,
        "chosen_evidence": best.evidence,
        "reason": f"highest-priority signal was {best.kind}",
    })
    return years, detail
