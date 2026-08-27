"""Structured profile extraction from PDF text using Qwen2.5.

The model is reached over an OpenAI-compatible /chat/completions endpoint
(Ollama, vLLM, LM Studio, or a hosted gateway). Base URL, API key, and model
name all come from the environment -- nothing is hardcoded.

If the model is unavailable or returns unusable output, a regex fallback fills
in what it can so a run never dies on one bad PDF.
"""
from __future__ import annotations

import json
import os
import re
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from . import config

FIELDS = (
    "first_name", "last_name", "full_name", "email", "phone",
    "city", "state_code", "zip_code", "headline", "specialty",
    "profession_type", "work_authorization", "years_experience", "bio",
)

# List-valued fields that end up in resume_sections / education.
LIST_FIELDS = ("certifications", "skills", "languages", "education", "positions")

SYSTEM_PROMPT = (
    "You are a resume parser. You extract structured candidate data from the "
    "text of one resume, which may have come from OCR and may contain noise.\n"
    "Absolute rules:\n"
    "1. Reply with a single JSON object and nothing else - no prose, no code fence.\n"
    "2. Copy values from the resume. Never invent, guess, infer, or complete a "
    "value that is not written in the text.\n"
    "3. Use null for a missing string or number, and [] for a missing list. "
    "Never use placeholders such as \"N/A\", \"unknown\", \"Not specified\", or \"\".\n"
    "4. Include every job in \"positions\" - do not stop early and do not summarise.\n"
    "5. Drop OCR garbage: stray single letters, icon leftovers, the U+FFFD "
    "character, and lines of unrelated capitals. Remove exact duplicates from lists."
)

USER_PROMPT_TEMPLATE = """Extract this candidate's profile from the resume text below.

Return exactly this JSON shape and these keys:
{{
  "first_name": string|null,
  "last_name": string|null,
  "full_name": string|null,
  "email": string|null,
  "phone": string|null,
  "city": string|null,                // the candidate's own city, not an employer's
  "state_code": string|null,          // 2-letter US state code, uppercase
  "zip_code": string|null,
  "headline": string|null,            // current or most recent job title, verbatim
  "specialty": string|null,           // primary clinical/professional specialty or field
  "profession_type": string|null,     // e.g. "Nurse", "Physician", "Teacher", "Driver"
  "work_authorization": string|null,  // only if the resume states it
  "total_years_experience": number|null,  // TOTAL career experience, not one job
  "bio": string|null,                 // the resume's own summary, or 2-4 factual sentences drawn only from it
  "positions": [                      // EVERY job entry, newest first, no duplicates
    {{
      "title": string|null,
      "employer": string|null,
      "city": string|null,
      "state_code": string|null,
      "start_date": string|null,      // "YYYY-MM" or "YYYY" exactly as datable from the text
      "end_date": string|null,        // same format, or "Present" if current
      "description": string|null      // the duties text for this job, cleaned up
    }}
  ],
  "certifications": [string],         // licences and certifications, deduplicated
  "education": [
    {{"degree": string|null, "field": string|null, "school": string|null, "year": string|null}}
  ],
  "skills": [string],                 // individual skills, deduplicated, no sentences
  "languages": [string]
}}

Rules for total_years_experience:
- Prefer a stated total/overall career length (for example "12+ years of nursing experience").
- If only individual jobs are listed, use the span from the earliest start to the
  latest end without double-counting overlapping dates.
- Never report a single job's duration as the total when the resume shows a longer career.
- Use a plain number of years. Use null if the resume gives no basis for it.

Rules for positions:
- One entry per job heading. A heading usually looks like "Title - Employer".
- Keep the dates as written; do not calculate, shift, or fill in a missing date.
- If the same job appears twice, keep it once.

Rules for lists:
- certifications, skills and languages hold short items, each written once.
- A skill is a short phrase, not a sentence and not a duty description.
- If a section is absent from the resume, return [] for it.

RESUME TEXT:
---
{text}
---
JSON:"""

US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "district of columbia": "DC",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID", "illinois": "IL",
    "indiana": "IN", "iowa": "IA", "kansas": "KS", "kentucky": "KY", "louisiana": "LA",
    "maine": "ME", "maryland": "MD", "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT", "virginia": "VA",
    "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
}
STATE_CODES = set(US_STATES.values())


class LlmError(RuntimeError):
    pass


def llm_enabled() -> bool:
    return config.env_bool("LLM_ENABLED", True)


def model_name() -> str:
    return os.environ.get("LLM_MODEL", "")


def _endpoint() -> str:
    base = (os.environ.get("LLM_BASE_URL") or "").strip().rstrip("/")
    if not base:
        raise LlmError("LLM_BASE_URL is not set.")
    if base.endswith("/chat/completions"):
        return base
    return urljoin(base + "/", "chat/completions")


def _call_model(messages: list[dict], *, timeout: int, temperature: float) -> str:
    payload = {
        "model": model_name(),
        "messages": messages,
        "temperature": temperature,
        "stream": False,
        # Too low and the JSON reply is cut off mid-object, which costs the
        # whole extraction and silently drops the run to regex fallback.
        "max_tokens": config.env_int("LLM_MAX_TOKENS", 1024),
        # Honoured by Ollama/vLLM; harmlessly ignored elsewhere.
        "response_format": {"type": "json_object"},
    }
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    api_key = os.environ.get("LLM_API_KEY")
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    last_error: Exception | None = None
    for attempt in range(1, 4):
        req = Request(_endpoint(), data=data, headers=headers, method="POST")
        try:
            with urlopen(req, timeout=timeout) as resp:
                body = json.loads(resp.read().decode("utf-8", errors="replace") or "{}")
            choices = body.get("choices") or []
            if not choices:
                raise LlmError(f"Model returned no choices: {str(body)[:200]}")
            return (choices[0].get("message") or {}).get("content") or ""
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:300]
            last_error = LlmError(f"HTTP {exc.code} from model endpoint: {detail}")
            if exc.code not in {408, 429, 500, 502, 503, 504} or attempt == 3:
                raise last_error from exc
        except (URLError, TimeoutError) as exc:
            last_error = LlmError(f"Could not reach the model at {_endpoint()}: {exc}")
            if attempt == 3:
                raise last_error from exc
        time.sleep(min(20, 2 ** attempt))

    raise last_error or LlmError("Model call failed.")


def _repair_truncated_json(fragment: str) -> dict | None:
    """Recover a reply that ran out of output tokens mid-object.

    A resume with a dozen jobs produces a long object, and losing the whole
    extraction over the last few characters would drop us to regex for a reply
    that was 95% complete. Trailing partial values are discarded, then the open
    brackets are closed.
    """
    # Walk the text tracking string state, and remember the last position where
    # the structure was safe to cut: just after a completed element.
    depth: list[str] = []
    in_string = False
    escaped = False
    cut = None
    for index, char in enumerate(fragment):
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
                cut = index + 1
            continue
        if char == '"':
            in_string = True
        elif char in "{[":
            depth.append("}" if char == "{" else "]")
        elif char in "}]":
            if depth:
                depth.pop()
            cut = index + 1
        elif char in ",":
            cut = index  # drop the dangling comma with the partial element
        elif char.isdigit() or char in "eE.+-" or char.isalpha():
            cut = index + 1

    if cut is None:
        return None
    candidate = fragment[:cut].rstrip().rstrip(",")
    # Re-derive the still-open brackets for the trimmed text.
    depth = []
    in_string = False
    escaped = False
    for char in candidate:
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char in "{[":
            depth.append("}" if char == "{" else "]")
        elif char in "}]" and depth:
            depth.pop()
    if in_string:
        candidate += '"'
    candidate += "".join(reversed(depth))

    try:
        parsed = json.loads(candidate)
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def _parse_json_object(raw: str) -> dict:
    text = (raw or "").strip()
    fenced = re.search(r"```(?:json)?\s*(.+?)```", text, re.DOTALL)
    if fenced:
        text = fenced.group(1).strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end > start:
            try:
                parsed = json.loads(text[start:end + 1])
            except json.JSONDecodeError:
                parsed = _repair_truncated_json(text[start:])
        elif start != -1:
            # No closing brace at all: the reply hit the token limit.
            parsed = _repair_truncated_json(text[start:])
        else:
            raise LlmError(f"Model did not return JSON: {text[:200]}")
        if parsed is None:
            raise LlmError(f"Model returned invalid JSON: {text[:200]}")
    if not isinstance(parsed, dict):
        raise LlmError("Model returned JSON that was not an object.")
    return parsed


def extract_profile(text: str) -> dict:
    """Ask Qwen2.5 for the structured profile. Raises LlmError on failure."""
    max_chars = config.env_int("LLM_MAX_INPUT_CHARS", 18000)
    cleaned = scrub_ocr_noise(text)
    trimmed = cleaned if len(cleaned) <= max_chars else cleaned[:max_chars] + "\n[...truncated...]"
    content = _call_model(
        [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": USER_PROMPT_TEMPLATE.format(text=trimmed)},
        ],
        timeout=config.env_int("LLM_TIMEOUT_SECONDS", 180),
        temperature=config.env_float("LLM_TEMPERATURE", 0.0),
    )
    return _parse_json_object(content)


# ---------------------------------------------------------------------------
# OCR clean-up
#
# Tesseract turns the icon glyphs a resume uses for bullets into replacement
# characters and short runs of unrelated capitals ("Dn Hb HB BD YH DW DW").
# Left in, the model copies them into names and skills.

_OCR_ICON_CHARS = "�■●▪•"

# Symbol characters an icon glyph decodes to. Letters and ordinary punctuation
# are left alone so real content is never touched.
_ICON_SYMBOLS = re.compile(r"[^\w\s,.:;'\"()/&+#%$@!?\[\]{}=*<>|~`^\\-]")

# A line that is only short letter tokens is an icon run, not content
# ("Dn Hb HB BD YH DW DW" is what a column of bullet icons OCRs to).
_ICON_RUN = re.compile(r"^(?:[A-Za-z]{1,3}[ ,.]+){2,}[A-Za-z]{1,3}[.]?$")

# Leading punctuation left where an icon used to be ("= five towns college").
_LEADING_PUNCT = re.compile(r"^[\W_]+")

# A resume's bullet icons OCR as a stray 1-2 letter token before the real text
# ("Fy Infant room teacher"). Genuine short leading words must survive.
_REAL_SHORT_WORDS = {
    "i", "a", "an", "at", "as", "in", "on", "of", "to", "up", "by", "or", "is", "it",
    "dr", "mr", "ms", "mx", "st", "us", "rn", "md", "do", "np", "pa", "bs", "ba", "ms",
}
_LEADING_ICON_TOKEN = re.compile(r"^([A-Za-z]{1,2})\s+(?=[A-Za-z].*\s)")


def _strip_leading_icon(line: str) -> str:
    match = _LEADING_ICON_TOKEN.match(line)
    if match and match.group(1).lower() not in _REAL_SHORT_WORDS:
        return line[match.end():]
    return line


def scrub_ocr_noise(text: str) -> str:
    """Drop icon glyphs and icon-run lines so the model sees only real content."""
    lines: list[str] = []
    for raw in (text or "").split("\n"):
        line = raw
        for char in _OCR_ICON_CHARS:
            line = line.replace(char, "")
        line = line.replace("’", "'").replace("‘", "'")
        line = line.replace("“", '"').replace("”", '"')
        # Tesseract routinely reads a sentence-initial "I" as "|" or "l".
        line = re.sub(r"^[|]\s+(?=[a-z])", "I ", line)
        # Runs of spaces are kept (squashed to exactly two) because they are the
        # only surviving trace of a column break in an OCR'd skills grid.
        line = _ICON_SYMBOLS.sub(" ", line)
        line = re.sub(r"\t", " ", line)
        line = re.sub(r" {2,}", "  ", line).strip()
        if not line:
            lines.append("")
            continue
        if _ICON_RUN.match(line):
            continue
        line = _strip_leading_icon(_LEADING_PUNCT.sub("", line)).strip()
        if line:
            lines.append(line)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


# ---------------------------------------------------------------------------
# Normalisation and regex fallback


def _clean(value) -> str | None:
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value)).strip().strip(",;|")
    if not text or text.lower() in {"null", "none", "n/a", "na", "unknown", "-"}:
        return None
    return text


def normalize_state(value) -> str | None:
    text = _clean(value)
    if not text:
        return None
    upper = text.upper()
    if upper in STATE_CODES:
        return upper
    return US_STATES.get(text.lower())


def _title_case(value: str | None) -> str | None:
    if not value:
        return None
    return " ".join(part.capitalize() if part.islower() or part.isupper() else part
                    for part in value.split())


def split_full_name(full_name: str | None) -> tuple[str | None, str | None]:
    text = _clean(full_name)
    if not text:
        return None, None
    text = re.split(r"[,(]", text)[0].strip()
    parts = [p.strip(".,") for p in text.split() if p.strip(".,")]
    if not parts:
        return None, None
    if len(parts) == 1:
        return _title_case(parts[0]), None
    return _title_case(parts[0]), _title_case(parts[-1])


_NAME_LINE = re.compile(r"^[A-Z][A-Za-z'\-]+(?: [A-Z][A-Za-z'\-.]+){1,3}$")
_CITY_STATE = re.compile(r"\b([A-Z][A-Za-z .'\-]{2,30}),\s*([A-Z]{2})\b")

# Post-nominals that follow a name. Several collide with state codes (MD, PA,
# OR, IN...), so a "City, ST" match has to be checked against them.
CREDENTIALS = {
    "RN", "BSN", "MSN", "DNP", "APRN", "NP", "FNP", "PNP", "AGNP", "CRNA", "CNM",
    "LPN", "LVN", "CNA", "MD", "DO", "PA", "PAC", "PT", "OT", "RT", "MBA", "MPH",
    "PHD", "MS", "BS", "BA", "CCRN", "ACLS", "BLS", "FACS", "FACP", "ESQ", "JR", "SR",
}
_SECTION_HEADINGS = {
    "professional summary", "summary", "profile", "objective", "experience",
    "work experience", "professional experience", "education", "skills",
    "certifications", "licenses", "employment history", "contact", "references",
}


def _strip_credentials(line: str) -> str:
    """'Maria Gonzalez, RN, BSN' -> 'Maria Gonzalez'."""
    head = line.split(",")[0].strip()
    tokens = [t for t in head.split() if t.strip(".,").upper() not in CREDENTIALS]
    return " ".join(tokens).strip()


def regex_profile(text: str) -> dict:
    """Best-effort structured fields without the model."""
    profile: dict = {}
    # Same clean-up the model sees, so both paths agree on section boundaries.
    text = scrub_ocr_noise(text)
    lines = [line.strip() for line in (text or "").split("\n")]

    name_line_index = None
    for index, stripped in enumerate(lines[:12]):
        if stripped.lower() in _SECTION_HEADINGS:
            continue
        # An all-caps line is a heading, not a name.
        if stripped.isupper():
            continue
        candidate = _strip_credentials(stripped)
        if 4 <= len(candidate) <= 60 and _NAME_LINE.match(candidate):
            profile["full_name"] = candidate
            name_line_index = index
            break

    for index, line in enumerate(lines):
        if index == name_line_index:
            continue
        match = _CITY_STATE.search(line)
        if not match:
            continue
        city, state = match.group(1).strip(), match.group(2)
        if state not in STATE_CODES:
            continue
        # "Gonzalez, RN, BSN" is a credential list, not a city and state.
        if state in CREDENTIALS and any(
            token.strip(".,").upper() in CREDENTIALS for token in line.split(",")[2:3] + line.split()[-2:]
        ):
            continue
        if len(city.split()) > 3:
            continue
        profile["city"] = _title_case(city)
        profile["state_code"] = state
        break

    specialty = re.search(
        r"(?:specialty|specialt(?:y|ies)|department|practice area)\s*[:\-]\s*([^\n]{2,60})",
        text or "", re.IGNORECASE,
    )
    if specialty:
        profile["specialty"] = _clean(specialty.group(1))

    # The summary section only. The first long paragraph of a resume is usually
    # the newest job's duties, which is not a bio.
    summary = " ".join(_section_lines(text, (
        "professional summary", "summary", "profile", "objective", "about",
        "career summary", "overview",
    )))
    summary = _clean(summary)
    if summary and len(summary) > 20:
        profile["bio"] = summary[:1200]

    profile.update(_regex_contact(text))
    profile.update(_regex_sections(text))
    profile["positions"] = _regex_positions(text)
    if profile.get("positions"):
        profile["headline"] = profile["positions"][0].get("title")
    return profile


# --- regex extraction of the remaining schema fields -----------------------

_EMAIL = re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]{2,}\b")
# US phone: optional +1, then 3-3-4 with any of space, dot, dash or parens.
_PHONE = re.compile(r"(?<!\d)(?:\+?1[\s.-]?)?\(?([2-9]\d{2})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})(?!\d)")
_ZIP = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
_WORK_AUTH = re.compile(
    r"^.{0,40}\b(authorized to work[^\n.]{0,60}|us citizen|green card[^\n.]{0,30}"
    r"|permanent resident|h-?1b[^\n.]{0,20}|opt\b[^\n.]{0,20}|ead\b[^\n.]{0,20})",
    re.IGNORECASE | re.MULTILINE,
)


def _regex_contact(text: str) -> dict:
    found: dict = {}
    email = _EMAIL.search(text or "")
    if email:
        found["email"] = email.group(0).lower()
    phone = _PHONE.search(text or "")
    if phone:
        found["phone"] = f"({phone.group(1)}) {phone.group(2)}-{phone.group(3)}"
    # Only trust a ZIP that sits right after a state code, so years and counts
    # elsewhere in the resume are never mistaken for one.
    zip_match = re.search(r"\b[A-Z]{2}\s+" + _ZIP.pattern, text or "")
    if zip_match:
        found["zip_code"] = zip_match.group(1)
    auth = _WORK_AUTH.search(text or "")
    if auth:
        found["work_authorization"] = _clean(auth.group(1))
    return found


# Section headings as they appear in a resume, mapped to our field names.
_SECTION_MAP = (
    ("certifications", (
        "certifications & licenses", "certifications and licenses", "certifications",
        "licenses & certifications", "licenses and certifications", "licenses",
        "licensure", "credentials",
    )),
    ("skills", ("skills", "technical skills", "core competencies", "competencies")),
    ("languages", ("languages", "language proficiency")),
    ("education", ("education", "education & training", "academic background")),
)
_ALL_HEADINGS = tuple(
    heading for _field, headings in _SECTION_MAP for heading in headings
) + (
    "experience", "work experience", "professional experience", "employment history",
    "work history", "professional summary", "summary", "additional information",
    "objective", "profile", "references", "awards", "publications", "contact",
)

# Boilerplate the source job board prints under each section heading.
_BOILERPLATE = re.compile(
    r"^related .{0,60} (?:come|comes) from the candidate.{0,30}$", re.IGNORECASE
)


def _section_lines(text: str, headings: tuple[str, ...], *, keep_blanks: bool = False) -> list[str]:
    """Lines between the given heading and the next heading for a *different* field.

    A resume often nests a heading of the same family inside its own section
    ("Certifications & licenses" ... "Certifications"). Stopping at those would
    truncate the section to its first entry, so only a foreign heading ends it.
    """
    lines = [line.strip() for line in (text or "").split("\n")]
    start = None
    for index, line in enumerate(lines):
        if line.lower().strip(":") in headings:
            start = index + 1
            break
    if start is None:
        return []
    collected: list[str] = []
    for line in lines[start:]:
        key = line.lower().strip(":")
        if key in headings:
            continue  # a nested heading of the same family
        if key in _ALL_HEADINGS:
            break
        if _BOILERPLATE.match(line):
            continue
        if line or keep_blanks:
            collected.append(line)
    return collected


def _dedupe(items) -> list[str]:
    """Case-insensitive de-duplication that keeps the first spelling seen."""
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        value = _clean(item)
        if not value:
            continue
        key = re.sub(r"[^a-z0-9]+", "", value.lower())
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


# A resume's skills grid OCRs as one line per row, with the column break shown
# only by a word-boundary: "Toddler Care Infant Care English Preschool experience".
# Splitting where a lowercase word is followed by a capitalised one recovers the
# items. It is a heuristic for the no-model fallback; the model does this better.
_SKILL_COLUMN_BREAK = re.compile(r"(?<=[a-z])\s+(?=[A-Z])")


def _split_skill_line(line: str) -> list[str]:
    """A skills line is comma-, bullet-, wide-space- or column-separated."""
    parts: list[str] = []
    for chunk in re.split(r"\s*[,;|•·]\s*|\s{2,}|\.\s+", line):
        chunk = chunk.strip(" .")
        if not chunk:
            continue
        # Only worth splitting a long fused row; short entries stay intact.
        parts.extend(_SKILL_COLUMN_BREAK.split(chunk) if len(chunk) > 30 else [chunk])
    return [part.strip(" .") for part in parts if part.strip(" .")]


def _regex_sections(text: str) -> dict:
    found: dict = {}

    certifications = _dedupe(
        line for line in _section_lines(text, _SECTION_MAP[0][1])
        if 2 < len(line) <= 80 and line.lower() != "certifications"
    )
    if certifications:
        found["certifications"] = certifications

    skills: list[str] = []
    for line in _section_lines(text, _SECTION_MAP[1][1]):
        # A long sentence in the skills section is a duty, not a skill.
        if len(line) > 200:
            continue
        skills.extend(part for part in _split_skill_line(line) if 1 < len(part) <= 60)
    skills = _dedupe(skills)
    if skills:
        found["skills"] = skills

    languages = _dedupe(
        part for line in _section_lines(text, _SECTION_MAP[2][1])
        for part in _split_skill_line(line) if 1 < len(part) <= 40
    )
    if languages:
        found["languages"] = languages

    education_lines = _dedupe(
        line for line in _section_lines(text, _SECTION_MAP[3][1]) if 2 < len(line) <= 120
    )
    if education_lines:
        found["education"] = [_education_entry(line) for line in education_lines]

    return found


_DEGREE_WORDS = (
    "associate", "bachelor", "master", "doctor", "phd", "ph.d", "md", "bsn", "msn",
    "dnp", "adn", "rn", "diploma", "certificate", "ged", "high school", "some college",
    "b.s", "b.a", "m.s", "m.a", "mba",
)


def _education_entry(line: str) -> dict:
    """Turn one education line into the {degree, field, school, year} shape."""
    year = re.search(r"\b(19|20)\d{2}\b", line)
    lowered = line.lower()
    is_degree = any(word in lowered for word in _DEGREE_WORDS)
    return {
        "degree": _clean(line) if is_degree else None,
        "field": None,
        "school": None if is_degree else _clean(line),
        "year": year.group(0) if year else None,
    }


# "Title - Employer" heading, then an optional date line, then optional "City, ST".
# Commas are excluded from the title: a wrapped duty sentence that happens to
# contain " - " ("...spectrum , non - verbal...") otherwise parses as a new job.
_POSITION_HEADING = re.compile(r"^(?P<title>[^\n,;.]{2,60}?)\s+[-–—]\s+(?P<employer>[^\n;]{2,70})$")
_DATE_LINE = re.compile(
    r"^(?:(?P<start>(?:[A-Z][a-z]+\s+)?(?:19|20)\d{2}|Through\s+(?:19|20)\d{2})"
    r"\s*[-–—]+\s*(?P<end>Present|(?:[A-Z][a-z]+\s+)?(?:19|20)\d{2})"
    r"|(?P<only>Through\s+(?:19|20)\d{2}|(?:19|20)\d{2}))\b",
    re.IGNORECASE,
)


def _regex_positions(text: str) -> list[dict]:
    """Parse the experience section into position entries without the model."""
    lines = _section_lines(text, (
        "experience", "work experience", "professional experience",
        "employment history", "work history",
    ), keep_blanks=True)
    positions: list[dict] = []
    current: dict | None = None
    at_block_start = True
    for line in lines:
        if not line:
            at_block_start = True
            continue
        starts_block, at_block_start = at_block_start, False
        heading = _POSITION_HEADING.match(line)
        # A date line can also contain a dash, so check dates first.
        date_match = _DATE_LINE.match(line)
        # A job heading always opens a block; mid-paragraph it is a wrapped line.
        if heading and not date_match and starts_block:
            current = {
                "title": _clean(heading.group("title")),
                "employer": _clean(heading.group("employer")),
                "city": None, "state_code": None,
                "start_date": None, "end_date": None, "description": None,
            }
            positions.append(current)
            continue
        if current is None:
            continue
        if date_match:
            current["start_date"] = _clean(date_match.group("start") or date_match.group("only"))
            current["end_date"] = _clean(date_match.group("end"))
            continue
        city_state = _CITY_STATE.match(line)
        if city_state and city_state.group(2) in STATE_CODES and not current["city"]:
            current["city"] = _title_case(city_state.group(1).strip())
            current["state_code"] = city_state.group(2)
            continue
        current["description"] = f"{current['description']} {line}".strip() if current["description"] else line
    return _dedupe_positions(positions)


def _position_key(position: dict) -> str:
    return re.sub(
        r"[^a-z0-9]+", "",
        f"{position.get('title') or ''}|{position.get('employer') or ''}".lower(),
    )


def _dedupe_positions(positions) -> list[dict]:
    """One entry per title+employer; the richer duplicate wins."""
    merged: dict[str, dict] = {}
    for position in positions:
        if not isinstance(position, dict):
            continue
        entry = {
            "title": _clean(position.get("title")),
            "employer": _clean(position.get("employer")),
            "city": _title_case(_clean(position.get("city"))),
            "state_code": normalize_state(position.get("state_code")),
            "start_date": _clean(position.get("start_date") or position.get("start_year")),
            "end_date": _clean(position.get("end_date") or position.get("end_year")),
            "description": _clean(position.get("description")),
        }
        if not entry["title"] and not entry["employer"]:
            continue
        key = _position_key(entry)
        existing = merged.get(key)
        if existing is None:
            merged[key] = entry
            continue
        for field, value in entry.items():
            if value and not existing.get(field):
                existing[field] = value
    return list(merged.values())


def _merge_lists(model_value, fallback_value, *, limit: int) -> list[str]:
    """Model list first, regex list second, deduplicated and capped."""
    model_items = model_value if isinstance(model_value, list) else []
    fallback_items = fallback_value if isinstance(fallback_value, list) else []
    return _dedupe(list(model_items) + list(fallback_items))[:limit]


def _normalize_education(model_value, fallback_value) -> list[dict]:
    entries = model_value if isinstance(model_value, list) else []
    cleaned: list[dict] = []
    seen: set[str] = set()
    for entry in entries:
        if isinstance(entry, str):
            entry = _education_entry(entry)
        if not isinstance(entry, dict):
            continue
        item = {
            "degree": _clean(entry.get("degree")),
            "field": _clean(entry.get("field")),
            "school": _clean(entry.get("school")),
            "year": _clean(entry.get("year")),
        }
        if not any(item.values()):
            continue
        key = re.sub(r"[^a-z0-9]+", "", "".join(str(v or "") for v in item.values()).lower())
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(item)
    if cleaned:
        return cleaned
    return fallback_value if isinstance(fallback_value, list) else []


def normalize_profile(raw: dict, text: str) -> dict:
    """Merge model output with regex fallbacks into the canonical field set."""
    raw = raw or {}
    fallback = regex_profile(text)

    full_name = _clean(raw.get("full_name")) or _clean(fallback.get("full_name"))
    first = _clean(raw.get("first_name"))
    last = _clean(raw.get("last_name"))
    if not first or not last:
        derived_first, derived_last = split_full_name(full_name)
        first = first or derived_first
        last = last or derived_last
    if not full_name and (first or last):
        full_name = " ".join(part for part in (first, last) if part)

    bio = _clean(raw.get("bio")) or _clean(fallback.get("bio"))
    if bio and len(bio) > 4000:
        bio = bio[:4000].rsplit(" ", 1)[0]

    # Positions from the model win, but the regex pass fills in a missing list
    # and its entries top up any field the model left blank.
    model_positions = raw.get("positions") if isinstance(raw.get("positions"), list) else []
    positions = _dedupe_positions(model_positions) or []
    fallback_positions = fallback.get("positions") or []
    if not positions:
        positions = fallback_positions
    else:
        by_key = {_position_key(p): p for p in positions}
        for candidate in fallback_positions:
            existing = by_key.get(_position_key(candidate))
            if existing is None:
                positions.append(candidate)
                continue
            for field, value in candidate.items():
                if value and not existing.get(field):
                    existing[field] = value

    headline = (
        _clean(raw.get("headline"))
        or (positions[0].get("title") if positions else None)
        or _clean(fallback.get("headline"))
    )

    email = _clean(raw.get("email")) or fallback.get("email")
    if email and not _EMAIL.fullmatch(email):
        email = fallback.get("email")

    return {
        "first_name": _title_case(first),
        "last_name": _title_case(last),
        "full_name": _title_case(full_name),
        "email": (email or "").lower() or None,
        "phone": _clean(raw.get("phone")) or fallback.get("phone"),
        "city": _title_case(_clean(raw.get("city")) or fallback.get("city")),
        "state_code": normalize_state(raw.get("state_code")) or fallback.get("state_code"),
        "zip_code": _clean(raw.get("zip_code")) or fallback.get("zip_code"),
        "headline": headline,
        "specialty": _clean(raw.get("specialty")) or _clean(fallback.get("specialty")),
        "profession_type": _clean(raw.get("profession_type")),
        "work_authorization": (
            _clean(raw.get("work_authorization")) or fallback.get("work_authorization")
        ),
        "bio": bio,
        "certifications": _merge_lists(
            raw.get("certifications"), fallback.get("certifications"), limit=100),
        "skills": _merge_lists(raw.get("skills"), fallback.get("skills"), limit=200),
        "languages": _merge_lists(raw.get("languages"), fallback.get("languages"), limit=30),
        "education": _normalize_education(raw.get("education"), fallback.get("education")),
        "positions": positions,
        "_model_total_years": raw.get("total_years_experience"),
        "_model_positions": positions,
    }
