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
    "first_name", "last_name", "full_name", "city", "state_code",
    "specialty", "years_experience", "bio",
)

SYSTEM_PROMPT = (
    "You extract structured candidate data from resume text. "
    "Reply with a single JSON object and nothing else. "
    "Use null for anything the resume does not state. Never invent values."
)

USER_PROMPT_TEMPLATE = """Extract this candidate's profile from the resume text below.

Return exactly this JSON shape:
{{
  "first_name": string|null,
  "last_name": string|null,
  "full_name": string|null,
  "city": string|null,
  "state_code": string|null,          // 2-letter US state code, uppercase
  "specialty": string|null,           // primary clinical or professional specialty
  "total_years_experience": number|null,  // TOTAL relevant professional experience across the whole career, not one job
  "positions": [                      // every job/experience entry you can see
    {{"title": string|null, "employer": string|null, "start_year": number|string|null, "end_year": number|string|null}}
  ],
  "bio": string|null                  // 2-4 sentence professional summary in third person
}}

Rules for total_years_experience:
- Prefer a stated total/overall career length (for example "12+ years of nursing experience").
- If only individual jobs are listed, sum their spans without double-counting overlapping dates.
- Never report a single job's duration as the total when the resume shows a longer career.
- Use a plain number of years. Use null if the resume gives no basis for it.

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


def _parse_json_object(raw: str) -> dict:
    text = (raw or "").strip()
    fenced = re.search(r"```(?:json)?\s*(.+?)```", text, re.DOTALL)
    if fenced:
        text = fenced.group(1).strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start == -1 or end <= start:
            raise LlmError(f"Model did not return JSON: {text[:200]}")
        try:
            parsed = json.loads(text[start:end + 1])
        except json.JSONDecodeError as exc:
            raise LlmError(f"Model returned invalid JSON: {text[:200]}") from exc
    if not isinstance(parsed, dict):
        raise LlmError("Model returned JSON that was not an object.")
    return parsed


def extract_profile(text: str) -> dict:
    """Ask Qwen2.5 for the structured profile. Raises LlmError on failure."""
    max_chars = config.env_int("LLM_MAX_INPUT_CHARS", 18000)
    trimmed = text if len(text) <= max_chars else text[:max_chars] + "\n[...truncated...]"
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

    paragraphs = [p.strip() for p in (text or "").split("\n\n") if len(p.strip()) > 120]
    if paragraphs:
        profile["bio"] = paragraphs[0][:1200]

    return profile


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

    return {
        "first_name": _title_case(first),
        "last_name": _title_case(last),
        "full_name": _title_case(full_name),
        "city": _title_case(_clean(raw.get("city")) or fallback.get("city")),
        "state_code": normalize_state(raw.get("state_code")) or fallback.get("state_code"),
        "specialty": _clean(raw.get("specialty")) or _clean(fallback.get("specialty")),
        "bio": bio,
        "_model_total_years": raw.get("total_years_experience"),
        "_model_positions": raw.get("positions") if isinstance(raw.get("positions"), list) else [],
    }
