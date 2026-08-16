from __future__ import annotations

import argparse
import hashlib
import importlib
import io
import json
import mimetypes
import os
import re
import subprocess
import sys
import uuid
from datetime import date, datetime, timezone
from getpass import getpass
from pathlib import Path
from time import monotonic
from urllib.parse import urlparse


SCRIPT_DIR = Path(__file__).resolve().parent
SUPPORTED = {".pdf", ".docx"}
DEFAULT_PREFIX = "resumes/import"
DEFAULT_MANIFEST = "data_upload_manifest.jsonl"

PIP_PACKAGES = {
    "sqlalchemy": "sqlalchemy",
    "psycopg": "psycopg[binary]",
    "boto3": "boto3",
    "pypdf": "pypdf",
    "docx": "python-docx",
}

REQUIRED_UPLOAD_ENV = (
    ("DATABASE_URL", "Neon DATABASE_URL", True),
    ("S3_ENDPOINT_URL", "Cloudflare R2 endpoint URL", False),
    ("S3_BUCKET", "Cloudflare R2 bucket name", False),
    ("S3_ACCESS_KEY", "Cloudflare R2 access key", True),
    ("S3_SECRET_KEY", "Cloudflare R2 secret key", True),
)

DEFAULT_UPLOAD_ENV = {
    "STORAGE_ENABLED": "true",
    "S3_REGION": "auto",
    "S3_PUBLIC": "false",
    "S3_ACL": "",
}


# ---------------------------------------------------------------------------
# Environment and dependency helpers

def _import_or_exit(module_name: str):
    try:
        return importlib.import_module(module_name)
    except ModuleNotFoundError as exc:
        package = PIP_PACKAGES.get(module_name, module_name)
        raise SystemExit(
            f"ERROR: Missing Python package '{package}'. Install dependencies with:\n"
            f"  {sys.executable} -m pip install sqlalchemy \"psycopg[binary]\" boto3 pypdf python-docx"
        ) from exc


def _install_missing_dependencies() -> None:
    missing = []
    for module_name, package in PIP_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ModuleNotFoundError:
            missing.append(package)
    if not missing:
        print("All required Python packages are already installed.")
        return
    cmd = [sys.executable, "-m", "pip", "install", *missing]
    print("Installing missing packages:", " ".join(missing))
    subprocess.check_call(cmd)


def _env_path(path: str) -> Path:
    candidate = Path(path).expanduser()
    return candidate if candidate.is_absolute() else SCRIPT_DIR / candidate


def _load_env(path: str = ".env") -> None:
    env_path = _env_path(path)
    if not env_path.exists():
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, value)


def _prompt_env_value(key: str, label: str, secret: bool) -> str:
    prompt = f"{label} ({key}): "
    value = getpass(prompt) if secret else input(prompt)
    value = value.strip().strip('"').strip("'")
    if not value:
        raise SystemExit(f"ERROR: {key} is required.")
    return value


def _append_env_values(env_path: Path, values: dict[str, str]) -> None:
    if not values:
        return
    env_path.parent.mkdir(parents=True, exist_ok=True)
    needs_newline = env_path.exists() and env_path.stat().st_size > 0
    with open(env_path, "a", encoding="utf-8") as f:
        if needs_newline:
            f.write("\n")
        f.write("# Added by data_upload.py first-run setup\n")
        for key, value in values.items():
            f.write(f"{key}={value}\n")


def _prepare_runtime_env(args: argparse.Namespace) -> None:
    env_file = getattr(args, "env_file", ".env")
    _load_env(env_file)
    if getattr(args, "dry_run", False):
        return

    env_updates: dict[str, str] = {}
    for key, value in DEFAULT_UPLOAD_ENV.items():
        if key not in os.environ:
            os.environ[key] = value
            env_updates[key] = value

    missing = [(k, label, secret) for k, label, secret in REQUIRED_UPLOAD_ENV if not os.environ.get(k)]
    if missing:
        print("Enter upload credentials. Secrets are hidden while typing.")
    for key, label, secret in missing:
        value = _prompt_env_value(key, label, secret)
        os.environ[key] = value
        env_updates[key] = value

    if env_updates and not getattr(args, "no_save_credentials", False):
        answer = input(f"Save these settings to {_env_path(env_file)} for future runs? [y/N]: ")
        if answer.strip().lower() in {"y", "yes"}:
            _append_env_values(_env_path(env_file), env_updates)


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _normalize_db_url(url: str) -> str:
    if url.startswith("postgres://"):
        return "postgresql+psycopg://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        return "postgresql+psycopg://" + url[len("postgresql://"):]
    return url


def _safe_db_label(url: str) -> str:
    parsed = urlparse(url)
    if not parsed.scheme:
        return url
    host = parsed.hostname or ""
    db = (parsed.path or "").lstrip("/")
    return f"{parsed.scheme}://{host}/{db}"


def _choose_folder() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(title="Select resume data folder")
        root.destroy()
        if folder:
            return folder
    except Exception:
        pass

    folder = input("Enter resume data folder path: ").strip().strip('"')
    if not folder:
        raise SystemExit("ERROR: no folder selected.")
    return folder


# ---------------------------------------------------------------------------
# Healthcare resume parsing

PROFESSIONS = {
    "RN": ["registered nurse", " rn ", " rn,", "rn ", "bsn", "registered nurse (rn)"],
    "LPN": ["licensed practical nurse", " lpn", "lvn"],
    "CNA": ["certified nursing assistant", " cna"],
    "NP": ["nurse practitioner", " np ", "fnp", "acnp", "pmhnp"],
    "CRNA": ["nurse anesthetist", "crna"],
    "CNM": ["certified nurse midwife", "nurse midwife", " cnm"],
    "MD": ["physician", " md ", " m.d", "doctor of medicine", "mbchb"],
    "PA": ["physician assistant", "pa-c"],
    "RT": ["respiratory therapist", "registered respiratory"],
    "PT": ["physical therapist", " dpt", "physical therapy"],
    "OT": ["occupational therapist", "occupational therapy"],
    "PharmD": ["pharmacist", "pharm.d", "pharmd"],
}

SPECIALTIES = {
    "ICU": ["icu", "intensive care", "critical care", "ccu", "sicu", "micu"],
    "ER": ["emergency", " er ", " ed ", "emergency department", "emergency room"],
    "OR": ["operating room", "perioperative", " or nurse", "surgical services"],
    "PICU": ["picu", "pediatric intensive"],
    "NICU": ["nicu", "neonatal"],
    "Labor & Delivery": ["labor and delivery", "labor & delivery", "l&d", "postpartum", "mother baby"],
    "Med-Surg": ["med surg", "med-surg", "medical surgical", "medical-surgical"],
    "Telemetry": ["telemetry", "tele unit", "step down", "stepdown", "pcu"],
    "Oncology": ["oncology", "hematology", "chemo"],
    "Cath Lab": ["cath lab", "cardiac cath", "interventional"],
    "PACU": ["pacu", "post anesthesia", "recovery room"],
    "Dialysis": ["dialysis", "nephrology", "hemodialysis"],
    "Home Health": ["home health", "home care"],
    "Psych": ["psychiatric", "behavioral health", "psych "],
}

CERTIFICATIONS = [
    "BLS", "ACLS", "PALS", "CCRN", "TNCC", "NRP", "CEN", "ATLS", "CNOR",
    "RNC", "CPN", "CMSRN", "PCCN", "TCRN", "AWHONN", "STABLE", "NIHSS", "EKG",
]

US_STATES = {
    "AL": "alabama", "AK": "alaska", "AZ": "arizona", "AR": "arkansas",
    "CA": "california", "CO": "colorado", "CT": "connecticut", "DE": "delaware",
    "FL": "florida", "GA": "georgia", "HI": "hawaii", "ID": "idaho",
    "IL": "illinois", "IN": "indiana", "IA": "iowa", "KS": "kansas",
    "KY": "kentucky", "LA": "louisiana", "ME": "maine", "MD": "maryland",
    "MA": "massachusetts", "MI": "michigan", "MN": "minnesota", "MS": "mississippi",
    "MO": "missouri", "MT": "montana", "NE": "nebraska", "NV": "nevada",
    "NH": "new hampshire", "NJ": "new jersey", "NM": "new mexico", "NY": "new york",
    "NC": "north carolina", "ND": "north dakota", "OH": "ohio", "OK": "oklahoma",
    "OR": "oregon", "PA": "pennsylvania", "RI": "rhode island", "SC": "south carolina",
    "SD": "south dakota", "TN": "tennessee", "TX": "texas", "UT": "utah",
    "VT": "vermont", "VA": "virginia", "WA": "washington", "WV": "west virginia",
    "WI": "wisconsin", "WY": "wyoming",
}

CREDENTIALS = {
    "MD", "DO", "MBBS", "MBCHB", "DMD", "DDS", "DPM", "DPT", "PHARMD", "PSYD",
    "PHD", "MPH", "MSC", "MS", "MBA", "MHA", "MFA", "BFA", "BA", "BS",
    "RN", "BSN", "MSN", "NP",
    "FNP", "DNP", "PA", "PA-C", "FACP", "FAAP", "FACS", "FACC", "FAAD",
    "FACRO", "FASN", "FCAAAI", "FCAAI",
    "FAAAAI", "FACAAI", "MSCR", "FACE", "I", "II", "III", "IV",
}

PROFESSION_FROM_CRED = ["DO", "MD", "MBBS", "MBCHB", "DPM", "DMD", "DDS", "PharmD", "DPT", "DNP", "NP", "PA", "RN"]

APP_CODES = {"NP", "CRNA", "CNM", "FNP", "AGNP", "PMHNP", "ACNP", "AGACNP", "DNP", "PA", "PA-C", "APRN", "APP"}
APP_KW = ["nurse practitioner", "nurse anesthetist", "crna", "certified nurse midwife", "nurse midwife", "advanced practice", "physician assistant"]
PHYS_CODES = {"MD", "DO", "MBBS", "MBCHB"}
PHYS_KW = ["physician", "family medicine", "internal medicine"]
NURSE_CODES = {"RN", "LPN", "LVN", "CNA", "BSN", "MSN", "ADN"}
NURSE_KW = ["registered nurse", "licensed practical nurse", "licensed vocational nurse", "nursing assistant"]
ALLIED_CODES = {"RT", "PT", "OT", "RAD", "RADTECH", "RTR", "ARRT", "RDMS", "RVT", "RCIS", "CNMT", "RDCS"}
ALLIED_KW = [
    "radiologic technologist", "rad tech", "radiographer", "x-ray", "x ray",
    "ct technologist", "ct tech", "mri technologist", "mri tech", "ultrasound",
    "sonographer", "echo tech", "vascular technologist", "nuclear medicine",
    "interventional radiology", "cardiac cath", "cath lab", "respiratory therapist",
    "physical therapist", "occupational therapist", "surgical technologist", "allied",
]

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}")
YEARS_RE = re.compile(r"(\d{1,2})\+?\s*(?:years|yrs?)\b", re.IGNORECASE)
NPI_RE = re.compile(r"\bNPI[:#\s]*([0-9]{10})\b", re.IGNORECASE)
STATE_CODE_RE = "|".join(US_STATES)
STATE_NAME_RE = "|".join(re.escape(name) for name in US_STATES.values())
CITY_STATE_RE = re.compile(
    rf"\b([A-Z][A-Za-z.' -]{{1,45}}?),\s*({STATE_CODE_RE})\b(?:\s+\d{{5}}(?:-\d{{4}})?)?"
)
CITY_FULL_STATE_RE = re.compile(
    rf"\b([A-Z][A-Za-z.' -]{{1,45}}?),\s*({STATE_NAME_RE})\b(?:\s+\d{{5}}(?:-\d{{4}})?)?",
    re.IGNORECASE,
)
LOCATION_RE = re.compile(r"(.+?)\s*[\u2022\u00b7|]\s*(.+?),\s*([A-Z]{2})\b")
LOCATION_FALLBACK_RE = re.compile(r"^(.+?),\s*([A-Z]{2})\b")
STATE_LICENSE_RE = re.compile(r"\b([A-Z]{2})\s+State Medical License", re.IGNORECASE)
LICENSE_YEARS_RE = re.compile(r"(\d{4})\s*[-\u2013]\s*(\d{4})")
SECTION_HEADERS = (
    "education & training", "certifications & licensure", "awards",
    "publications", "professional memberships", "languages", "experience",
)
SECTION_TEXT = {
    "summary", "objective", "profile", "membership", "memberships",
    "organizational", "education", "education & training", "licensure",
    "licensure & certifications", "certification", "certifications",
    "certifications & licensure", "licenses", "skills", "experience",
    "clinical experience", "professional experience", "healthcare experience",
    "work experience", "employment", "references",
}
NOISY_WORDS = {
    "summary", "membership", "organizational", "education", "licensure",
    "certification", "certifications", "experience", "social media",
    "proficiency", "training", "provider training", "surgical", "clinical",
}
ADDRESS_WORDS = {
    "road", "rd", "street", "st", "drive", "dr", "court", "ct", "way",
    "avenue", "ave", "blvd", "boulevard", "lane", "ln", "apt", "suite",
    "floor", "unit", "po box", "p.o. box",
}
CITY_BAD_WORDS = SECTION_TEXT | {
    "provider", "hospital", "clinic", "department", "surgical", "summary",
    "education", "experience", "certification", "certifications", "licensure",
    "congress", "conference", "academy", "association", "society", "university",
    "college", "dermatology", "pediatric", "medical", "region",
}


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        pypdf = _import_or_exit("pypdf")
        reader = pypdf.PdfReader(str(path))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    if suffix == ".docx":
        docx = _import_or_exit("docx")
        doc = docx.Document(str(path))
        lines = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                lines.append(" ".join(c.text for c in row.cells))
        return "\n".join(lines)
    raise ValueError(f"Unsupported resume type: {suffix} (use .pdf or .docx)")


def classify_provider(profession_type=None, specialty=None, headline=None, title=None) -> str:
    code = (profession_type or "").upper().strip().strip(".")
    text = " ".join(x for x in (profession_type, specialty, headline, title) if x).lower()
    if code in APP_CODES or any(k in text for k in APP_KW):
        return "APP"
    if code in PHYS_CODES or any(k in text for k in PHYS_KW) or re.search(r"\bm\.?d\.?\b", text):
        return "Physicians"
    if code in NURSE_CODES or any(k in text for k in NURSE_KW):
        return "Nursing"
    if code in ALLIED_CODES or any(k in text for k in ALLIED_KW):
        return "Allied"
    return "Other"


def primary_american_board(cert_names) -> str | None:
    for cert in cert_names or []:
        value = str(cert).strip()
        lower = value.lower()
        if lower.startswith("american board of") and lower != "american board of physician specialties":
            return value
    return None


def _name_from_filename(path: Path) -> tuple[str, str]:
    stem = re.sub(r"(?i)(resume|cv|_|-)", " ", path.stem).strip()
    parts = [p for p in stem.split() if p.replace(".", "").isalpha()]
    parts = [p for p in parts if p.upper().strip(".") not in CREDENTIALS]
    if len(parts) >= 2:
        return parts[0].title(), parts[-1].title()
    if len(parts) == 1:
        return parts[0].title(), "Candidate"
    return "Unknown", "Candidate"


def _guess_name(text: str, path: Path) -> tuple[str, str]:
    for line in (ln.strip() for ln in text.splitlines()[:10]):
        name = _name_from_comma_header(line)
        if name:
            return name
        name = _candidate_name_from_line(line)
        if name:
            return name
    return _name_from_filename(path)


def _detect(text_lower: str, vocab: dict[str, list[str]]) -> str | None:
    for label, needles in vocab.items():
        if any(n in text_lower for n in needles):
            return label
    return None


def _clean_text(value: str | None) -> str:
    text = str(value or "")
    text = text.replace("â€¢", " ").replace("â€“", "-").replace("â€”", "-")
    text = text.replace("\u2022", " ").replace("\u2013", "-").replace("\u2014", "-")
    text = re.sub(r"[_\-]{4,}", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" \t\r\n-|,;")


def _is_noisy_card_text(value: str | None) -> bool:
    text = _clean_text(value)
    if not text:
        return True
    lower = text.lower()
    if len(text) > 90 or EMAIL_RE.search(text) or PHONE_RE.search(text):
        return True
    if re.search(r"^\d+\s+", text):
        return True
    if any(word in lower for word in NOISY_WORDS):
        return True
    if any(re.search(rf"\b{re.escape(word)}\b", lower) for word in ADDRESS_WORDS):
        return True
    return bool(re.search(r"[\u2022|]{1,}", text))


def _title_words(value: str | None) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    return " ".join(part[:1].upper() + part[1:].lower() for part in text.split())


# Words that appear in a bad "name" but never in a real person's name. We detect
# junk by STRUCTURE (role/résumé word, run-on, or digit), not by listing names.
JUNK_NAME_WORDS = {
    "unknown", "candidate", "provider", "providers", "resume", "cv", "n/a", "na",
    "none", "member", "applicant", "profile", "portfolio",
    "registered", "certified", "licensed", "nurse", "nursing", "physician",
    "surgeon", "doctor", "technician", "technologist", "practitioner",
    "assistant", "associate", "professional", "administrator", "administrative",
    "coordinator", "specialist", "therapist", "pharmacist", "director", "manager",
    "management", "supervisor", "consultant", "clinician", "caregiver", "aide",
    "worker", "staff", "curriculum", "vitae", "objective", "summary", "references",
    "reference", "experience", "experienced", "qualifications", "education",
    "skills", "skill", "certifications", "licensure", "employment", "history",
    "healthcare", "medical", "clinical", "hospital", "university", "college",
    "career", "seeking", "dedicated", "motivated", "organized", "acquired",
    "regional", "center", "travel", "staffing", "solutions", "services",
    "department", "unit", "team", "group",
    "adn", "bsn", "msn", "dnp", "aprn", "faan", "mba", "mph", "phd",
}
NAME_PLACEHOLDERS = JUNK_NAME_WORDS   # back-compat alias
_JUNK_SUBSTRINGS = tuple(w for w in JUNK_NAME_WORDS if len(w) >= 6)


def _bad_name(first: str | None, last: str | None) -> bool:
    """Structural junk-name test (mirrors app.importers.parsing.is_real_name)."""
    f = _clean_text(first).lower().strip(".")
    if not f or not any(ch.isalpha() for ch in f):
        return True
    full = f"{f} {_clean_text(last).lower().strip('.')}".strip()
    if any(ch.isdigit() for ch in full):
        return True
    for w in re.split(r"[\s\-]+", full):
        w = w.strip(".")
        if not w:
            continue
        if w in JUNK_NAME_WORDS or len(w) >= 16:
            return True
        if len(w) >= 12 and any(sub in w for sub in _JUNK_SUBSTRINGS):
            return True
    return False


def _title_city(value: str | None) -> str | None:
    city = _clean_text(value)
    if not city or len(city) <= 2 or len(city) > 60 or any(ch.isdigit() for ch in city):
        return None
    if len(city.split()) > 4:
        return None
    lower = city.lower()
    if any(re.search(rf"\b{re.escape(word)}\b", lower) for word in CITY_BAD_WORDS):
        return None
    if any(re.search(rf"\b{re.escape(word)}\b", lower) for word in ADDRESS_WORDS):
        return None
    compact = city.upper().replace(".", "").replace(" ", "")
    if compact in {c.upper().replace(".", "") for c in CREDENTIALS} | {"MFA", "BFA", "BA", "BS"}:
        return None
    if "." in city and not re.search(r"\bSt\.", city):
        return None
    return _title_words(city)


def _clean_state(value: str | None) -> str | None:
    state = _clean_text(value).upper()
    if state in US_STATES:
        return state
    lower = _clean_text(value).lower()
    for code, name in US_STATES.items():
        if lower == name:
            return code
    return None


def _clean_profession(value: str | None) -> str | None:
    text = _clean_text(value)
    if not text or _is_noisy_card_text(text):
        return None
    upper = text.upper().replace(".", "")
    if upper == "PA C":
        return "PA-C"
    if upper == "PHARMD":
        return "PharmD"
    if len(upper) <= 12 and re.fullmatch(r"[A-Z-]+", upper):
        return upper
    return text[:50]


def _line_tokens_for_name(line: str) -> list[str]:
    cleaned = _clean_text(line).replace(",", " ")
    cleaned = re.split(r"\s+\d", cleaned, 1)[0]
    tokens = []
    for token in re.split(r"\s+", cleaned):
        letters = token.replace(".", "").replace("'", "").replace("-", "")
        if letters.isalpha():
            tokens.append(token)
    return [t for t in tokens if t.upper().strip(".") not in CREDENTIALS]


def _candidate_name_from_line(line: str) -> tuple[str, str] | None:
    explicit = re.search(r"\bname\s*:\s*([^,@|0-9]+)", line, re.IGNORECASE)
    if explicit:
        tokens = _line_tokens_for_name(explicit.group(1))
        if 2 <= len(tokens) <= 5:
            return _title_words(tokens[0]) or tokens[0], _title_words(tokens[-1]) or tokens[-1]
    if _is_noisy_card_text(line) and not re.match(r"^[^\d@|,]+?\s+\d", line):
        return None
    tokens = _line_tokens_for_name(line)
    if 2 <= len(tokens) <= 5:
        return _title_words(tokens[0]) or tokens[0], _title_words(tokens[-1]) or tokens[-1]
    lower = _clean_text(line).lower()
    if lower in SECTION_TEXT or any(h in lower for h in SECTION_TEXT):
        return None
    return None


def _extract_profession(text: str, fallback: str | None = None) -> str | None:
    profession = _clean_profession(fallback)
    if profession:
        return profession
    header = "\n".join(text.splitlines()[:18])
    for code in ("CRNA", "CNM", "PA-C", "PharmD", "LPN", "LVN", "CNA", "RN", "NP", "PA", "RT", "PT", "OT", "MD", "DO", "MBBS"):
        if re.search(rf"\b{re.escape(code)}\b", header, re.IGNORECASE):
            return "PharmD" if code.lower() == "pharmd" else code.upper()
    return _clean_profession(_detect(" " + text.lower() + " ", PROFESSIONS))


def _clean_specialty(value: str | None) -> str | None:
    text = _clean_text(value)
    if _is_noisy_card_text(text):
        return None
    parts = [_clean_text(part) for part in text.split(",")]
    parts = [part for part in parts if part]
    for idx, part in enumerate(parts[:-1]):
        if part.upper().replace(".", "") in CREDENTIALS:
            text = ", ".join(parts[idx + 1:]).strip()
            break
    text = re.split(r"\s+\?\s+", text, maxsplit=1)[0].strip(" ?")
    return text[:100]


def _extract_location(text: str) -> tuple[str | None, str | None]:
    best: tuple[int, str, str] | None = None
    lines = [_clean_text(ln) for ln in text.splitlines() if _clean_text(ln)]
    for idx, line in enumerate(lines[:40]):
        for match in list(CITY_STATE_RE.finditer(line)) + list(CITY_FULL_STATE_RE.finditer(line)):
            city = _title_city(match.group(1))
            state = _clean_state(match.group(2))
            if not city or not state:
                continue
            has_location_context = (
                bool(re.search(r"\b\d{5}(?:-\d{4})?\b", line))
                or EMAIL_RE.search(line)
                or PHONE_RE.search(line)
                or "|" in line
                or any(re.search(rf"\b{re.escape(word)}\b", line.lower()) for word in ADDRESS_WORDS)
            )
            if state == "MD" and not has_location_context:
                continue
            if idx > 12 and not has_location_context:
                continue
            score = 100 - idx
            if re.search(r"\b\d{5}(?:-\d{4})?\b", line):
                score += 20
            if EMAIL_RE.search(line) or PHONE_RE.search(line) or "|" in line:
                score += 10
            if best is None or score > best[0]:
                best = (score, city, state)
    if best:
        return best[1], best[2]
    for code, name in US_STATES.items():
        if re.search(rf"\b{code}\b", text) or re.search(rf"\b{re.escape(name)}\b", text, re.IGNORECASE):
            return None, code
    return None, None


def _specialty_from_compact_location_line(line: str, city: str | None, state_code: str | None) -> str | None:
    """Extract specialty from lines like 'Oral Surgery Cape Elizabeth, ME'."""
    if not (city and state_code):
        return None
    text = _clean_text(line)
    city_state = f"{city}, {state_code}"
    idx = text.lower().rfind(city_state.lower())
    if idx <= 0:
        return None
    specialty = text[:idx].strip(" ,-|")
    return _clean_specialty(specialty)


def _specialty_from_comma_header(line: str) -> str | None:
    """Extract specialty from headers like 'Jane Smith, MD, Pediatric Gastroenterology'."""
    parts = [_clean_text(part) for part in str(line or "").split(",")]
    parts = [part for part in parts if part]
    if len(parts) < 3:
        return None
    credential = parts[1].upper().replace(".", "")
    if credential not in CREDENTIALS and credential not in {"PA-C", "PA C"}:
        return None
    return _clean_specialty(parts[2])


def _name_from_comma_header(line: str) -> tuple[str, str] | None:
    parts = [_clean_text(part) for part in str(line or "").split(",")]
    parts = [part for part in parts if part]
    if len(parts) < 2:
        return None
    credential = parts[1].upper().replace(".", "")
    if credential not in CREDENTIALS and credential not in {"PA-C", "PA C"}:
        return None
    tokens = _line_tokens_for_name(parts[0])
    if 2 <= len(tokens) <= 5:
        return _title_words(tokens[0]) or tokens[0], _title_words(tokens[-1]) or tokens[-1]
    return None


def _extract_years(text: str, fallback: int | None = None) -> int:
    value = int(fallback or 0)
    if value < 0 or value > 60:
        value = 0
    matches = [int(m.group(1)) for m in YEARS_RE.finditer(text)]
    matches = [m for m in matches if 0 < m <= 60]
    return max(matches) if matches else value


def format_resume_fields(fields: dict, text: str, path: Path) -> dict:
    out = dict(fields)
    first, last = _guess_name(text, path)
    if first and last and _bad_name(out.get("first_name"), out.get("last_name")):
        out["first_name"], out["last_name"] = first[:100], last[:100]
    else:
        out["first_name"] = (_title_words(out.get("first_name")) or "Unknown")[:100]
        out["last_name"] = (_title_words(out.get("last_name")) or "Provider")[:100]

    email_m = EMAIL_RE.search(text)
    phone_m = PHONE_RE.search(text)
    if email_m:
        out["email"] = email_m.group(0)[:255]
    if phone_m:
        out["phone"] = phone_m.group(0)[:30]
    out["profession_type"] = (_extract_profession(text, out.get("profession_type")) or None)
    specialty = _clean_specialty(out.get("specialty")) or _detect(" " + text.lower() + " ", SPECIALTIES)
    if not specialty:
        for line in text.splitlines()[:8]:
            specialty = _specialty_from_comma_header(line)
            if specialty:
                break
    parsed_city, parsed_state = _extract_location(text)
    city = _title_city(out.get("city")) or parsed_city
    state = _clean_state(out.get("state_code")) or parsed_state
    if parsed_city and parsed_state and city == parsed_city:
        state = parsed_state
    if not specialty and city and state:
        for line in text.splitlines()[:12]:
            specialty = _specialty_from_compact_location_line(line, city, state)
            if specialty:
                break
    out["specialty"] = specialty[:100] if specialty else None
    out["city"] = city[:120] if city else None
    out["state_code"] = state
    out["years_experience"] = _extract_years(text, out.get("years_experience"))
    headline = _clean_text(out.get("headline"))
    if _is_noisy_card_text(headline):
        headline = None
    if not headline:
        headline = " ".join(b for b in [out.get("specialty"), out.get("profession_type")] if b) or None
    out["headline"] = headline[:255] if headline else None
    out["american_board"] = primary_american_board(out.get("certifications")) or out.get("american_board")
    category = classify_provider(out.get("profession_type"), out.get("specialty"), out.get("headline"))
    out["provider_category"] = category if category != "Other" else (out.get("provider_category") or "Other")
    return out


def _strip_credentials(name_line: str) -> tuple[str, str, str | None]:
    cleaned = name_line.replace(",", " ")
    tokens = [t for t in re.split(r"\s+", cleaned) if t]
    profession = None
    for cred in PROFESSION_FROM_CRED:
        if any(t.upper().strip(".") == cred.upper() for t in tokens):
            profession = cred
            break
    name_tokens = [
        t for t in tokens
        if t.upper().strip(".") not in CREDENTIALS and "(" not in t and ")" not in t
    ]
    if not name_tokens:
        name_tokens = tokens
    first = name_tokens[0].title() if name_tokens else "Unknown"
    last = name_tokens[-1].title() if len(name_tokens) > 1 else "Provider"
    return first, last, profession


def _looks_structured(lines: list[str]) -> bool:
    head = "\n".join(lines[:6]).lower()
    all_text = "\n".join(lines).lower()
    return (
        "certifications & licensure" in all_text
        or "education & training" in head
        or bool(len(lines) > 1 and LOCATION_RE.search(lines[1]))
    )


def parse_structured(text: str, path: Path) -> dict:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    first, last, profession = _strip_credentials(lines[0]) if lines else ("Unknown", "Provider", None)

    specialty = city = state_code = headline = None
    loc_idx = None
    for i in range(1, min(len(lines), 8)):
        match = LOCATION_RE.search(lines[i])
        if match:
            specialty = _clean_specialty(match.group(1))
            city = match.group(2).strip()[:120]
            state_code = match.group(3).upper()
            loc_idx = i
            break

    if loc_idx is None:
        for i in range(1, min(len(lines), 8)):
            match = LOCATION_FALLBACK_RE.search(lines[i])
            if match and not any(h in lines[i].lower() for h in SECTION_HEADERS):
                city = match.group(1).strip()[:120]
                state_code = match.group(2).upper()
                loc_idx = i
                break

    if loc_idx is not None and loc_idx + 1 < len(lines):
        nxt = lines[loc_idx + 1]
        if not any(h in nxt.lower() for h in SECTION_HEADERS):
            headline = nxt[:255]
    if not headline:
        headline = specialty

    licenses = []
    seen_states = set()
    for i, line in enumerate(lines):
        match = STATE_LICENSE_RE.search(line)
        if not match:
            continue
        st = match.group(1).upper()
        if st in seen_states:
            continue
        seen_states.add(st)
        expiry_year = None
        for nxt in lines[i + 1: i + 3]:
            year_match = LICENSE_YEARS_RE.search(nxt)
            if year_match:
                expiry_year = int(year_match.group(2))
                break
        licenses.append({
            "state_code": st,
            "license_type": profession or "MD",
            "expiry_year": expiry_year,
        })

    if state_code is None and licenses:
        state_code = licenses[0]["state_code"]

    if not specialty and city and state_code:
        for line in lines[1:min(len(lines), 8)]:
            specialty = _specialty_from_compact_location_line(line, city, state_code)
            if specialty:
                break

    board_certs = [
        ln.strip()[:100] for ln in lines
        if ln.lower().startswith("american board of") or ln.lower().startswith("abms")
    ]
    board_certs = list(dict.fromkeys(board_certs))[:6]
    email_m = EMAIL_RE.search(text)
    phone_m = PHONE_RE.search(text)
    npi_m = NPI_RE.search(text)

    bio_parts = []
    if headline:
        bio_parts.append(headline)
    for i, line in enumerate(lines):
        if line.lower().startswith("education & training") and i + 1 < len(lines):
            bio_parts.append(f"Trained at {lines[i + 1]}")
            break
    bio = " - ".join(bio_parts) if bio_parts else None

    prof = profession or "MD"
    fields = {
        "first_name": first[:100],
        "last_name": last[:100],
        "email": email_m.group(0)[:255] if email_m else None,
        "phone": phone_m.group(0)[:30] if phone_m else None,
        "profession_type": prof[:50],
        "specialty": specialty,
        "city": city,
        "state_code": state_code,
        "certifications": board_certs,
        "licenses": licenses,
        "years_experience": 0,
        "npi_number": npi_m.group(1) if npi_m else None,
        "headline": headline[:255] if headline else None,
        "bio": bio,
        "american_board": primary_american_board(board_certs),
        "provider_category": classify_provider(prof, specialty, headline),
    }
    return format_resume_fields(fields, text, path)


def parse_resume(text: str, path: Path) -> dict:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if _looks_structured(lines):
        return parse_structured(text, path)

    tl = " " + text.lower() + " "
    first, last = _guess_name(text, path)
    email_m = EMAIL_RE.search(text)
    phone_m = PHONE_RE.search(text)
    years_m = YEARS_RE.search(text)
    npi_m = NPI_RE.search(text)

    profession = _detect(tl, PROFESSIONS)
    if profession is None:
        for code in ("CRNA", "CNM", "PharmD", "LPN", "CNA", "RN", "NP", "PA", "RT", "PT", "OT", "MD"):
            if re.search(rf"\b{re.escape(code)}\b", text):
                profession = code
                break

    specialty = _detect(tl, SPECIALTIES)
    certs = [c for c in CERTIFICATIONS if re.search(rf"\b{re.escape(c)}\b", text, re.IGNORECASE)]

    state_code = None
    for code, name in US_STATES.items():
        if re.search(rf"\b{code}\b", text) or name in tl:
            state_code = code
            break

    headline_bits = [b for b in [specialty, profession] if b]
    headline = " ".join(headline_bits) if headline_bits else None
    if headline and years_m:
        headline += f" - {years_m.group(1)} yrs"

    fields = {
        "first_name": first[:100],
        "last_name": last[:100],
        "email": email_m.group(0)[:255] if email_m else None,
        "phone": phone_m.group(0)[:30] if phone_m else None,
        "profession_type": profession[:50] if profession else None,
        "specialty": specialty,
        "city": None,
        "certifications": certs,
        "licenses": [],
        "years_experience": int(years_m.group(1)) if years_m else 0,
        "state_code": state_code,
        "npi_number": npi_m.group(1) if npi_m else None,
        "headline": headline[:255] if headline else None,
        "bio": None,
        "american_board": primary_american_board(certs),
        "provider_category": classify_provider(profession, specialty, headline),
    }
    return format_resume_fields(fields, text, path)


# ---------------------------------------------------------------------------
# Optional LLM extraction  (works with ANY OpenAI-compatible endpoint)
#
# Set these in .env to turn it on — no code change to switch providers:
#   LLM_ENABLED=true
#   LLM_BASE_URL=http://localhost:11434/v1      # Ollama (free, local)
#                https://api.groq.com/openai/v1  # Groq (free tier)
#                https://api.deepseek.com        # DeepSeek (very cheap)
#                https://api.openai.com/v1       # OpenAI (gpt-4o-mini)
#   LLM_API_KEY=...        # "ollama" or blank for a local server
#   LLM_MODEL=llama3.1:8b  # or llama-3.1-8b-instant / deepseek-chat / gpt-4o-mini
#
# The LLM fills the "hard" fields (name, specialty, city/state/zip, profession,
# years, contact, board); heuristics still handle licenses/certs. Every LLM value
# is re-validated with the same cleaners, and on ANY failure we fall back to the
# heuristic result — so enabling this can only improve data quality, never break.

import urllib.request

_LLM_SYSTEM = (
    "You extract structured data from a healthcare professional's resume. "
    "Respond with ONLY one JSON object and no other text."
)
_LLM_INSTRUCTIONS = (
    "Extract these fields from the resume. Use null when a field is absent — never guess.\n"
    'Return JSON exactly like:\n'
    '{"first_name":null,"last_name":null,"profession":null,"specialty":null,'
    '"city":null,"state":null,"zip":null,"years_experience":null,'
    '"email":null,"phone":null,"board_certification":null,"npi":null}\n'
    "- first_name/last_name = the ACTUAL person's name. If the top of the resume is a document "
    "title like 'Curriculum Vitae', 'Resume', 'Professional Summary', or a section header, "
    "set both names to null.\n"
    "- profession = their license/credential: MD, DO, RN, LPN, CNA, NP, CRNA, PA, RT, PT, OT, PharmD, etc.\n"
    "- state = 2-letter US code. zip = 5 digits. years_experience = a whole number.\n"
)


def _llm_config() -> dict:
    return {
        "enabled": _env_bool("LLM_ENABLED"),
        "base_url": (os.environ.get("LLM_BASE_URL") or "").rstrip("/"),
        "api_key": os.environ.get("LLM_API_KEY", ""),
        "model": os.environ.get("LLM_MODEL", ""),
        "timeout": float(os.environ.get("LLM_TIMEOUT", "60")),
    }


def _llm_raw(text: str, cfg: dict) -> dict:
    body = json.dumps({
        "model": cfg["model"],
        "temperature": 0,
        "messages": [
            {"role": "system", "content": _LLM_SYSTEM},
            {"role": "user", "content": _LLM_INSTRUCTIONS + "\n\nRESUME:\n" + text[:6000]},
        ],
        "response_format": {"type": "json_object"},
    }).encode("utf-8")
    req = urllib.request.Request(
        cfg["base_url"] + "/chat/completions", data=body, method="POST",
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + (cfg["api_key"] or "x"),
        },
    )
    with urllib.request.urlopen(req, timeout=cfg["timeout"]) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    content = payload["choices"][0]["message"]["content"]
    start, end = content.find("{"), content.rfind("}")
    return json.loads(content[start:end + 1])


def _llm_fields(text: str) -> dict | None:
    """Return validated fields from the LLM, or None if disabled/failed/invalid."""
    cfg = _llm_config()
    if not (cfg["enabled"] and cfg["base_url"] and cfg["model"]):
        return None
    try:
        raw = _llm_raw(text, cfg)
    except Exception as exc:  # noqa: BLE001 — never let the LLM break the import
        print(f"  (LLM extract failed: {str(exc)[:120]}; using heuristics)", file=sys.stderr)
        return None

    out: dict = {}
    first, last = _title_words(raw.get("first_name")), _title_words(raw.get("last_name"))
    if first and last and not _bad_name(first, last):
        out["first_name"], out["last_name"] = first[:100], last[:100]
    prof = _clean_profession(raw.get("profession"))
    if prof:
        out["profession_type"] = prof
    spec = _clean_specialty(raw.get("specialty"))
    if spec:
        out["specialty"] = spec[:100]
    city = _title_city(raw.get("city"))
    if city:
        out["city"] = city[:120]
    state = _clean_state(raw.get("state"))
    if state:
        out["state_code"] = state
    ye = raw.get("years_experience")
    if isinstance(ye, (int, float)) and 0 < ye <= 60:
        out["years_experience"] = int(ye)
    if raw.get("email"):
        em = EMAIL_RE.search(str(raw["email"]))
        if em:
            out["email"] = em.group(0)[:255]
    if raw.get("phone"):
        pm = PHONE_RE.search(str(raw["phone"]))
        if pm:
            out["phone"] = pm.group(0)[:30]
    board = raw.get("board_certification")
    if board and str(board).lower().startswith("american board of"):
        out["american_board"] = str(board)[:150]
    if raw.get("npi") and re.fullmatch(r"\d{10}", str(raw["npi"]).strip()):
        out["npi_number"] = str(raw["npi"]).strip()
    if raw.get("zip") and re.fullmatch(r"\d{5}", str(raw["zip"]).strip()):
        out["zip_code"] = str(raw["zip"]).strip()
    return out or None


def parse_resume_smart(text: str, path: Path) -> dict:
    """Heuristic parse, refined by the LLM when configured (validated + fallback)."""
    fields = parse_resume(text, path)
    llm = _llm_fields(text)
    if llm:
        fields.update({k: v for k, v in llm.items() if v})
        # Re-derive the category from the (now more accurate) profession/specialty.
        fields["provider_category"] = classify_provider(
            fields.get("profession_type"), fields.get("specialty"), fields.get("headline"))
    return fields


# ---------------------------------------------------------------------------
# Cloudflare R2

_S3_CLIENT = None


def _s3_client():
    global _S3_CLIENT
    if _S3_CLIENT is None:
        boto3 = _import_or_exit("boto3")
        botocore_config = _import_or_exit("botocore.config")
        _S3_CLIENT = boto3.client(
            "s3",
            endpoint_url=os.environ.get("S3_ENDPOINT_URL") or None,
            region_name=os.environ.get("S3_REGION", "auto"),
            aws_access_key_id=os.environ.get("S3_ACCESS_KEY"),
            aws_secret_access_key=os.environ.get("S3_SECRET_KEY"),
            config=botocore_config.Config(signature_version="s3v4"),
        )
    return _S3_CLIENT


def _content_type(path: Path) -> str:
    if path.suffix.lower() == ".pdf":
        return "application/pdf"
    if path.suffix.lower() == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return mimetypes.guess_type(path.name)[0] or "application/octet-stream"


def _url_for_key(key: str) -> str:
    public_base = (os.environ.get("S3_PUBLIC_BASE_URL") or "").rstrip("/")
    if _env_bool("S3_PUBLIC") and public_base:
        return f"{public_base}/{key}"
    return f"/files/{key}"


def _object_exists(key: str) -> bool:
    botocore_exceptions = _import_or_exit("botocore.exceptions")
    try:
        _s3_client().head_object(Bucket=os.environ["S3_BUCKET"], Key=key)
        return True
    except botocore_exceptions.ClientError as exc:
        code = str(exc.response.get("Error", {}).get("Code", ""))
        if code in {"404", "NoSuchKey", "NotFound"}:
            return False
        raise


def _upload_if_needed(path: Path, key: str) -> str:
    if _object_exists(key):
        return _url_for_key(key)

    extra = {"ContentType": _content_type(path)}
    acl = os.environ.get("S3_ACL", "").strip()
    if acl:
        extra["ACL"] = acl

    with open(path, "rb") as f:
        _s3_client().upload_fileobj(f, os.environ["S3_BUCKET"], key, ExtraArgs=extra)
    return _url_for_key(key)


# ---------------------------------------------------------------------------
# Neon raw SQL

REQUIRED_PROFILE_COLUMNS = {
    "profile_id", "user_id", "first_name", "last_name", "headline", "bio",
    "phone", "email", "specialty", "profession_type", "provider_category",
    "american_board", "years_experience", "city", "state_code", "lat", "lng",
    "open_to_work", "job_type_prefs", "pay_min_hourly", "available_date",
    "npi_number", "profile_photo_url", "resume_url", "completion_score",
    "source", "search_text", "created_at", "updated_at", "is_listable",
}


def _create_engine():
    sqlalchemy = _import_or_exit("sqlalchemy")
    db_url = _normalize_db_url(os.environ.get("DATABASE_URL", ""))
    if not db_url:
        raise SystemExit("ERROR: DATABASE_URL is missing.")
    if not db_url.startswith("postgresql"):
        raise SystemExit("ERROR: DATABASE_URL must point to Neon Postgres.")
    return sqlalchemy.create_engine(db_url, future=True, pool_pre_ping=True)


def _text(sql: str):
    sqlalchemy = _import_or_exit("sqlalchemy")
    return sqlalchemy.text(sql)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _validate_schema(engine) -> None:
    required_tables = {"profiles", "certifications", "licenses", "profile_skills"}
    with engine.connect() as conn:
        tables = set(conn.execute(_text("""
            SELECT table_name
            FROM information_schema.tables
            WHERE table_schema = 'public'
        """)).scalars().all())
        missing_tables = sorted(required_tables - tables)
        if missing_tables:
            raise SystemExit(
                "ERROR: Neon database is missing HealthBoard table(s): "
                + ", ".join(missing_tables)
            )

        profile_columns = set(conn.execute(_text("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = 'profiles'
        """)).scalars().all())
        missing_columns = sorted(REQUIRED_PROFILE_COLUMNS - profile_columns)
        if missing_columns:
            raise SystemExit(
                "ERROR: profiles table is missing column(s): "
                + ", ".join(missing_columns)
                + ". Deploy/run the latest HealthBoard migrations before uploading."
            )


def _existing_profile_id(conn, key: str) -> str | None:
    row = conn.execute(
        _text("SELECT profile_id FROM profiles WHERE resume_url LIKE :needle LIMIT 1"),
        {"needle": f"%{key}%"},
    ).first()
    return row[0] if row else None


# --- Person-level de-duplication -------------------------------------------
# A résumé for someone already in the DB (different file: PDF vs DOCX, an
# updated copy, a second upload) must NOT create a second profile. We match on
# NPI, then email, then last name + phone (last 10 digits) — the same keys the
# app/dedup_profiles.py cleanup uses. The existing identities are loaded into
# memory once so each file is an O(1) set lookup, not a DB query.

def _phone10(v: str | None) -> str | None:
    digits = re.sub(r"\D", "", v or "")
    return digits[-10:] if len(digits) >= 10 else None


def _identity_keys(fields: dict):
    """Return (npi, email, (last_name, phone10)) — any of which may be None."""
    npi = (str(fields.get("npi_number") or "").strip()) or None
    email = (str(fields.get("email") or "").strip().lower()) or None
    if email and "@" not in email:
        email = None
    ph = _phone10(fields.get("phone"))
    last = (str(fields.get("last_name") or "").strip().lower()) or None
    phone_key = (last, ph) if (ph and last) else None
    return npi, email, phone_key


def _load_existing_identity(engine) -> dict:
    seen = {"npi": set(), "email": set(), "phone": set()}
    with engine.begin() as conn:
        rows = conn.execute(
            _text("SELECT npi_number, email, last_name, phone FROM profiles")).all()
    for npi, email, last, phone in rows:
        if npi:
            seen["npi"].add(str(npi).strip())
        if email and "@" in str(email):
            seen["email"].add(str(email).strip().lower())
        ph = _phone10(phone)
        if ph and last:
            seen["phone"].add((str(last).strip().lower(), ph))
    return seen


def _identity_dup(seen: dict, fields: dict) -> str | None:
    """Return which key matched an existing person, or None if new."""
    npi, email, phone_key = _identity_keys(fields)
    if npi and npi in seen["npi"]:
        return "npi"
    if email and email in seen["email"]:
        return "email"
    if phone_key and phone_key in seen["phone"]:
        return "phone"
    return None


def _identity_remember(seen: dict, fields: dict) -> None:
    npi, email, phone_key = _identity_keys(fields)
    if npi:
        seen["npi"].add(npi)
    if email:
        seen["email"].add(email)
    if phone_key:
        seen["phone"].add(phone_key)


def _unique_npi(conn, npi: str | None) -> str | None:
    if not npi:
        return None
    row = conn.execute(
        _text("SELECT profile_id FROM profiles WHERE npi_number = :npi LIMIT 1"),
        {"npi": npi},
    ).first()
    return None if row else npi


def _completion(fields: dict, has_resume: bool = True) -> int:
    score = 20 if has_resume else 0
    if fields.get("headline"):
        score += 10
    if fields.get("bio"):
        score += 10
    if fields.get("specialty"):
        score += 15
    if fields.get("profession_type"):
        score += 10
    if fields.get("years_experience"):
        score += 10
    if fields.get("city") and fields.get("state_code"):
        score += 10
    if fields.get("email") or fields.get("phone"):
        score += 5
    if fields.get("certifications"):
        score += 5
    if fields.get("licenses"):
        score += 10
    return min(score, 100)


LICENSE_FULL_NAMES = {
    "RN": "Registered Nurse", "LPN": "Licensed Practical Nurse",
    "LVN": "Licensed Vocational Nurse", "CNA": "Certified Nursing Assistant",
    "NP": "Nurse Practitioner", "FNP": "Family Nurse Practitioner",
    "DNP": "Doctor of Nursing Practice", "CRNA": "Certified Registered Nurse Anesthetist",
    "CNM": "Certified Nurse Midwife", "PA": "Physician Assistant", "MD": "Physician",
    "DO": "Doctor of Osteopathic Medicine", "RT": "Respiratory Therapist",
    "PT": "Physical Therapist", "OT": "Occupational Therapist",
}


def _search_text(fields: dict) -> str:
    parts = [
        fields.get("first_name"), fields.get("last_name"), fields.get("headline"),
        fields.get("bio"), fields.get("specialty"), fields.get("profession_type"),
        fields.get("city"), fields.get("state_code"), fields.get("american_board"),
        fields.get("provider_category"),
        # Full license name so "registered nurse" matches an RN profile.
        LICENSE_FULL_NAMES.get(str(fields.get("profession_type") or "").upper().strip(".")),
    ]
    return " ".join(str(p) for p in parts if p).lower()


def _insert_profile(conn, path: Path, key: str, resume_url: str, fields: dict) -> str:
    profile_id = str(uuid.uuid4())
    now = _utcnow()
    params = {
        "profile_id": profile_id,
        "first_name": fields["first_name"],
        "last_name": fields["last_name"],
        "headline": fields.get("headline"),
        "bio": fields.get("bio"),
        "phone": fields.get("phone"),
        "email": fields.get("email"),
        "specialty": fields.get("specialty"),
        "profession_type": fields.get("profession_type"),
        "provider_category": fields.get("provider_category"),
        "american_board": fields.get("american_board"),
        "years_experience": fields.get("years_experience") or 0,
        "city": fields.get("city"),
        "state_code": fields.get("state_code"),
        "npi_number": _unique_npi(conn, fields.get("npi_number")),
        "resume_url": resume_url,
        "completion_score": _completion(fields),
        "search_text": _search_text(fields),
        "created_at": now,
        "updated_at": now,
        "job_type_prefs": json.dumps([]),
        "source": "resume_parse",
        # Hide parser-junk names from the directory at insert time.
        "is_listable": not _bad_name(fields.get("first_name"), fields.get("last_name")),
    }
    conn.execute(_text("""
        INSERT INTO profiles (
            profile_id, user_id, first_name, last_name, headline, bio, phone, email,
            specialty, profession_type, provider_category, american_board,
            years_experience, city, state_code, lat, lng, open_to_work,
            job_type_prefs, pay_min_hourly, available_date, npi_number,
            profile_photo_url, resume_url, completion_score, source, search_text,
            is_listable, created_at, updated_at
        )
        VALUES (
            :profile_id, NULL, :first_name, :last_name, :headline, :bio, :phone, :email,
            :specialty, :profession_type, :provider_category, :american_board,
            :years_experience, :city, :state_code, NULL, NULL, TRUE,
            CAST(:job_type_prefs AS JSON), NULL, NULL, :npi_number,
            NULL, :resume_url, :completion_score, CAST(:source AS profilesource), :search_text,
            :is_listable, :created_at, :updated_at
        )
    """), params)

    for cert in fields.get("certifications") or []:
        conn.execute(_text("""
            INSERT INTO certifications (
                cert_id, profile_id, cert_name, issuing_body, issue_date,
                expiry_date, cert_number, created_at
            )
            VALUES (
                :cert_id, :profile_id, :cert_name, NULL, NULL, NULL, NULL, :created_at
            )
        """), {
            "cert_id": str(uuid.uuid4()),
            "profile_id": profile_id,
            "cert_name": str(cert)[:100],
            "created_at": now,
        })

    for lic in fields.get("licenses") or []:
        expiry = date(int(lic["expiry_year"]), 12, 31) if lic.get("expiry_year") else None
        conn.execute(_text("""
            INSERT INTO licenses (
                license_id, profile_id, license_type, license_number, state_code,
                status, issued_date, expiry_date, verified_at, verification_source,
                is_compact, created_at
            )
            VALUES (
                :license_id, :profile_id, :license_type, :license_number, :state_code,
                CAST(:status AS licensestatus), NULL, :expiry_date, NULL, :verification_source,
                FALSE, :created_at
            )
        """), {
            "license_id": str(uuid.uuid4()),
            "profile_id": profile_id,
            "license_type": (lic.get("license_type") or fields.get("profession_type") or "MD")[:50],
            "license_number": "(from resume)",
            "state_code": str(lic["state_code"]).upper()[:2],
            "status": "active",
            "expiry_date": expiry,
            "verification_source": "resume",
            "created_at": now,
        })

    if fields.get("specialty"):
        conn.execute(_text("""
            INSERT INTO profile_skills (skill_id, profile_id, name, years)
            VALUES (:skill_id, :profile_id, :name, :years)
        """), {
            "skill_id": str(uuid.uuid4()),
            "profile_id": profile_id,
            "name": fields["specialty"][:100],
            "years": fields.get("years_experience") or None,
        })

    return profile_id


# ---------------------------------------------------------------------------
# Import loop

def _iter_files(folder: str, limit: int | None = None) -> list[Path]:
    root = Path(folder).expanduser()
    if not root.is_dir():
        raise SystemExit(f"ERROR: not a folder: {root}")
    files = sorted(p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED)
    return files[:limit] if limit else files


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _key_for(path: Path, digest: str, prefix: str) -> str:
    return f"{prefix.strip('/')}/{digest[:2]}/{digest}{path.suffix.lower()}"


def _load_done_hashes(path: Path) -> set[str]:
    if not path.exists():
        return set()
    done: set[str] = set()
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("status") == "imported" and rec.get("sha256"):
                done.add(rec["sha256"])
    return done


def _append_manifest(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, sort_keys=True) + "\n")


def _validate_upload_env(args: argparse.Namespace) -> None:
    if args.dry_run:
        return
    db_url = _normalize_db_url(os.environ.get("DATABASE_URL", ""))
    if not db_url.startswith("postgresql"):
        raise SystemExit("ERROR: DATABASE_URL must point to Neon Postgres.")
    if not _env_bool("STORAGE_ENABLED"):
        raise SystemExit("ERROR: STORAGE_ENABLED must be true.")
    missing = []
    for key, _label, _secret in REQUIRED_UPLOAD_ENV:
        if not os.environ.get(key):
            missing.append(key)
    if missing:
        raise SystemExit("ERROR: Missing required .env setting(s): " + ", ".join(missing))


def import_folder(args: argparse.Namespace) -> dict:
    files = _iter_files(args.folder, args.limit)
    manifest_path = _env_path(args.manifest)
    done_hashes = set() if args.ignore_manifest else _load_done_hashes(manifest_path)
    db_url = _normalize_db_url(os.environ.get("DATABASE_URL", ""))

    print(f"Database: {_safe_db_label(db_url) if db_url else '(dry-run/no database)'}")
    print(f"Storage enabled: {_env_bool('STORAGE_ENABLED')}")
    if os.environ.get("S3_BUCKET"):
        print(f"R2/S3 bucket: {os.environ.get('S3_BUCKET')}")
    if os.environ.get("S3_ENDPOINT_URL"):
        print(f"R2/S3 endpoint: {urlparse(os.environ.get('S3_ENDPOINT_URL', '')).netloc}")
    print(f"Found {len(files)} supported resume file(s).")

    _validate_upload_env(args)
    engine = None if args.dry_run else _create_engine()
    if engine is not None:
        _validate_schema(engine)

    # Load existing identities so a person already in the DB isn't re-imported
    # as a duplicate from a different file. Grows as we insert, so dupes within
    # this same run are caught too.
    seen_identity = {"npi": set(), "email": set(), "phone": set()}
    if engine is not None:
        seen_identity = _load_existing_identity(engine)
        print(f"Identity index: {len(seen_identity['email']):,} emails, "
              f"{len(seen_identity['phone']):,} phones, {len(seen_identity['npi']):,} NPIs.")

    stats = {"created": 0, "skipped": 0, "failed": 0, "total": len(files)}
    started = monotonic()

    for i, path in enumerate(files, start=1):
        try:
            digest = _sha256_file(path)
            key = _key_for(path, digest, args.prefix)

            if digest in done_hashes:
                stats["skipped"] += 1
                print(f"[{i}/{len(files)}] SKIP manifest {path.name}")
                continue

            if engine is not None:
                with engine.begin() as conn:
                    existing_id = _existing_profile_id(conn, key)
                if existing_id:
                    stats["skipped"] += 1
                    print(f"[{i}/{len(files)}] SKIP existing {path.name} -> {existing_id}")
                    continue

            text = extract_text(path)
            fields = parse_resume_smart(text, path)
            label = f"{fields['first_name']} {fields['last_name']}"

            dup_by = _identity_dup(seen_identity, fields)
            if dup_by:
                stats["skipped"] += 1
                print(f"[{i}/{len(files)}] SKIP dup-{dup_by} {label:24.24s} <- {path.name}")
                continue

            if args.dry_run:
                stats["created"] += 1
                _identity_remember(seen_identity, fields)
                print(
                    f"[{i}/{len(files)}] WOULD IMPORT {path.name} -> "
                    f"{label} | {fields.get('profession_type') or '?'} | "
                    f"email={bool(fields.get('email'))} phone={bool(fields.get('phone'))}"
                )
                continue

            resume_url = _upload_if_needed(path, key)
            with engine.begin() as conn:
                existing_id = _existing_profile_id(conn, key)
                if existing_id:
                    stats["skipped"] += 1
                    print(f"[{i}/{len(files)}] SKIP existing {path.name} -> {existing_id}")
                    continue
                profile_id = _insert_profile(conn, path, key, resume_url, fields)

            _identity_remember(seen_identity, fields)
            stats["created"] += 1
            _append_manifest(manifest_path, {
                "status": "imported",
                "sha256": digest,
                "file": str(path),
                "key": key,
                "profile_id": profile_id,
            })
            print(f"[{i}/{len(files)}] IMPORTED {label:30.30s} <- {path.name}")

        except Exception as exc:
            stats["failed"] += 1
            print(f"[{i}/{len(files)}] FAIL {path.name}: {exc}", file=sys.stderr)

    elapsed = max(monotonic() - started, 0.001)
    print(f"\nSummary: {stats}")
    print(f"Elapsed: {elapsed / 60:.1f} min ({len(files) / elapsed:.2f} files/sec)")
    return stats


def main() -> int:
    parser = argparse.ArgumentParser(description="Standalone upload of resumes to Cloudflare R2 and Neon")
    parser.add_argument("folder", nargs="?", help="Folder containing .pdf/.docx resumes")
    parser.add_argument("--dry-run", action="store_true", help="Parse only; do not upload or write DB rows")
    parser.add_argument("--limit", type=int, default=None, help="Only process first N files")
    parser.add_argument("--prefix", default=DEFAULT_PREFIX, help="R2 key prefix")
    parser.add_argument("--manifest", default=DEFAULT_MANIFEST, help="JSONL progress manifest path")
    parser.add_argument("--ignore-manifest", action="store_true", help="Do not skip hashes in manifest")
    parser.add_argument("--env-file", default=".env", help="Env file path, relative to this script by default")
    parser.add_argument("--no-save-credentials", action="store_true", help="Do not offer to save prompted credentials")
    parser.add_argument("--install-deps", action="store_true", help="Install missing Python packages, then continue")
    args = parser.parse_args()

    if args.install_deps:
        _install_missing_dependencies()

    if not args.folder:
        args.folder = _choose_folder()

    _prepare_runtime_env(args)
    import_folder(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
