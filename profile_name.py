"""Person-name detection shared by the screenshot converter and the uploader.

The name of the person on a profile page is available from several
independent places: the page layout (the largest text near the top), the
page's own copy ("<Name> is on Doximity", "See <Name>'s full profile"), the
rendered DOCX title paragraph, and the source filename. Any one of them can
be missing or wrong on a given image (OCR drops a line, an avatar glyph
glues itself to the name, an old DOCX has no title), so nothing here trusts a
single line. Every candidate is first checked for the *shape* of a person's
name, then scored by how many independent sources agree with it. The result
carries a confidence and a human-readable evidence trail so a bad detection
can be flagged for review instead of written to the database.
"""
from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from pathlib import Path

# A detection at or above this confidence is stored without review.
REVIEW_BELOW = 0.6

# Degree / licence tokens that trail a name ("Jane Smith NP", "MD, PhD").
CREDENTIALS = {
    "md", "do", "mbbs", "mbchb", "dds", "dmd", "dpm", "dpt", "dc", "nd", "od",
    "pharmd", "mba", "mph", "msc", "ms", "ma", "phd", "np", "pa", "pa-c", "pac",
    "dnp", "fnp", "fnp-c", "fnp-bc", "crna", "cnm", "rn", "aprn", "aprn-bc",
    "cnp", "whnp", "whnp-bc", "pmhnp", "pmhnp-bc", "agnp", "acnp", "agacnp",
    "anp", "gnp", "pnp", "cns", "bsn", "msn", "faan", "facs", "facc", "facp",
    "facog", "frcpc", "facaai", "faap", "mhs", "mha", "med", "edd", "psyd",
    "ot", "pt", "rd", "ldn", "lpn", "lvn", "cna", "emt", "rt", "jd", "dvm",
    "mdiv", "mdcm", "bs", "ba", "bsc", "dsc", "mmed", "mmsc", "mbe", "mhsa",
    "msed", "mhpe", "lcdr", "esq",
}

# Skipped but, unlike a credential, they come *before* the name.
HONORIFICS = {"dr", "prof", "mr", "mrs", "ms", "miss", "sir", "dame"}

# Word classes that occur in the resume lines a mis-detected "name" comes from
# (institutions, roles, section titles, specialties, page chrome) but never
# inside a person's name. Matching is per token, so it is layout-independent
# and does not depend on any particular profile.
NON_NAME_WORDS = {
    # institutions / organisations
    "university", "universite", "college", "school", "academy", "institute",
    "institution", "hospital", "hospitals", "medical", "medicine", "center",
    "centre", "clinic", "clinics", "health", "healthcare", "nursing",
    "department", "faculty", "foundation", "association", "society", "board",
    "american", "national", "international", "regional", "county", "state",
    "city", "group", "associates", "partners", "practice", "services",
    "system", "systems", "network", "program", "programs", "campus", "office",
    "corporation", "inc", "llc", "ltd", "company",
    # roles / titles
    "nurse", "nurses", "practitioner", "practitioners", "physician",
    "physicians", "doctor", "surgeon", "assistant", "specialist", "therapist",
    "pharmacist", "resident", "director", "manager", "coordinator",
    "professor", "attending", "registered", "certified", "licensed",
    "provider", "providers", "candidate", "member", "staff", "owner",
    "technician", "technologist", "clinician",
    # section / résumé words
    "education", "training", "certifications", "certification", "licensure",
    "license", "licenses", "awards", "honors", "honours", "recognition",
    "publications", "presentations", "memberships", "membership", "languages",
    "experience", "summary", "profile", "address", "contact", "phone", "fax",
    "email", "about", "overview", "biography", "bio", "skills", "references",
    "objective", "employment", "history", "curriculum", "vitae", "resume",
    "masters", "bachelor", "bachelors", "doctorate", "residency", "fellowship",
    "internship", "degree", "graduate", "undergraduate",
    # specialties (lines like "Pediatric Cardiology" / "Women's Health ...")
    "cardiology", "oncology", "surgery", "surgical", "psychiatry",
    "anesthesiology", "dermatology", "radiology", "neurology", "orthopedic",
    "orthopedics", "obstetrics", "gynecology", "allergy", "immunology",
    "asthma", "conditions", "pediatric", "pediatrics", "internal", "emergency",
    "urgent", "primary", "care", "telehealth", "clinical", "womens", "women's",
    "women’s", "mens", "men's", "adult", "geriatric", "hospice", "palliative",
    "urology", "nephrology", "pulmonary", "critical", "endocrinology",
    "gastroenterology", "hematology", "infectious", "disease", "diseases",
    "rheumatology", "ophthalmology", "otolaryngology", "pathology",
    # street-address lines from the contact card
    "street", "st", "ave", "avenue", "blvd", "boulevard", "road", "rd", "drive",
    "lane", "ln", "court", "ct", "suite", "ste", "highway", "hwy", "parkway",
    "pkwy", "place", "pl", "way", "north", "south", "east", "west", "floor",
    "building", "bldg", "plaza", "square", "box",
    # page chrome / placeholders
    "doximity", "logo", "join", "view", "full", "similar", "see", "already",
    "account", "unknown", "na", "n/a", "none", "null", "image", "screenshot",
    "test", "sample",
}

# Function words: never inside a person's name, common inside an institution
# ("University of X", "X and Y Associates"). Only checked after the first
# token, so given names such as "An" or particles like "De" are untouched.
FUNCTION_WORDS = {"of", "and", "the", "for", "at", "in", "on", "to", "with", "from", "by", "or"}

# Section headings that mark the end of the profile header; a person's name
# is never found below the first one. Compared on a normalised line.
SECTION_HEADINGS = {
    "education & training", "education and training", "education",
    "certifications & licensure", "certifications and licensure",
    "certifications", "licensure", "awards, honors, & recognition", "awards",
    "publications & presentations", "publications", "professional memberships",
    "memberships", "hospital affiliations", "board certifications",
    "clinical interests", "practice locations", "practice address", "languages",
    "summary", "experience", "skills", "references", "objective", "employment",
    "work experience", "professional experience", "clinical experience",
}

GENERATION_SUFFIXES = {"jr", "sr", "i", "ii", "iii", "iv", "v"}

_TOKEN_RE = re.compile(r"^[A-Za-z][A-Za-z'’.\-]*$")
_HEX_RE = re.compile(r"^[0-9a-f]{6,}$")
_PAREN_RE = re.compile(r"\([^)]*(?:\)|$)")     # "(Uygungil)" and a trailing "(He/Him"

# Page copy that repeats the person's name. These are the join/promo lines the
# converter otherwise discards, so they are an independent read of the name.
PAGE_TEXT_PATTERNS = [
    re.compile(r"^(?:Dr\.?\s+)?(.{2,80}?)\s+is on\s+Doximity\b", re.I),
    re.compile(r"\bSee\s+(?:Dr\.?\s+)?(.{2,80}?)['’]?s\s+full\s+profile\b", re.I),
]


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9&' ]", "", (text or "").lower()).strip()


def is_section_heading(line: str) -> bool:
    n = _norm(line).replace("  ", " ")
    if not n:
        return False
    if n in SECTION_HEADINGS:
        return True
    # A heading may carry OCR noise at the end ("Education & Training —").
    return any(n.startswith(h) and len(n) <= len(h) + 4 for h in SECTION_HEADINGS if len(h) >= 8)


def is_honorific(token: str) -> bool:
    return token.strip(".,;:").lower() in HONORIFICS


def is_credential(token: str) -> bool:
    t = token.strip(".,;:").lower().replace(".", "")   # "M.D." / "D.Sc" -> "md" / "dsc"
    if not t:
        return False
    if t in CREDENTIALS:
        return True
    # "FNP-BC", "PMHNP-C", "RN-BC": a licence suffix, whatever the prefix.
    return bool(re.fullmatch(r"[a-z]{2,6}-(?:bc|c)", t))


def strip_leading_junk(text: str) -> str:
    """Remove OCR artifacts glued to the front of a name line.

    On a profile page the name sits beside an avatar. When the avatar is an
    initial on a coloured disc, OCR reads it as a digit or a lone letter ("4
    Aaron Pinion", "A Amanda Trott"); when it is a broken image, as a run of
    symbols or a lowercase echo of the name ("__ Adrianne", "abbie Abbie").
    Only tokens that cannot start a real name are dropped, so a genuine
    initial ("J. Thaddaeus Abbott") survives.
    """
    text = (text or "").strip()
    for _ in range(4):
        m = re.match(r"^(\S+)\s+(\S+)(.*)$", text, re.S)
        if not m:
            break
        first, nxt, rest = m.group(1), m.group(2), m.group(3)
        drop = False
        if not re.search(r"[A-Za-z]", first):
            drop = True                                   # "4", "__", "*"
        elif re.search(r"[0-9@{}<>~^`\\=|_()]", first):
            drop = True                                   # "R==~", "a1", "CY)"
        elif (re.fullmatch(r"[a-z][a-z'’.\-]*", first) and nxt[:1].isupper()
              and (len(first) <= 2 or nxt.lower().startswith(first.lower()) or first.lower().startswith(nxt.lower()))):
            drop = True                                   # "abbie Abbie", "ea Fellowship"
        elif (len(a := re.sub(r"[^A-Za-z]", "", first).lower()) >= 3
              and len(b := re.sub(r"[^A-Za-z]", "", nxt).lower()) >= 3
              and a != b and a.endswith(b) and len(a) - len(b) <= 2):
            drop = True                                   # avatar initial glued: "Ajing Jing", "AZjosh Josh"
        elif (re.fullmatch(r"[A-Za-z]", first) and nxt[:1].isalpha()
              and nxt[:1].upper() == first.upper() and len(nxt) >= 2):
            drop = True                                   # avatar initial: "A Amanda"
        if not drop:
            break
        text = f"{nxt}{rest}".strip()
    return text


def name_tokens(text: str) -> list[str]:
    """Split a name line into its name words, dropping credentials, a maiden
    name in parentheses and anything that is not a word."""
    text = _PAREN_RE.sub(" ", text or "")
    text = strip_leading_junk(text)
    # "Jane Smith, MD, Pediatric Gastroenterology": the name ends at the
    # first comma-separated segment that is a credential.
    segments = [seg.strip() for seg in text.split(",")]
    for i, seg in enumerate(segments[1:], start=1):
        if seg and all(is_credential(t) for t in seg.split()):
            text = ", ".join(segments[:i])
            break
    mixed_case = any(c.islower() for c in text)
    tokens: list[str] = []
    for raw in re.split(r"[\s,;/]+", text):
        tok = raw.strip(".,;:()[]\"“”")
        if not tok or is_honorific(tok):
            continue
        # A credential glued to the surname by OCR ("LiMD", "SmithNP").
        glued = re.match(r"^(.+?[a-z])((?:MD|DO|NP|PA|RN|DNP|PhD|DDS|DMD|DPM|DPT)\.?)$", tok)
        if glued and not is_credential(tok):
            tokens.append(glued.group(1))
            break
        # The name ends at the first credential; whatever follows ("MD PhD",
        # "MD Pediatrician", "MD, FACS") is not part of it.
        if is_credential(tok):
            break
        # Fellowship / society designations (FACS, FCAAAI) that were not in
        # the list are all-caps on an otherwise mixed-case line.
        letters = re.sub(r"[^A-Za-z]", "", tok)
        if mixed_case and len(letters) >= 4 and letters.isupper():
            break
        if raw.endswith(".") and len(tok) == 1 and tok.isalpha():
            tok = f"{tok}."                              # keep a real initial marker
        tokens.append(tok)
    # Generation suffixes are not part of the surname; OCR reads "III" as
    # "Ill" / "lll", so any short run of I/l at the end counts too.
    while len(tokens) > 2 and (
        tokens[-1].lower().strip(".") in GENERATION_SUFFIXES
        or re.fullmatch(r"[Il|1]{2,4}", tokens[-1].strip("."))
    ):
        tokens.pop()
    return tokens


def validate_name_tokens(tokens: list[str]) -> tuple[bool, str]:
    """Shape test for a person's name. Returns (ok, reason)."""
    if len(tokens) < 2:
        return False, f"only {len(tokens)} name word(s)"
    if len(tokens) > 5:
        return False, f"{len(tokens)} words is too many for a name"
    real = 0
    for i, tok in enumerate(tokens):
        if not _TOKEN_RE.match(tok):
            return False, f"non-name token {tok!r}"
        if len(tok) > 25:
            return False, f"token too long {tok!r}"
        letters = re.sub(r"[^A-Za-z]", "", tok)
        if len(letters) == 1:
            continue                                     # an initial ("A", "J.")
        low = tok.lower().strip(".'’-")
        if low in NON_NAME_WORDS or (i > 0 and low in FUNCTION_WORDS):
            return False, f"contains non-name word {tok!r}"
        real += 1
    if real < 2:
        return False, "fewer than two full name words"
    last = re.sub(r"[^A-Za-z]", "", tokens[-1])
    if len(last) < 2:
        return False, f"last name {tokens[-1]!r} is an initial"
    return True, "ok"


def split_first_last(tokens: list[str]) -> tuple[str, str]:
    first = tokens[0]
    # A leading initial ("J. Thaddaeus Abbott"): the given name is the next
    # full word when there is one.
    if len(tokens) >= 3 and len(re.sub(r"[^A-Za-z]", "", first)) == 1:
        first = tokens[1]
    return _title(first.rstrip(".")), _title(tokens[-1])


def _title(word: str) -> str:
    # Title-case per hyphen/apostrophe segment, leaving mixed-case names alone
    # ("McDonald", "VonKaenel", "O'Brien").
    if word[:1].isupper() and any(c.islower() for c in word[1:]):
        return word
    return re.sub(r"[A-Za-z][a-z]*", lambda m: m.group(0)[:1].upper() + m.group(0)[1:].lower(), word.lower())


def _key(word: str) -> str:
    return re.sub(r"[^a-z]", "", word.lower())


def _edit_distance(a: str, b: str) -> int:
    if abs(len(a) - len(b)) > 1:
        return 2
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def names_agree(a: list[str], b: list[str]) -> str | None:
    """'full' when the last names and first names/initials match, 'last' when
    only the last names do, None otherwise."""
    if not a or not b:
        return None
    la, lb = _key(a[-1]), _key(b[-1])
    if not la or not lb:
        return None
    last_ok = la == lb or (len(la) >= 4 and len(lb) >= 4 and (la.endswith(lb) or lb.endswith(la)))
    if not last_ok and len(la) >= 4 and len(lb) >= 4 and _edit_distance(la, lb) <= 1:
        last_ok = True                      # one OCR character off ("lacobelli" / "iacobelli")
    if not last_ok:
        # Hyphenated / two-part surnames: any shared part of 4+ letters.
        parts_a = {p for p in re.split(r"[^a-z]+", a[-1].lower()) if len(p) >= 4}
        parts_b = {p for p in re.split(r"[^a-z]+", b[-1].lower()) if len(p) >= 4}
        last_ok = bool(parts_a & parts_b)
    if not last_ok:
        return None
    fa, fb = _key(split_first_last(a)[0]), _key(split_first_last(b)[0])
    if fa and fb and (fa == fb or fa[0] == fb[0]):
        return "full"
    return "last"


def name_from_source_name(source_name: str | None) -> list[str]:
    """Name words from a screenshot/DOCX filename such as
    'Abbie_Smith_abbie-smith-np.png' or 'Alan_Lawrence_Aarons_MD_alan-aarons-md'."""
    if not source_name:
        return []
    stem = Path(str(source_name)).stem
    words: list[str] = []
    slug: list[str] = []
    for part in re.split(r"[_\s]+", stem):
        if re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)+", part):
            slug.extend(part.split("-"))
        elif part:
            words.append(part)

    def clean(parts: list[str]) -> list[str]:
        out = []
        for w in parts:
            w = w.strip(".,()")
            if not w or w.isdigit() or _HEX_RE.match(w.lower()):
                continue
            if is_credential(w) or w.lower() in {"slash", "copy", "screenshot", "image", "resume", "cv"}:
                continue
            if not _TOKEN_RE.match(w):
                continue
            out.append(w)
        return out

    cand = clean(words)
    if len(cand) < 2:
        cand = [_title(w) for w in clean(slug)]
    while len(cand) > 2 and cand[-1].lower() in GENERATION_SUFFIXES:
        cand.pop()
    return cand if len(cand) >= 2 else []


def page_text_names(lines) -> list[str]:
    """Name strings repeated in the page's own copy (join/promo lines)."""
    found: list[str] = []
    for line in lines:
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        for pat in PAGE_TEXT_PATTERNS:
            m = pat.search(text)
            if m:
                cand = m.group(1).strip(" ,.-")
                if cand and cand not in found:
                    found.append(cand)
    return found


@dataclass
class NameCandidate:
    source: str            # layout | page_text | docx_title | text_line | filename | converter
    raw: str
    weight: float
    tokens: list[str] = field(default_factory=list)
    valid: bool = False
    reason: str = ""

    def __post_init__(self):
        if not self.tokens:
            self.tokens = name_tokens(self.raw)
        self.valid, self.reason = validate_name_tokens(self.tokens)


@dataclass
class NameDetection:
    first: str | None
    last: str | None
    display: str | None
    confidence: float
    method: str
    evidence: list[str]
    needs_review: bool
    reason: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict | None) -> "NameDetection | None":
        if not isinstance(data, dict) or "confidence" not in data:
            return None
        try:
            return cls(
                first=data.get("first"),
                last=data.get("last"),
                display=data.get("display"),
                confidence=float(data.get("confidence") or 0.0),
                method=str(data.get("method") or "unknown"),
                evidence=[str(e) for e in (data.get("evidence") or [])],
                needs_review=bool(data.get("needs_review", True)),
                reason=data.get("reason"),
            )
        except (TypeError, ValueError):
            return None

    def summary(self) -> str:
        if not self.display:
            return f"NO NAME (confidence={self.confidence:.2f}; {self.reason})"
        flag = " NEEDS_REVIEW" if self.needs_review else ""
        return (f"{self.display!r} -> first={self.first!r} last={self.last!r} "
                f"confidence={self.confidence:.2f} via {self.method}{flag}")


# Independent sources. Agreement is only counted between different families,
# so two page-text lines saying the same thing do not vouch for each other.
_FAMILY = {
    "layout": "layout", "converter": "layout", "docx_title": "layout",
    "page_text": "page_text", "text_line": "text_line", "filename": "filename",
}


_STRONG_FAMILIES = {"layout", "page_text", "filename"}


def _same_read(a: NameCandidate, b: NameCandidate) -> bool:
    """Two candidates that carry the exact same words are one observation
    (the DOCX title paragraph is also the first text line), not two."""
    same_document = {_FAMILY.get(a.source), _FAMILY.get(b.source)} <= {"layout", "text_line"}
    return same_document and [t.lower() for t in a.tokens] == [t.lower() for t in b.tokens]


def detect_name(candidates: list[NameCandidate]) -> NameDetection:
    """Pick the person's name from every available candidate.

    Each valid candidate starts at its source weight and gains for every other
    source family that agrees with it (+0.2 full, +0.12 last name only) and
    loses 0.15 for each family that clearly disagrees. Invalid candidates never
    win but are kept in the evidence so a failure is explainable.
    """
    evidence: list[str] = []
    valid = [c for c in candidates if c.valid]
    for c in candidates:
        state = "valid" if c.valid else f"rejected: {c.reason}"
        evidence.append(f"{c.source}: {c.raw!r} -> {state}")

    if not valid:
        return NameDetection(
            first=None, last=None, display=None, confidence=0.0, method="none",
            evidence=evidence, needs_review=True,
            reason="no candidate has the shape of a person's name",
        )

    best: tuple[float, NameCandidate, list[str]] | None = None
    for c in valid:
        score = c.weight
        agreeing = {_FAMILY.get(c.source, c.source)}
        seen_family: dict[str, str] = {}
        for other in valid:
            fam = _FAMILY.get(other.source, other.source)
            if other is c or fam in agreeing or _same_read(other, c):
                continue
            verdict = names_agree(c.tokens, other.tokens)
            # One vote per family: the strongest agreement wins for that family.
            prev = seen_family.get(fam)
            if verdict == "full" or (verdict == "last" and prev is None) or (verdict is None and prev is None):
                seen_family[fam] = verdict or "conflict"
        for fam, verdict in seen_family.items():
            if verdict == "full":
                score += 0.2
                agreeing.add(fam)
            elif verdict == "last":
                score += 0.12
                agreeing.add(fam)
            elif fam in _STRONG_FAMILIES:
                # Only a source that itself claims to be the name can veto;
                # a plain text line below the title is expected to differ.
                score -= 0.15
        score = max(0.0, min(0.99, score))
        if best is None or score > best[0]:
            best = (score, c, sorted(agreeing))

    assert best is not None
    score, chosen, agreeing = best
    first, last = split_first_last(chosen.tokens)
    # If the chosen line's first word is not backed by any agreeing source but
    # one of its inner words is ("Zamir Amir Kamel X" beside filename "Amir X"),
    # the inner word is the given name and the first one is an OCR artifact.
    supporters = [
        o for o in valid
        if o is not chosen and not _same_read(o, chosen)
        and _FAMILY.get(o.source) in agreeing and names_agree(chosen.tokens, o.tokens)
    ]
    backed = {_key(split_first_last(o.tokens)[0]): split_first_last(o.tokens)[0] for o in supporters}
    # The filename comes from the profile URL, so when OCR's surname is one
    # character off it (an l/I or rn/m confusion) the filename spelling wins.
    for o in supporters:
        if o.source == "filename" and _key(o.tokens[-1]) != _key(last) and _edit_distance(_key(o.tokens[-1]), _key(last)) == 1:
            evidence.append(f"last name {last!r} is one character off the filename; using {o.tokens[-1]!r}")
            last = _title(o.tokens[-1])
            break
    if backed and _key(first) not in backed:
        inner = next((t for t in chosen.tokens[1:-1] if _key(t) in backed), None)
        if inner:
            evidence.append(f"first name {first!r} not corroborated; using {inner!r} backed by another source")
            first = _title(inner)
        else:
            # One OCR character off a corroborated spelling ("llene" / "Ilene").
            near = next((v for k, v in backed.items() if len(k) >= 4 and _edit_distance(k, _key(first)) <= 1), None)
            if near is None:
                near = next((v for t in chosen.tokens[1:-1] for k, v in backed.items()
                             if len(k) >= 4 and _edit_distance(k, _key(t)) <= 1), None)
            if near:
                evidence.append(f"first name {first!r} is one character off a corroborated {near!r}; using it")
                first = _title(near)
    display = " ".join(chosen.tokens)
    conflicts = [
        f"{o.source} says {' '.join(o.tokens)!r}"
        for o in valid
        if o is not chosen and _FAMILY.get(o.source) not in agreeing
        and _FAMILY.get(o.source) in _STRONG_FAMILIES
        and names_agree(chosen.tokens, o.tokens) is None
    ]
    if conflicts:
        evidence.append("conflict: " + "; ".join(conflicts))
    needs_review = score < REVIEW_BELOW
    reason = None
    if needs_review:
        reason = (f"confidence {score:.2f} below {REVIEW_BELOW}: only {'+'.join(agreeing)} "
                  f"supports {display!r}" + (f"; {conflicts[0]}" if conflicts else ""))
    return NameDetection(
        first=first, last=last, display=display, confidence=round(score, 2),
        method="+".join(agreeing), evidence=evidence, needs_review=needs_review,
        reason=reason,
    )
