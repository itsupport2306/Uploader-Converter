"""Re-run the screenshot -> DOCX -> Cloudflare R2 -> Neon flow for profiles whose
first/last name was mis-detected (contains "university", "assistant" or
"certified"). Their stored DOCX only holds the certification section, so every
one is rebuilt from the original PNG.

    python reprocess.py --dry-run --limit 3
    python reprocess.py --workers 4
    python reprocess.py --source-root-url "https://radixsol-my.sharepoint.com/..."

Nothing in the parent folder is modified: this script imports
"complete_process copy.py" (and through it the converter and uploader) and
reuses their functions.

For each matching profile:
  1. read resume_sections.source_file (a OneDrive sync path or a web URL)
  2. fetch the PNG: local disk, --local-root, or Microsoft Graph
  3. OCR + convert to DOCX in memory (same converter/args as complete_process)
  4. upload the DOCX to R2, overwriting any object at the new key
  5. UPDATE the same profile_id with the re-parsed fields and new file info
  6. delete the old R2 object, if its key differs and no other row uses it
"""
from __future__ import annotations

import argparse
import importlib.util
import io
import json
import os
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from time import monotonic
from urllib.parse import quote, urlparse

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
PROCESS_SCRIPT = ROOT / "complete_process copy.py"
DEFAULT_ENV_FILE = ROOT / "word_profile_pipeline" / ".env"
DEFAULT_MANIFEST = HERE / "logs" / "reprocess_manifest.jsonl"
DEFAULT_RUN_LOG_DIR = HERE / "logs"
# Folder name of the OneDrive sync root inside a stored Windows path, e.g.
#   C:\Users\admin\OneDrive - Radixsol\Physical_A\...\file.png
# Everything after it is the path inside the drive.
DEFAULT_SYNC_ROOT_NAME = "OneDrive - Radixsol"

NAME_TERMS = ("university", "assistant", "certified")
SELECT_SQL = """
    SELECT profile_id, first_name, last_name, resume_url, resume_sections::text
    FROM {table}
    WHERE first_name ILIKE '%university%' OR last_name ILIKE '%university%'
       OR first_name ILIKE '%assistant%'  OR last_name ILIKE '%assistant%'
       OR first_name ILIKE '%certified%'  OR last_name ILIKE '%certified%'
    ORDER BY updated_at NULLS LAST, profile_id
"""


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load module from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


# Loads the converter and "data_upload 5.py" as a side effect.
cp = _load_module("complete_process_copy", PROCESS_SCRIPT)
converter = cp.converter
uploader = cp.uploader
DOCX_CT = cp.DOCX_CONTENT_TYPE


@dataclass(frozen=True)
class BadProfile:
    profile_id: str
    first_name: str | None
    last_name: str | None
    resume_url: str | None
    sections: dict

    @property
    def source_file(self) -> str:
        return str(self.sections.get("source_file") or "").strip()

    @property
    def old_key(self) -> str | None:
        key = self.sections.get("cloudflare_key")
        if key:
            return str(key)
        # Older rows may only carry resume_url = /files/<key> or <public>/<key>.
        url = self.resume_url or ""
        marker = "/resumes/"
        if marker in url:
            return url[url.index(marker) + 1:]
        return None

    @property
    def label(self) -> str:
        return f"{self.first_name or ''} {self.last_name or ''}".strip() or self.profile_id


class Tee:
    """Mirror stdout/stderr into a run log so worker output is kept."""

    def __init__(self, stream, path: Path):
        self.stream = stream
        self.fh = path.open("a", encoding="utf-8")
        self.lock = threading.Lock()

    def write(self, data):
        with self.lock:
            self.stream.write(data)
            self.fh.write(data)

    def flush(self):
        self.stream.flush()
        self.fh.flush()


def _utc_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _append_jsonl(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, sort_keys=True, default=str) + "\n")


def _done_ids(path: Path) -> set[str]:
    done: set[str] = set()
    if not path.exists():
        return done
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("status") == "processed" and rec.get("profile_id"):
                done.add(rec["profile_id"])
    return done


# ---------------------------------------------------------------------------
# Selecting the profiles

def _fetch_bad_profiles(engine, table: str, limit: int | None) -> list[BadProfile]:
    sql = SELECT_SQL.format(table=cp._quoted_table(table))
    with engine.connect() as conn:
        rows = conn.execute(uploader._text(sql)).all()
    profiles: list[BadProfile] = []
    for profile_id, first, last, resume_url, sections_text in rows:
        try:
            sections = json.loads(sections_text) if sections_text else {}
        except json.JSONDecodeError:
            sections = {}
        if not isinstance(sections, dict):
            sections = {}
        profiles.append(BadProfile(str(profile_id), first, last, resume_url, sections))
    return profiles[:limit] if limit else profiles


# ---------------------------------------------------------------------------
# Fetching the source PNG

class SourceFetcher:
    """Resolves a stored source_file to image bytes.

    Order: the path as-is on this machine -> --local-root remap ->
    Microsoft Graph (web URL via /shares, or sync path under --source-root-url
    or --source-user).
    """

    def __init__(self, args: argparse.Namespace):
        self.args = args
        self.client = None
        self.client_lock = threading.Lock()
        self.root_item: dict | None = None
        self.local_root = Path(args.local_root).expanduser() if args.local_root else None
        self.sync_root_name = args.sync_root_name.strip().lower()

    # -- Graph --------------------------------------------------------------
    def _graph(self):
        with self.client_lock:
            if self.client is None:
                self.client = cp._build_onedrive_client(self.args)
                if self.args.source_root_url:
                    self.root_item = cp._resolve_shared_item(self.client, self.args.source_root_url)
            return self.client

    def _relative_in_drive(self, source: str) -> str | None:
        """C:/.../OneDrive - Radixsol/A/b.png -> A/b.png"""
        parts = [p for p in source.replace("\\", "/").split("/") if p]
        for index, part in enumerate(parts):
            if part.strip().lower() == self.sync_root_name:
                rest = parts[index + 1:]
                return "/".join(rest) if rest else None
        return None

    def _download_by_share_url(self, url: str) -> bytes:
        client = self._graph()
        item = cp._resolve_shared_item(client, url)
        return cp._onedrive_request(
            client, "GET",
            f"/drives/{quote(cp._item_drive_id(item), safe='')}/items/{quote(str(item['id']), safe='')}/content",
            expect_json=False, ok_statuses={200},
        )

    def _download_by_drive_path(self, rel_path: str) -> bytes:
        if not self.args.source_root_url and not self.args.source_user:
            raise RuntimeError(
                "source image is not on this machine; pass --source-root-url (share link of the "
                f"'{self.args.sync_root_name}' folder) or --source-user (owner's UPN) to fetch it from Graph"
            )
        client = self._graph()
        encoded = quote(rel_path, safe="/")
        if self.root_item is not None:
            drive_id = cp._item_drive_id(self.root_item)
            path = (f"/drives/{quote(drive_id, safe='')}/items/{quote(str(self.root_item['id']), safe='')}"
                    f":/{encoded}:/content")
        else:
            path = f"/users/{quote(self.args.source_user, safe='')}/drive/root:/{encoded}:/content"
        return cp._onedrive_request(client, "GET", path, expect_json=False, ok_statuses={200})

    # -- public -------------------------------------------------------------
    def fetch(self, source: str) -> tuple[bytes, str]:
        """Return (image_bytes, how)."""
        if not source:
            raise RuntimeError("resume_sections.source_file is empty")
        parsed = urlparse(source)
        if parsed.scheme in {"http", "https"} and parsed.netloc:
            return self._download_by_share_url(source), "graph-url"

        local = Path(source)
        if local.is_file():
            return local.read_bytes(), "local"

        rel = self._relative_in_drive(source)
        if self.local_root is not None and rel:
            candidate = self.local_root / Path(rel)
            if candidate.is_file():
                return candidate.read_bytes(), "local-root"

        if rel is None:
            raise RuntimeError(
                f"could not find '{self.args.sync_root_name}' in source path and file is not local: {source}"
            )
        return self._download_by_drive_path(rel), "graph-path"


# ---------------------------------------------------------------------------
# R2 helpers (force overwrite; the uploader's helpers skip existing keys)

def _put_docx(data: bytes, key: str) -> str:
    extra = {"ContentType": DOCX_CT}
    acl = os.environ.get("S3_ACL", "").strip()
    if acl:
        extra["ACL"] = acl
    uploader._s3_client().upload_fileobj(io.BytesIO(data), os.environ["S3_BUCKET"], key, ExtraArgs=extra)
    return uploader._url_for_key(key)


def _delete_key(key: str) -> bool:
    if not uploader._object_exists(key):
        return False
    uploader._s3_client().delete_object(Bucket=os.environ["S3_BUCKET"], Key=key)
    return True


def _key_still_referenced(conn, table: str, key: str) -> int:
    row = conn.execute(uploader._text(f"""
        SELECT count(*) FROM {cp._quoted_table(table)}
        WHERE resume_sections::text LIKE :like OR resume_url LIKE :like
    """), {"like": f"%{key}%"}).scalar()
    return int(row or 0)


# ---------------------------------------------------------------------------
# One profile

def _reprocess_one(profile: BadProfile, fetcher: SourceFetcher, engine, args) -> dict:
    base = {
        "profile_id": profile.profile_id,
        "old_first_name": profile.first_name,
        "old_last_name": profile.last_name,
        "old_key": profile.old_key,
        "source_file": profile.source_file,
        "updated_at": _utc_iso(),
    }
    source = profile.source_file
    print(f"Fetching source image: {profile.label} <- {source}")
    image_bytes, how = fetcher.fetch(source)
    base["fetched_via"] = how

    # Same digest rule as complete_process: hash the source image so the key is
    # stable across reconversions.
    digest = uploader._sha256_bytes(image_bytes)
    display_name = Path(source.replace("\\", "/")).name or f"{profile.profile_id}.png"
    print(f"Running OCR and DOCX conversion: {display_name}")
    docx_bytes, _ = converter.convert_bytes_result(
        image_bytes, display_name,
        scale=args.scale, cutoff_ratio=args.cutoff_ratio,
        keep_promo=args.keep_promo, min_conf=args.min_conf, debug=args.debug,
    )
    docx_profile = uploader.extract_docx_profile_bytes(docx_bytes)
    text = docx_profile.text
    docx_parse_path = Path(display_name).with_suffix(".docx")
    fields = uploader.parse_resume_smart(text, docx_parse_path, docx_profile)

    print(
        f"NAME {display_name}: {fields.get('name_display')!r} -> "
        f"first={fields['first_name']!r} last={fields['last_name']!r} "
        f"confidence={fields.get('name_confidence', 0):.2f} via {fields.get('name_method')}"
        + (f" | REVIEW: {fields.get('review_reason')}" if fields.get("needs_review") else "")
    )
    new_key = cp._cloudflare_key_for(
        docx_parse_path, digest,
        fields.get("provider_category"), fields.get("profession_type"), fields.get("specialty"),
    )
    base.update({
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "specialty": fields.get("specialty"),
        "needs_review": bool(fields.get("needs_review")),
        "review_reason": fields.get("review_reason"),
        "docx_sha256": digest,
        "new_key": new_key,
        "docx_bytes": len(docx_bytes),
    })

    if args.save_docx:
        out = Path(args.save_docx).expanduser()
        out.mkdir(parents=True, exist_ok=True)
        (out / f"{docx_parse_path.stem}_{profile.profile_id[:8]}.docx").write_bytes(docx_bytes)

    if args.dry_run:
        return {**base, "status": "would_process"}

    print(f"Uploading DOCX to Cloudflare R2: {new_key}")
    resume_url = _put_docx(docx_bytes, new_key)

    print(f"Updating profile {profile.profile_id} in Neon")
    with engine.begin() as conn:
        cp._upsert_candidate(
            conn,
            table_name=args.target_table,
            row_id=profile.profile_id,
            fields=fields,
            text=text,
            source_file=source,
            docx_file=f"memory://{docx_parse_path.as_posix()}",
            digest=digest,
            key=new_key,
            resume_url=resume_url,
        )

    old_key = profile.old_key
    deleted = False
    if old_key and old_key != new_key and not args.keep_old:
        with engine.connect() as conn:
            refs = _key_still_referenced(conn, args.target_table, old_key)
        if refs == 0:
            deleted = _delete_key(old_key)
            print(f"Deleted old R2 object: {old_key}" if deleted else f"Old R2 object already gone: {old_key}")
        else:
            print(f"Kept old R2 object (still referenced by {refs} row(s)): {old_key}")
    elif old_key == new_key:
        print("Old and new keys match; the bad DOCX was overwritten in place.")

    return {
        **base,
        "status": "processed",
        "resume_url": resume_url,
        "old_key_deleted": deleted,
        "uploaded_at": _utc_iso(),
    }


# ---------------------------------------------------------------------------

def run(args: argparse.Namespace) -> dict:
    uploader._prepare_runtime_env(args)
    uploader._validate_upload_env(args)

    engine = uploader._create_engine()
    cp._ensure_target_table(engine, args.target_table)
    profiles = _fetch_bad_profiles(engine, args.target_table, args.limit)
    done = set() if args.retry_all else _done_ids(args.manifest)
    todo = [p for p in profiles if p.profile_id not in done]
    print(f"Matched {len(profiles)} profile(s) with {'/'.join(NAME_TERMS)} in the name; "
          f"{len(todo)} to process ({len(profiles) - len(todo)} already done in manifest).")
    if not todo:
        return {"processed": 0, "failed": 0, "total": 0}
    if args.dry_run:
        print("Dry run: fetch + convert + parse only; no upload, no DB write, no delete.")

    converter.configure_tesseract(args.tesseract)
    fetcher = SourceFetcher(args)
    stats = {"processed": 0, "needs_review": 0, "failed": 0, "total": len(todo)}
    started = monotonic()
    workers = max(1, min(args.workers, len(todo)))
    print(f"Starting with {workers} worker(s)...")

    def _ok(index: int, profile: BadProfile, record: dict) -> None:
        _append_jsonl(args.manifest, record)
        stats["processed"] += 1
        if record.get("needs_review"):
            stats["needs_review"] += 1
        print(f"[{index}/{len(todo)}] {record['status'].upper()} {profile.label} -> "
              f"{record.get('first_name')} {record.get('last_name')} | {record.get('new_key')}")

    def _fail(index: int, profile: BadProfile, exc: Exception) -> None:
        stats["failed"] += 1
        _append_jsonl(args.manifest, {
            "profile_id": profile.profile_id, "source_file": profile.source_file,
            "old_key": profile.old_key, "status": "failed", "error": str(exc),
            "updated_at": _utc_iso(),
        })
        print(f"[{index}/{len(todo)}] FAIL {profile.label} ({profile.profile_id}): {exc}", file=sys.stderr)

    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="reprocess") as pool:
        futures = {
            pool.submit(_reprocess_one, profile, fetcher, engine, args): (i, profile)
            for i, profile in enumerate(todo, start=1)
        }
        for future in as_completed(futures):
            index, profile = futures[future]
            try:
                _ok(index, profile, future.result())
            except Exception as exc:  # keep the other workers going
                _fail(index, profile, exc)

    elapsed = max(monotonic() - started, 0.001)
    print(f"\nSummary: {stats}")
    if stats["needs_review"]:
        print(f"{stats['needs_review']} profile(s) flagged needs_review in the manifest - check their names.")
    print(f"Elapsed: {elapsed / 60:.1f} min")
    return stats


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--env-file", default=str(DEFAULT_ENV_FILE), help=f"default: {DEFAULT_ENV_FILE}")
    parser.add_argument("--workers", type=int, default=None, help="Parallel workers (default: PIPELINE_WORKERS from .env, else 2)")
    parser.add_argument("--limit", type=int, default=None, help="Only handle the first N matching profiles")
    parser.add_argument("--dry-run", action="store_true", help="Fetch/convert/parse only; no upload, DB write or delete")
    parser.add_argument("--save-docx", default=None, help="Also write each generated DOCX into this folder (useful with --dry-run)")
    parser.add_argument("--keep-old", action="store_true", help="Do not delete the old R2 object")
    parser.add_argument("--retry-all", action="store_true", help="Ignore the manifest and redo every match")
    parser.add_argument("--manifest", default=str(DEFAULT_MANIFEST))
    parser.add_argument("--no-run-log", action="store_true", help="Do not mirror console output to logs/run_<ts>.log")
    parser.add_argument("--target-table", default=None, help='default: TARGET_TABLE from .env, else "profiles"')

    src = parser.add_argument_group("source image location")
    src.add_argument("--sync-root-name", default=DEFAULT_SYNC_ROOT_NAME,
                     help="Folder name in stored paths that marks the OneDrive root")
    src.add_argument("--local-root", default=None,
                     help="Local folder that mirrors the OneDrive root (or REPROCESS_LOCAL_ROOT in .env)")
    src.add_argument("--source-root-url", default=None,
                     help="SharePoint/OneDrive share link of the root folder (or REPROCESS_SOURCE_ROOT_URL)")
    src.add_argument("--source-user", default=None,
                     help="Owner UPN of the OneDrive, e.g. admin@radixsol.com (or REPROCESS_SOURCE_USER)")
    src.add_argument("--onedrive-client-id", default=None)
    src.add_argument("--onedrive-client-secret", default=None)
    src.add_argument("--onedrive-tenant", default=None)
    src.add_argument("--no-browser", action="store_true")

    conv = parser.add_argument_group("converter (same defaults as complete_process)")
    conv.add_argument("--tesseract", help="Path to tesseract.exe")
    conv.add_argument("--scale", type=float, default=None)
    conv.add_argument("--cutoff-ratio", type=float, default=None)
    conv.add_argument("--min-conf", type=int, default=35)
    conv.add_argument("--keep-promo", action="store_true")
    conv.add_argument("--debug", action="store_true")
    parser.add_argument("--install-deps", action="store_true")
    args = parser.parse_args(argv)

    args.manifest = Path(args.manifest).expanduser()
    if not Path(args.env_file).expanduser().is_absolute():
        args.env_file = str((HERE / args.env_file).resolve())

    # Load the env first so the remaining defaults come from the same .env.
    uploader._load_env(args.env_file)
    env = os.environ.get
    args.workers = args.workers or int(env("PIPELINE_WORKERS", "2") or 2)
    args.target_table = args.target_table or env("TARGET_TABLE", "profiles")
    args.onedrive_client_id = args.onedrive_client_id or env("ONEDRIVE_CLIENT_ID")
    args.onedrive_client_secret = args.onedrive_client_secret or env("ONEDRIVE_CLIENT_SECRET")
    args.onedrive_tenant = args.onedrive_tenant or env("ONEDRIVE_TENANT", "common")
    args.source_root_url = args.source_root_url or env("REPROCESS_SOURCE_ROOT_URL")
    args.source_user = args.source_user or env("REPROCESS_SOURCE_USER")
    args.local_root = args.local_root or env("REPROCESS_LOCAL_ROOT")
    if args.workers < 1:
        raise SystemExit("ERROR: --workers must be 1 or greater.")

    if args.install_deps:
        uploader._install_missing_dependencies()

    if not args.no_run_log:
        DEFAULT_RUN_LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = DEFAULT_RUN_LOG_DIR / f"run_{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S')}.log"
        sys.stdout = Tee(sys.stdout, log_path)
        sys.stderr = Tee(sys.stderr, log_path)
        print(f"Run log: {log_path}")

    stats = run(args)
    return 1 if stats["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
