"""Delete profiles whose resume_sections mentions "Physical_A", along with
their resume object in Cloudflare R2.

    python delete_physical_a.py --dry-run
    python delete_physical_a.py --dry-run --limit 5
    python delete_physical_a.py

Reads ../word_profile_pipeline/.env by default (same DATABASE_URL / S3_*
settings as the rest of the pipeline). Does not import or modify any other
script in the repo.

For each matching profile:
  1. SELECT profile_id, resume_sections FROM profiles
     WHERE resume_sections::text ILIKE '%Physical_A%'
  2. read resume_sections.cloudflare_bucket / cloudflare_key
  3. delete that object from Cloudflare R2 (S3-compatible API)
  4. only if step 3 did not raise -> DELETE FROM profiles WHERE profile_id = ...
  5. append a record to logs/delete_manifest.jsonl either way

Re-running skips profile_ids already marked "deleted" in the manifest unless
--retry-all is passed.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DEFAULT_ENV_FILE = ROOT / "word_profile_pipeline" / ".env"
DEFAULT_MANIFEST = HERE / "logs" / "delete_manifest.jsonl"

MATCH_TERM = "Physical_A"
TABLE = "profiles"
SELECT_SQL = f"""
    SELECT profile_id, resume_sections::text
    FROM {TABLE}
    WHERE resume_sections::text ILIKE :pattern
    ORDER BY profile_id
"""


def _load_env(path: Path) -> None:
    if not path.exists():
        return
    with open(path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, value)


def _create_engine():
    import sqlalchemy
    db_url = os.environ.get("DATABASE_URL", "")
    if not db_url:
        raise SystemExit("ERROR: DATABASE_URL is missing from the env file.")
    return sqlalchemy.create_engine(db_url, future=True, pool_pre_ping=True)


_S3_CLIENT = None


def _s3_client():
    global _S3_CLIENT
    if _S3_CLIENT is None:
        import boto3
        from botocore.config import Config
        _S3_CLIENT = boto3.client(
            "s3",
            endpoint_url=os.environ.get("S3_ENDPOINT_URL") or None,
            region_name=os.environ.get("S3_REGION", "auto"),
            aws_access_key_id=os.environ.get("S3_ACCESS_KEY"),
            aws_secret_access_key=os.environ.get("S3_SECRET_KEY"),
            config=Config(signature_version="s3v4"),
        )
    return _S3_CLIENT


def _object_exists(bucket: str, key: str) -> bool:
    from botocore.exceptions import ClientError
    try:
        _s3_client().head_object(Bucket=bucket, Key=key)
        return True
    except ClientError as exc:
        code = str(exc.response.get("Error", {}).get("Code", ""))
        if code in {"404", "NoSuchKey", "NotFound"}:
            return False
        raise


def _delete_object(bucket: str, key: str) -> bool:
    """Delete the R2 object. Returns True if an object was actually removed,
    False if it was already gone. Raises on any other failure."""
    existed = _object_exists(bucket, key)
    if not existed:
        return False
    _s3_client().delete_object(Bucket=bucket, Key=key)
    return True


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
            if rec.get("status") == "deleted" and rec.get("profile_id"):
                done.add(rec["profile_id"])
    return done


def _fetch_matches(engine, limit: int | None) -> list[dict]:
    from sqlalchemy import text
    with engine.connect() as conn:
        rows = conn.execute(text(SELECT_SQL), {"pattern": f"%{MATCH_TERM}%"}).all()
    matches = []
    for profile_id, sections_text in rows:
        try:
            sections = json.loads(sections_text) if sections_text else {}
        except json.JSONDecodeError:
            sections = {}
        if not isinstance(sections, dict):
            sections = {}
        matches.append({"profile_id": str(profile_id), "sections": sections})
    return matches[:limit] if limit else matches


def _process_one(engine, args, profile_id: str, sections: dict) -> dict:
    bucket = str(sections.get("cloudflare_bucket") or os.environ.get("S3_BUCKET") or "").strip()
    key = str(sections.get("cloudflare_key") or "").strip()
    base = {"profile_id": profile_id, "bucket": bucket, "key": key, "updated_at": _utc_iso()}

    if not key:
        return {**base, "status": "skipped", "reason": "no cloudflare_key in resume_sections"}
    if not bucket:
        return {**base, "status": "skipped", "reason": "no cloudflare_bucket and no S3_BUCKET fallback"}

    if args.dry_run:
        exists = _object_exists(bucket, key)
        return {**base, "status": "would_delete", "r2_object_exists": exists}

    from sqlalchemy import text
    try:
        r2_deleted = _delete_object(bucket, key)
    except Exception as exc:
        return {**base, "status": "failed", "stage": "r2_delete", "error": str(exc)}

    try:
        with engine.begin() as conn:
            result = conn.execute(
                text(f"DELETE FROM {TABLE} WHERE profile_id = :pid"),
                {"pid": profile_id},
            )
        db_deleted = result.rowcount > 0
    except Exception as exc:
        return {
            **base, "status": "failed", "stage": "db_delete", "error": str(exc),
            "r2_object_deleted": r2_deleted,
        }

    return {
        **base,
        "status": "deleted",
        "r2_object_deleted": r2_deleted,
        "db_row_deleted": db_deleted,
    }


def run(args: argparse.Namespace) -> dict:
    engine = _create_engine()
    matches = _fetch_matches(engine, args.limit)
    done = set() if args.retry_all else _done_ids(args.manifest)
    todo = [m for m in matches if m["profile_id"] not in done]
    print(f"Matched {len(matches)} profile(s) containing '{MATCH_TERM}' in resume_sections; "
          f"{len(todo)} to process ({len(matches) - len(todo)} already deleted per manifest).")
    if not todo:
        return {"deleted": 0, "failed": 0, "skipped": 0, "total": 0}
    if args.dry_run:
        print("Dry run: no R2 deletion, no DB deletion. Reporting only.")

    stats = {"deleted": 0, "failed": 0, "skipped": 0, "would_delete": 0, "total": len(todo)}
    for i, m in enumerate(todo, start=1):
        record = _process_one(engine, args, m["profile_id"], m["sections"])
        _append_jsonl(args.manifest, record)
        status = record["status"]
        stats[status if status in stats else "failed"] = stats.get(status, 0) + 1
        print(f"[{i}/{len(todo)}] {status.upper()} profile_id={m['profile_id']} "
              f"bucket={record.get('bucket')} key={record.get('key')}"
              + (f" ERROR={record.get('error')}" if record.get("error") else ""))

    print(f"\nSummary: {stats}")
    return stats


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--env-file", default=str(DEFAULT_ENV_FILE), help=f"default: {DEFAULT_ENV_FILE}")
    parser.add_argument("--dry-run", action="store_true", help="Report matches only; delete nothing")
    parser.add_argument("--limit", type=int, default=None, help="Only handle the first N matching profiles")
    parser.add_argument("--retry-all", action="store_true", help="Ignore the manifest and redo every match")
    parser.add_argument("--manifest", default=str(DEFAULT_MANIFEST))
    args = parser.parse_args(argv)

    args.manifest = Path(args.manifest).expanduser()
    _load_env(Path(args.env_file).expanduser())

    stats = run(args)
    return 1 if stats.get("failed") else 0


if __name__ == "__main__":
    raise SystemExit(main())
