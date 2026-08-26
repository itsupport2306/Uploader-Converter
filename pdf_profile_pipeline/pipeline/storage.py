"""Upload the original PDF to Cloudflare R2, byte for byte."""
from __future__ import annotations

import hashlib
import io
import os
import re
import threading

from . import config

PDF_CONTENT_TYPE = "application/pdf"

_client = None
_client_lock = threading.Lock()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _import(name: str):
    try:
        module = __import__(name, fromlist=["*"])
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise SystemExit(f"ERROR: {name} is required. Run: pip install -r requirements.txt") from exc
    return module


def client():
    global _client
    with _client_lock:
        if _client is None:
            boto3 = _import("boto3")
            botocore_config = _import("botocore.config")
            _client = boto3.client(
                "s3",
                endpoint_url=os.environ.get("S3_ENDPOINT_URL") or None,
                region_name=os.environ.get("S3_REGION", "auto"),
                aws_access_key_id=os.environ.get("S3_ACCESS_KEY"),
                aws_secret_access_key=os.environ.get("S3_SECRET_KEY"),
                config=botocore_config.Config(signature_version="s3v4"),
            )
        return _client


def safe_segment(value: str | None) -> str:
    text = re.sub(r"[^A-Za-z0-9]+", "_", value or "Unknown")
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:80] or "Unknown"


def key_for(digest: str, *segments: str | None) -> str:
    """Content-addressed key, so re-processing the same PDF reuses the object."""
    prefix = (os.environ.get("R2_KEY_PREFIX") or "resumes").strip("/")
    folder = safe_segment(next((value for value in segments if value), None))
    return f"{prefix}/{folder}/{digest}.pdf"


def url_for_key(key: str) -> str:
    public_base = (os.environ.get("S3_PUBLIC_BASE_URL") or "").rstrip("/")
    if config.env_bool("S3_PUBLIC") and public_base:
        return f"{public_base}/{key}"
    return f"/files/{key}"


def object_exists(key: str) -> bool:
    exceptions = _import("botocore.exceptions")
    try:
        client().head_object(Bucket=os.environ["S3_BUCKET"], Key=key)
        return True
    except exceptions.ClientError as exc:
        code = str(exc.response.get("Error", {}).get("Code", ""))
        if code in {"404", "NoSuchKey", "NotFound"}:
            return False
        raise


def upload_pdf(data: bytes, key: str) -> tuple[str, bool]:
    """Upload the PDF unless the identical object is already there.

    Returns (url, uploaded_now).
    """
    if object_exists(key):
        return url_for_key(key), False

    extra = {"ContentType": PDF_CONTENT_TYPE}
    acl = (os.environ.get("S3_ACL") or "").strip()
    if acl:
        extra["ACL"] = acl

    client().upload_fileobj(io.BytesIO(data), os.environ["S3_BUCKET"], key, ExtraArgs=extra)
    return url_for_key(key), True
