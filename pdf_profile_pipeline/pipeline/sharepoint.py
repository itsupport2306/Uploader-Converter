"""Microsoft Graph access for OneDrive / SharePoint shared folders.

Only PDFs are discovered; each PDF is downloaded as bytes and stays a PDF for
the whole pipeline. No conversion of any kind happens here.
"""
from __future__ import annotations

import base64
import json
import os
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urlparse
from urllib.request import Request, urlopen

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
MICROSOFT_LOGIN_BASE_URL = "https://login.microsoftonline.com"
GRAPH_APP_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_DELEGATED_SCOPES = "offline_access Files.Read.All Sites.Read.All User.Read"
PDF_EXTS = {".pdf"}


class GraphError(RuntimeError):
    def __init__(self, message: str, *, status: int | None = None, retry_after: int | None = None):
        super().__init__(message)
        self.status = status
        self.retry_after = retry_after


@dataclass(frozen=True)
class PdfSource:
    """One PDF found in the shared location (or on local disk)."""

    record_id: str
    source_file: str
    display_name: str
    file_name: str
    local_path: Path | None = None
    remote_drive_id: str | None = None
    remote_item_id: str | None = None

    @property
    def is_remote(self) -> bool:
        return bool(self.remote_drive_id and self.remote_item_id)


def _post_form(url: str, payload: dict) -> dict:
    data = urlencode(payload).encode("utf-8")
    req = Request(url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"}, method="POST")
    try:
        with urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8", errors="replace") or "{}")
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        try:
            parsed = json.loads(body)
            message = parsed.get("error_description") or parsed.get("error") or body
        except json.JSONDecodeError:
            message = body
        raise GraphError(str(message).strip(), status=exc.code) from exc
    except URLError as exc:
        raise GraphError(f"Could not reach Microsoft endpoint: {exc}") from exc


class GraphClient:
    """Token holder for either app-only or device-code delegated auth."""

    def __init__(self, *, client_id: str, tenant: str, client_secret: str | None = None, open_browser: bool = True):
        self.client_id = client_id
        self.tenant = (tenant or "common").strip().strip("/") or "common"
        self.client_secret = client_secret or None
        self.open_browser = open_browser
        self.access_token = ""
        self.refresh_token = ""
        self.expires_at = 0.0
        self._lock = threading.Lock()

    @property
    def oauth_base_url(self) -> str:
        return f"{MICROSOFT_LOGIN_BASE_URL}/{quote(self.tenant, safe='')}/oauth2/v2.0"

    def authenticate(self) -> None:
        if self.client_secret:
            self._authenticate_app_only()
        else:
            self._authenticate_device_code()

    def _authenticate_app_only(self) -> None:
        token = _post_form(f"{self.oauth_base_url}/token", {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": GRAPH_APP_SCOPE,
        })
        self._set_token(token)
        print("Microsoft Graph app-only login complete.")

    def _authenticate_device_code(self) -> None:
        start = _post_form(f"{self.oauth_base_url}/devicecode", {
            "client_id": self.client_id,
            "scope": GRAPH_DELEGATED_SCOPES,
        })
        print(start.get("message")
              or f"Open {start.get('verification_uri')} and enter code {start.get('user_code')}")
        if self.open_browser and start.get("verification_uri"):
            try:
                webbrowser.open(start["verification_uri"])
            except Exception:
                pass

        interval = int(start.get("interval", 5))
        deadline = time.time() + int(start.get("expires_in", 900))
        while time.time() < deadline:
            time.sleep(interval)
            try:
                token = _post_form(f"{self.oauth_base_url}/token", {
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "client_id": self.client_id,
                    "device_code": start["device_code"],
                })
            except GraphError as exc:
                message = str(exc).lower()
                if "pending" in message:
                    continue
                if "slow_down" in message:
                    interval += 5
                    continue
                raise
            self._set_token(token)
            print("Microsoft Graph device login complete.")
            return
        raise GraphError("Device login timed out before it was approved.")

    def _set_token(self, token: dict) -> None:
        self.access_token = token["access_token"]
        self.refresh_token = token.get("refresh_token", self.refresh_token)
        self.expires_at = time.time() + int(token.get("expires_in", 3600))

    def _refresh(self) -> bool:
        if not self.refresh_token:
            return False
        token = _post_form(f"{self.oauth_base_url}/token", {
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "refresh_token": self.refresh_token,
            "scope": GRAPH_DELEGATED_SCOPES,
        })
        self._set_token(token)
        return True

    def ensure_token(self) -> str:
        with self._lock:
            if self.access_token and time.time() < self.expires_at - 120:
                return self.access_token
            if not self._refresh():
                self.authenticate()
            return self.access_token


def graph_request(
    client: GraphClient,
    method: str,
    path_or_url: str,
    *,
    headers: dict[str, str] | None = None,
    expect_json: bool = True,
    ok_statuses: set[int] | None = None,
):
    """Graph call with token refresh plus throttle/5xx backoff."""
    ok_statuses = ok_statuses or {200, 201}
    retry_statuses = {429, 500, 502, 503, 504}
    url = path_or_url if path_or_url.startswith(("http://", "https://")) else f"{GRAPH_BASE_URL}{path_or_url}"

    for attempt in range(1, 9):
        request_headers = {"Authorization": f"Bearer {client.ensure_token()}"}
        if headers:
            request_headers.update(headers)
        req = Request(url, headers=request_headers, method=method)
        try:
            with urlopen(req, timeout=120) as resp:
                payload = resp.read()
                status = getattr(resp, "status", resp.getcode())
                if status not in ok_statuses:
                    raise GraphError(f"HTTP {status}", status=status)
                if not expect_json:
                    return payload
                return json.loads(payload.decode("utf-8", errors="replace")) if payload else {}
        except HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            retry_after = None
            header_value = exc.headers.get("Retry-After") if exc.headers else None
            if header_value and str(header_value).isdigit():
                retry_after = int(header_value)
            if exc.code == 401 and attempt < 8:
                client.access_token = ""
                continue
            if exc.code in retry_statuses and attempt < 8:
                wait = retry_after or min(120, 2 ** attempt)
                print(f"Graph is throttling (HTTP {exc.code}); waiting {wait}s before retry {attempt + 1}/8...")
                time.sleep(wait)
                continue
            raise GraphError(f"HTTP {exc.code}: {body[:400]}", status=exc.code, retry_after=retry_after) from exc
        except URLError as exc:
            if attempt < 8:
                time.sleep(min(60, 2 ** attempt))
                continue
            raise GraphError(f"Could not reach Microsoft Graph: {exc}") from exc

    raise GraphError("Microsoft Graph request failed after retries")


def is_remote_input(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def build_client(*, client_id: str | None, tenant: str, client_secret: str | None, open_browser: bool) -> GraphClient:
    if not client_id:
        raise SystemExit(
            "ERROR: SharePoint/OneDrive URL input requires ONEDRIVE_CLIENT_ID in .env "
            "(or --onedrive-client-id)."
        )
    client = GraphClient(
        client_id=client_id,
        tenant=tenant,
        client_secret=client_secret,
        open_browser=open_browser,
    )
    try:
        client.authenticate()
    except GraphError as exc:
        hint = ""
        text = str(exc)
        if "AADSTS50059" in text:
            hint = "\nHint: set ONEDRIVE_TENANT to your tenant ID instead of 'common'."
        elif "AADSTS7000218" in text:
            hint = "\nHint: set ONEDRIVE_CLIENT_SECRET, or enable public client flows on the app registration."
        raise SystemExit(f"Microsoft login failed: {exc}{hint}") from exc
    return client


def _encode_share_url(shared_url: str) -> str:
    encoded = base64.b64encode(shared_url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=").replace("/", "_").replace("+", "-")


def _item_drive_id(item: dict) -> str:
    for source in (item.get("parentReference") or {}, (item.get("remoteItem") or {}).get("parentReference") or {}):
        drive_id = source.get("driveId")
        if drive_id:
            return drive_id
    raise GraphError(f"Missing driveId for item: {item.get('name') or item.get('id')}")


def _resolve_shared_item(client: GraphClient, shared_url: str) -> dict:
    print("Resolving shared SharePoint/OneDrive link...")
    item = graph_request(
        client,
        "GET",
        f"/shares/{_encode_share_url(shared_url)}/driveItem",
        headers={"Prefer": "redeemSharingLink"},
    )
    kind = "folder" if "folder" in item else "file"
    print(f"Resolved shared item: {item.get('name') or item.get('id')} ({kind})")
    return item


def _iter_children(client: GraphClient, drive_id: str, item_id: str):
    next_url = (
        f"{GRAPH_BASE_URL}/drives/{quote(drive_id, safe='')}"
        f"/items/{quote(item_id, safe='')}/children?$top=200"
    )
    while next_url:
        payload = graph_request(client, "GET", next_url)
        for child in payload.get("value", []):
            yield child
        next_url = payload.get("@odata.nextLink")


def _walk(client: GraphClient, item: dict, *, parts: list[str], results: list[PdfSource], limit: int | None) -> None:
    if limit is not None and len(results) >= limit:
        return

    name = item.get("name") or item.get("id") or "item"
    current = parts + [name]

    if "file" in item:
        if Path(name).suffix.lower() in PDF_EXTS:
            display_name = "/".join(current)
            results.append(PdfSource(
                record_id=f"sharepoint://{_item_drive_id(item).lower()}/{str(item['id']).lower()}",
                source_file=item.get("webUrl") or display_name,
                display_name=display_name,
                file_name=Path(name).name,
                remote_drive_id=_item_drive_id(item),
                remote_item_id=str(item["id"]),
            ))
            if len(results) <= 5 or len(results) % 25 == 0:
                print(f"Discovered {len(results)} PDF(s)... latest: {display_name}")
        return

    if "folder" not in item:
        return

    children = sorted(
        _iter_children(client, _item_drive_id(item), str(item["id"])),
        key=lambda child: (child.get("name") or "").lower(),
    )
    for child in children:
        _walk(client, child, parts=current, results=results, limit=limit)
        if limit is not None and len(results) >= limit:
            return


def iter_remote_pdfs(client: GraphClient, shared_url: str, limit: int | None) -> list[PdfSource]:
    root = _resolve_shared_item(client, shared_url)
    print("Listing PDFs from the shared location...")
    results: list[PdfSource] = []
    _walk(client, root, parts=[], results=results, limit=limit)
    if not results:
        raise SystemExit(f"ERROR: no PDF files found in shared location: {shared_url}")
    print(f"Finished listing. Found {len(results)} PDF(s).")
    return results


def iter_local_pdfs(input_path: Path, limit: int | None) -> list[PdfSource]:
    if input_path.is_file():
        if input_path.suffix.lower() not in PDF_EXTS:
            raise SystemExit(f"ERROR: not a PDF: {input_path}")
        files = [input_path]
    elif input_path.is_dir():
        files = sorted(p for p in input_path.rglob("*") if p.is_file() and p.suffix.lower() in PDF_EXTS)
    else:
        raise SystemExit(f"ERROR: input path not found: {input_path}")
    if not files:
        raise SystemExit(f"ERROR: no PDF files found in {input_path}")
    selected = files[:limit] if limit else files
    return [
        PdfSource(
            record_id=str(path.resolve()).lower(),
            source_file=str(path),
            display_name=str(path),
            file_name=path.name,
            local_path=path,
        )
        for path in selected
    ]


def download_pdf_bytes(client: GraphClient, source: PdfSource) -> bytes:
    if not source.is_remote:
        raise ValueError("Remote download requested for a local PDF source.")
    print(f"Downloading PDF: {source.display_name}")
    return graph_request(
        client,
        "GET",
        f"/drives/{quote(source.remote_drive_id or '', safe='')}"
        f"/items/{quote(source.remote_item_id or '', safe='')}/content",
        expect_json=False,
        ok_statuses={200},
    )


def read_pdf_bytes(client: GraphClient | None, source: PdfSource) -> bytes:
    if source.is_remote:
        if client is None:
            raise RuntimeError("A Graph client is required to read remote PDFs.")
        return download_pdf_bytes(client, source)
    if source.local_path is None:
        raise RuntimeError("Local PDF source is missing a path.")
    return source.local_path.read_bytes()


def default_client_id() -> str | None:
    return os.environ.get("ONEDRIVE_CLIENT_ID")
