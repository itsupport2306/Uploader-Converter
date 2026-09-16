# reprocess_bad_names

Rebuilds the DOCX for profiles whose `first_name`/`last_name` contains
`university`, `assistant` or `certified`. Their current Cloudflare DOCX only
holds the certification section, so each one is regenerated from the original
PNG and the profile row is updated in place.

Reuses `../complete_process copy.py` (converter + uploader) without touching it,
and reads `../word_profile_pipeline/.env` by default.

## Flow per profile

1. `SELECT ... FROM profiles WHERE first_name/last_name ILIKE %university|assistant|certified%`
2. read `resume_sections.source_file`
3. fetch the PNG: local path -> `--local-root` remap -> Microsoft Graph
4. OCR + convert to DOCX (same converter and defaults as `complete_process copy.py`)
5. upload to R2 — **overwrites** the key if it already exists
6. `UPDATE profiles ... WHERE profile_id = <same id>` (name, resume_url, resume_sections, etc.)
7. delete the old R2 object if the key changed and no other row references it

## Run

```powershell
cd D:\Radixsol\reprocess_bad_names
python reprocess.py --dry-run --limit 3 --save-docx .\preview   # check a few first
python reprocess.py --workers 4
```

## Where the PNGs come from

`source_file` is stored as a OneDrive sync path, e.g.
`C:\Users\admin\OneDrive - Radixsol\Physical_A\Family_Medicine\X\x.png`.
On a machine where that path exists it is read directly. Otherwise:

| option | use when |
|---|---|
| `--local-root D:\OneDriveMirror` | the same tree is synced somewhere else on this PC |
| `--source-root-url <share link>` | you have a SharePoint/OneDrive share link for the `OneDrive - Radixsol` folder |
| `--source-user admin@radixsol.com` | app-only Graph access to that user's OneDrive (`/users/<upn>/drive/root:/...`) |

These can also be set in `.env` as `REPROCESS_LOCAL_ROOT`, `REPROCESS_SOURCE_ROOT_URL`,
`REPROCESS_SOURCE_USER`. Graph uses `ONEDRIVE_CLIENT_ID / ONEDRIVE_CLIENT_SECRET / ONEDRIVE_TENANT`
from the same `.env`. If `source_file` is an `https://` URL it is resolved via the Graph `/shares` API.

## Logs

- `logs/reprocess_manifest.jsonl` — one line per profile (`processed` / `would_process` / `failed`,
  old and new key, detected name, `needs_review`). Re-running skips `processed` rows; `--retry-all` redoes them.
- `logs/run_<timestamp>.log` — console transcript.

`--keep-old` skips the delete step. Rows still flagged `needs_review` are updated anyway
(the existing names are known-bad); review them from the manifest.
