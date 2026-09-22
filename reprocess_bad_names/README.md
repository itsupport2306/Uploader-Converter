# reprocess_bad_names

Rebuilds the DOCX for profiles whose `first_name`/`last_name` looks wrong.
Candidates are found by `qwen_detect.py`: it shows Qwen each profile's stored
name/headline/specialty plus an excerpt of its OCR'd resume text
(`resume_sections.raw_extracted_text`) and asks it to flag names that look
incorrect or corrupted. That catches junk a fixed word list misses — a
plausible-looking two-word "name" that is really an organization, a class,
a place, or simply doesn't match the person named in the resume ("Class Of" /
"Samaritan Vie" / "Santa Hospital", or a first/last name that disagrees with
the resume text) — not just names containing `university`, `assistant` or
`certified`. Their current Cloudflare DOCX often only holds the certification
section, so each flagged profile is regenerated from the original PNG and the
profile row is updated in place.

Reuses `../complete_process copy.py` (converter + uploader) without touching it,
and reads `../word_profile_pipeline/.env` by default — including the
`LLM_BASE_URL` / `LLM_MODEL` / `LLM_API_KEY` settings already there for the
Qwen2.5 server, which the detection pass talks to independently of that
pipeline's code.

## Flow

1. **Detect** (`qwen_detect.py`): scan every profile (or `--scan-limit` of
   them), a batch at a time, asking Qwen `{"bad_ids": [...]}` for the ones in
   each batch whose name looks wrong. Results are cached in
   `logs/qwen_scan_manifest.jsonl` (`verdict: "bad"|"ok"`), so a re-run only
   classifies profiles added since the last scan (`--rescan-all` to force a
   full reclassify). By default the legacy `university`/`assistant`/
   `certified` substring match is unioned in too (`--no-term-match` to drop
   it, `--term-match-only` to use only the legacy match and skip Qwen).
2. **Reprocess**, per flagged profile (unchanged from before):
   1. read `resume_sections.source_file`
   2. fetch the PNG: local path -> `--local-root` remap -> Microsoft Graph
   3. OCR + convert to DOCX (same converter and defaults as `complete_process copy.py`)
   4. upload to R2 — **overwrites** the key if it already exists
   5. `UPDATE profiles ... WHERE profile_id = <same id>` (name, resume_url, resume_sections, etc.)
   6. delete the old R2 object if the key changed and no other row references it

## Run

```powershell
cd D:\Radixsol\reprocess_bad_names
python reprocess.py --scan-limit 20 --detect-only          # preview what Qwen flags on a small sample first
python reprocess.py --dry-run --limit 3 --save-docx .\preview   # then check a few reprocessed ones
python reprocess.py --workers 4
```

## Splitting the work across servers

Each row's `created_at` (e.g. `2026-08-26 15:44:34.315109`) can partition the
table so several servers each own a disjoint slice and never touch the same
profile. Set a range in `.env`:

```env
START_DATE=2026-08-01 00:00:00
END_DATE=2026-08-15 23:59:59
```

or pass `--start-date` / `--end-date` (they override `.env`; both or neither
must be given). The range gates every query this script makes — the Qwen
scan, the legacy term match, and therefore the reprocessing list — so a
server given `2026-08-01..2026-08-15` never sees a row outside it. Give each
server its own non-overlapping range (and its own `.env`/`logs/` folder,
since the scan and reprocess manifests are local caches) and run them at the
same time:

```powershell
# server 1
python reprocess.py --start-date "2026-08-01 00:00:00" --end-date "2026-08-15 23:59:59" --workers 4
# server 2 (different machine or checkout)
python reprocess.py --start-date "2026-08-16 00:00:00" --end-date "2026-08-31 23:59:59" --workers 4
```

Within one server, `--workers` (reprocessing) and `--scan-concurrency` (the
Qwen scan) already bound how many run at once — the profiles are queued and
pulled off by that many worker threads, so `--workers 1` processes them
strictly one at a time.

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

- `logs/qwen_scan_manifest.jsonl` — one line per profile classified by Qwen (`verdict: "bad"|"ok"`,
  `reason`, `scanned_at`). Re-running only classifies profile_ids not yet in this file; `--rescan-all` redoes them.
- `logs/reprocess_manifest.jsonl` — one line per profile (`processed` / `would_process` / `failed`,
  old and new key, detected name, `needs_review`). Re-running skips `processed` rows; `--retry-all` redoes them.
- `logs/run_<timestamp>.log` — console transcript.

`--keep-old` skips the delete step. Rows still flagged `needs_review` are updated anyway
(the existing names are known-bad); review them from the manifest.
