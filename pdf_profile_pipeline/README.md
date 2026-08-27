# PDF Profile Pipeline

Reads resume **PDFs directly from SharePoint/OneDrive**, extracts profile details with a
**Qwen2.5** model, uploads the **original, untouched PDF** to **Cloudflare R2**, and
inserts/updates the profile in **Neon Postgres**.

There is no image conversion and no DOC/DOCX conversion anywhere in this project.
The PDF that comes out of SharePoint is byte-for-byte the PDF that lands in R2.

```
SharePoint PDF
      -> read PDF (text layer, in memory)
      -> extract profile details (Qwen2.5)
      -> resolve years of experience
      -> generate Profile ID
      -> upload the original PDF to Cloudflare R2
      -> insert/update the profile in Neon
      -> log every record
```

## What Gets Extracted

| Field | Source |
| --- | --- |
| `profile_id` | Generated: UUIDv5 over the PDF's SHA-256, so the same PDF always maps to the same row |
| `first_name` | Model, falling back to a parsed full name |
| `last_name` | Model, falling back to a parsed full name |
| `full_name` | Model; stored in `resume_sections.full_name` and `search_text` |
| `city` | Model, falling back to a `City, ST` match |
| `state_code` | Model, normalised to a 2-letter US code |
| `specialty` | Model, falling back to a labelled `Specialty:` line |
| `years_experience` | Resolved from several signals — see below |
| `bio` | Model-written professional summary, falling back to the first long paragraph |

## How Years of Experience Is Decided

A resume usually states experience in more than one way at once: a summary line
("12+ years of experience"), a single job's duration ("3 years at one clinic"), and a
list of employment date ranges. Taking the first match found is wrong, because it often
picks one job's length instead of the career total.

[pipeline/experience.py](pipeline/experience.py) collects **every** candidate value,
labels each by how it was stated, and applies this priority order:

1. **`explicit_total`** — a statement near words like *total*, *overall*, *combined*,
   *career* (e.g. "15 years of total nursing experience").
2. **`summary_statement`** — an experience figure inside the summary/profile section at
   the top of the resume.
3. **`model_total`** — the model's own `total_years_experience`, which it is prompted to
   compute across the whole career rather than per job.
4. **`date_range_span`** — the union of all employment date ranges, with overlapping
   jobs counted once (so two concurrent 2015–2020 roles are 5 years, not 10). Handles
   `Present`/`Current` end dates.
5. **`single_duration`** — the largest individual stated duration, used only when
   nothing better exists.

Within one tier the larger value wins, and anything outside `0 < years <= 60` is
discarded as implausible. The chosen value, the tier it came from, and every rejected
candidate are written to `resume_sections.experience_decision` and to the logs, so any
number in the database can be traced back to the text that produced it.

## Setup

```powershell
cd pdf_profile_pipeline
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env      # then fill in real values
```

### Qwen2.5 with Ollama

```powershell
ollama pull qwen2.5:7b
ollama serve
```

Then in `.env`:

```env
LLM_BASE_URL=http://localhost:11434/v1
LLM_API_KEY=ollama
LLM_MODEL=qwen2.5:7b
```

Any OpenAI-compatible endpoint works (vLLM, LM Studio, a hosted gateway) — only
`LLM_BASE_URL`, `LLM_API_KEY`, and `LLM_MODEL` change. No model details are hardcoded.

### SharePoint access

A Microsoft Entra app registration is required.

- **Device login** (interactive): set `ONEDRIVE_CLIENT_ID`, leave `ONEDRIVE_CLIENT_SECRET`
  blank, and enable public client flows on the app registration.
- **App-only** (unattended): set `ONEDRIVE_CLIENT_ID`, `ONEDRIVE_CLIENT_SECRET`, and
  `ONEDRIVE_TENANT` to your tenant ID, with `Files.Read.All` / `Sites.Read.All`
  application permissions granted.

## Usage

```powershell
# SharePoint / OneDrive shared folder
python run_pipeline.py "https://<your-share-link>"

# Try it first without writing anything
python run_pipeline.py "https://<your-share-link>" --limit 5 --dry-run

# Local PDFs
python run_pipeline.py "C:\path\to\pdfs"

# Faster on large folders
python run_pipeline.py "https://<your-share-link>" --workers 4

# Regex-only extraction, no model call
python run_pipeline.py "C:\path\to\pdfs" --no-llm
```

### Options

| Flag | Purpose |
| --- | --- |
| `--limit N` | Process only the first N PDFs |
| `--workers N` | Parallel download/extract/upload workers (default 1) |
| `--dry-run` | Read and extract only; no R2 upload, no database write |
| `--no-llm` | Skip Qwen2.5 and use regex extraction only |
| `--target-table` | Neon table to write (default `profiles`) |
| `--env-file PATH` | Use a different env file |
| `--manifest PATH` \| `none` | JSONL manifest location |
| `--upload-log PATH` \| `none` | CSV upload log location |
| `--run-log PATH` \| `none` | Console transcript location |
| `--ignore-manifest` | Do not skip records already marked processed |
| `--retry-all` | Reprocess everything, including completed records |
| `--onedrive-client-id/-secret/-tenant` | Graph credentials, overriding `.env` |
| `--no-browser` | Print the device login code instead of opening a browser |

## Logging

Every run writes three logs under `logs/` (all can be redirected or disabled):

| File | Contents |
| --- | --- |
| `logs/profiles_manifest.jsonl` | One JSON line per record: status, profile ID, PDF hash, page count, R2 key, resume URL, extraction mode, and the full experience decision. Reruns read this file and skip records already `processed` or `skipped_duplicate`. |
| `logs/upload_log.csv` | One row per record, spreadsheet-friendly: timestamp, status, profile ID, insert vs update, name, city/state, specialty, years and which signal produced them, PDF hash and pages, R2 key, whether the object was newly uploaded, resume URL, source file, extraction mode, error. |
| `logs/run_<timestamp>.log` | The complete console transcript of the run. |

Statuses you will see: `processed`, `would_process` (dry run), `skipped_unusable`
(no first/last name could be found, and those columns are `NOT NULL`), and
`failed` with the error text.

## Reruns and Duplicates

- **R2** keys are content-addressed (`<prefix>/<folder>/<sha256>.pdf`), so re-uploading
  an identical PDF reuses the existing object instead of duplicating it.
- **Neon** rows are matched on the generated profile ID, then the PDF hash inside
  `resume_sections`, then first + last name with a matching city. A match updates the
  existing row; otherwise a new row is inserted.
- **The manifest** skips records already completed, unless `--retry-all` or
  `--ignore-manifest` is passed.

## Verification

```sql
SELECT profile_id, first_name, last_name, city, state_code, specialty,
       years_experience, resume_url, updated_at
FROM profiles
WHERE capture_source = 'pdf_profile_pipeline'
ORDER BY updated_at DESC
LIMIT 20;
```

Trace one profile's experience decision:

```sql
SELECT profile_id,
       resume_sections -> 'experience_decision' AS experience_decision
FROM profiles
WHERE profile_id = '<profile-id>';
```

## Project Layout

```
pdf_profile_pipeline/
  run_pipeline.py          CLI entry point
  pipeline/
    config.py              .env loading, validation, settings
    sharepoint.py          Microsoft Graph auth, PDF discovery, download
    pdf_text.py            PDF text extraction (pypdf, in memory)
    llm_extract.py         Qwen2.5 prompt, JSON parsing, regex fallback
    experience.py          Years-of-experience resolution
    storage.py             Cloudflare R2 upload of the original PDF
    database.py            Neon profile insert/update
    run_log.py             Manifest, CSV, and run-log writers
    processor.py           Orchestration
  requirements.txt
  .env.example
  .gitignore
  README.md
```

## Notes

- **Extraction order:** PDF text layer -> Tesseract OCR -> Qwen 2.5 -> validation ->
  database. Each step only runs when the one before it came up empty, and a failure
  at any step is recorded on the record rather than ending it. A PDF is never
  dropped just for lacking a text layer.
- **OCR is on by default.** Scanned, image-only PDFs are rasterised in memory with
  PyMuPDF and read with Tesseract; the PDF itself is never modified. Use `--no-ocr`
  to turn it off, or `--require-ocr` to make a broken Tesseract fatal instead of a
  warning. At startup the pipeline runs `tesseract --version` and logs the binary
  path, version and tessdata folder.
- **Tesseract on Windows:** `TESSERACT_CMD` must point at `tesseract.exe`, but the
  install *directory* is accepted too and resolved to the executable. Pointing it at
  the folder is what produced `[WinError 5] Access is denied`: the path existed, so
  it was passed to `CreateProcess`, which cannot execute a directory. If
  `TESSERACT_CMD` is unset or wrong, the usual install locations are searched, and
  `TESSDATA_PREFIX` is set automatically from the binary's own `tessdata` folder.
- If the model is unreachable or returns unusable JSON, the run continues with regex
  extraction and records `extraction_mode = regex_fallback` plus the model error. A
  reply cut off by the token limit is repaired rather than discarded.
- Work history is written to `work_history`, one row per position, keyed by a
  deterministic `work_id` so reruns never duplicate. Existing rows are left as they
  are, so hand corrections survive a reparse.
- Updates never destroy data: every optional column is merged with
  `COALESCE(NULLIF(new, ''), existing)`, so a weaker rerun can only add to a row.
- The `profiles` table must already exist in Neon with the expected columns; the
  pipeline validates this at startup and exits with a clear message if not.
