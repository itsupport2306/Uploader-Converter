# Screenshot → Word converter

Converts **Doximity-style physician-profile screenshots** into clean, editable
Word (`.docx`) documents using local Tesseract OCR. The right-hand sidebar
(phone, fax, address, "Similar Physicians & HCPs") and Doximity's promo /
"Join to view" chrome are dropped automatically — only the physician's data is
kept, formatted to mirror the on-screen layout.

The same command also accepts NPPES-style `.xlsx`, `.xlsm`, and `.csv` provider
files. Spreadsheet inputs write one Word document per data row using the same
resume-style typography and section layout. The Excel mapper supports the
attached Part01 headers such as `NPI`, `First Name`, `Last Name`, `Provider Type`,
`Specialty`, `Street Address`, `License Number`, and
`Enumeration Date`.

## What it does

1. **De-inverts the dark header band** so the white-on-dark name is read.
2. **Detects the column gap** and crops away the right sidebar + contact block.
3. **Rebuilds structure** from OCR word boxes — name, specialty, section
   headings, and the entries inside each section (grouped by vertical spacing).
4. **Strips icon/logo artifacts** that OCR leaves at the start of entries
   (e.g. a `Hy`, `lg)`, `*` from the small institution logos).
5. **Renders a styled `.docx`**: large bold name, blue specialty line, bold
   section headings with divider rules, bold entry titles, gray detail lines.
6. **Reads Excel/CSV provider rows** and maps NPPES identity, address, phone,
   and identifier fields into the same styled resume format.

## Requirements

- **Python 3.10+**
- **Tesseract OCR** (the binary, installed separately):
  Windows build → https://github.com/UB-Mannheim/tesseract/wiki
  The tool auto-detects `C:\Program Files\Tesseract-OCR\tesseract.exe`; if yours
  is elsewhere, pass `--tesseract "C:\path\to\tesseract.exe"`.
- Python packages: `pip install -r requirements.txt`

## Usage

```bash
# Single screenshot -> Alan_Lawrence_Aarons_MD.docx next to it
python screenshot_to_word.py "samples/doximity_sample.png"

# Choose the output file
python screenshot_to_word.py "profile.png" -o "Dr_Aarons.docx"

# Convert every image in a folder (writes one .docx per image)
python screenshot_to_word.py "C:\screenshots" -o "C:\out_docs"

# Convert an NPPES Excel workbook (writes one .docx per data row)
python screenshot_to_word.py "C:\Downloads\NPPES_Provider_Database_Part01.xlsx" -o "C:\out_docs"

# Try only the first 10 provider rows
python screenshot_to_word.py "C:\Downloads\NPPES_Provider_Database_Part01.xlsx" -o "C:\out_docs" --limit 10

# Resume checkpoints are saved automatically for Excel/CSV runs.
# To force a specific row, pass --start-row. To ignore checkpoints, pass --no-resume.

# Upload generated docs to a OneDrive account chosen in the browser.
# Requires a Microsoft Graph public client ID.
python screenshot_to_word.py "C:\Downloads\NPPES_Provider_Database_Part01.xlsx" --limit 10 --onedrive-upload --onedrive-client-id "YOUR_CLIENT_ID" --onedrive-tenant consumers --onedrive-folder "NPPES Resumes" --onedrive-upload-delay 0.5

# See exactly what structure was detected (useful for tuning)
python screenshot_to_word.py "profile.png" --debug
```

### Options

| Flag | Default | Purpose |
|------|---------|---------|
| `-o, --output` | next to input | Output `.docx` (single input) or output folder (batch) |
| `--tesseract PATH` | auto-detect | Explicit path to `tesseract.exe` |
| `--scale FLOAT` | `1.5` | Upscale factor before OCR (raise for small/blurry shots) |
| `--cutoff-ratio FLOAT` | `0.66` | Fallback main-column width fraction if the gap isn't auto-found |
| `--min-conf INT` | `35` | Min OCR confidence for body text. Lower toward 0 for max recall, but expect icon/seal noise |
| `--keep-promo` | off | Keep Doximity join/promo lines instead of dropping them |
| `--sheet NAME` | active sheet | Worksheet name for Excel inputs |
| `--limit N` | all rows | Maximum number of Excel/CSV data rows to process |
| `--start-row N` | checkpoint or `2` | First Excel/CSV data row number to process |
| `--no-resume` | off | Ignore saved Excel/CSV row checkpoints |
| `--onedrive-upload` | off | Upload generated docs to the OneDrive account you sign into in the browser |
| `--onedrive-client-id ID` | `ONEDRIVE_CLIENT_ID` env var | Microsoft Graph public client ID used for browser/device login |
| `--onedrive-tenant VALUE` | `consumers` | Auth tenant: `consumers`, `common`, or a work/school tenant ID |
| `--onedrive-folder PATH` | `Converter Output` | Remote OneDrive folder under root |
| `--onedrive-upload-delay SEC` | `0` | Pause after each upload; useful for long OneDrive runs that get throttled |
| `--keep-local` | off | Keep local `.docx` files after successful OneDrive upload |
| `--debug` | off | Print the detected blocks |

## Notes & limitations

- Output quality is bounded by Tesseract OCR. Occasional character errors
  (e.g. `II IGF` read as `IIIGF`) are expected; the structure stays correct.
- A publication title that wraps to a second line shows the wrapped part as a
  detail line — content is preserved, structure is approximate.
- A faint glyph next to a logo/seal (e.g. the 2-letter state in "IN State
  Medical License") may be read at zero confidence and dropped at the default
  `--min-conf 35`. Lowering it recovers such tokens but lets in seal noise on
  other entries, so the clean default is preferred for batch runs.
- An entry logo occasionally OCRs into a short letter fragment that lands at the
  start of a line (e.g. "Jee Henry Ford..."). Symbol fragments are removed; plain
  letter fragments are left as-is because they're indistinguishable from real
  2-letter codes/initials (a state code, a middle initial).
- Tuned for the Doximity profile layout. For other layouts, adjust
  `KNOWN_HEADINGS` / `PROMO_SNIPPETS` near the top of `screenshot_to_word.py`.

`samples/out.docx` is an example output generated from `samples/doximity_sample.png`.
