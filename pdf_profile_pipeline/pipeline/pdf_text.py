"""Read text out of a PDF, in memory. The PDF itself is never modified."""
from __future__ import annotations

import io
import os
import re
import shutil
from pathlib import Path


class PdfTextError(RuntimeError):
    pass


def _import_pypdf():
    try:
        import pypdf  # noqa: PLC0415
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise SystemExit("ERROR: pypdf is required. Run: pip install -r requirements.txt") from exc
    return pypdf


def extract_text(pdf_bytes: bytes, *, max_pages: int | None = None) -> str:
    """Return the concatenated text layer of the PDF.

    Returns an empty string for scanned/image-only PDFs; the caller decides
    whether that is a failure. No OCR and no format conversion happens here.
    """
    pypdf = _import_pypdf()
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    except Exception as exc:
        raise PdfTextError(f"Could not open PDF: {exc}") from exc

    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception as exc:
            raise PdfTextError(f"PDF is encrypted and could not be opened: {exc}") from exc

    pages = reader.pages if max_pages is None else reader.pages[:max_pages]
    chunks: list[str] = []
    for index, page in enumerate(pages, start=1):
        try:
            chunks.append(page.extract_text() or "")
        except Exception as exc:
            chunks.append("")
            print(f"  warning: page {index} text extraction failed: {exc}")
    return normalize_text("\n".join(chunks))


def page_count(pdf_bytes: bytes) -> int:
    pypdf = _import_pypdf()
    try:
        return len(pypdf.PdfReader(io.BytesIO(pdf_bytes)).pages)
    except Exception:
        return 0


def normalize_text(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace(" ", " ").replace("–", "-").replace("—", "-")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return "\n".join(line.strip() for line in text.split("\n")).strip()


def looks_empty(text: str, *, min_chars: int = 40) -> bool:
    return len(re.sub(r"\s", "", text or "")) < min_chars


# --- OCR fallback for scanned/image-only PDFs -------------------------------
#
# Only used when the text layer is empty. Pages are rasterised in memory with
# PyMuPDF and read with Tesseract; the PDF itself is still never modified.

_WINDOWS_TESSERACT_PATHS = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
)


def tesseract_cmd() -> str | None:
    """Locate the Tesseract binary: TESSERACT_CMD, then PATH, then Program Files."""
    configured = (os.environ.get("TESSERACT_CMD") or "").strip().strip('"')
    if configured:
        return configured if Path(configured).exists() else None
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in _WINDOWS_TESSERACT_PATHS:
        if Path(candidate).exists():
            return candidate
    return None


def ocr_available() -> tuple[bool, str]:
    """Return (usable, reason) so callers can explain a disabled fallback."""
    try:
        import fitz  # noqa: F401, PLC0415
    except ImportError:
        return False, "PyMuPDF is not installed (pip install -r requirements.txt)"
    try:
        import pytesseract  # noqa: F401, PLC0415
    except ImportError:
        return False, "pytesseract is not installed (pip install -r requirements.txt)"
    if not tesseract_cmd():
        return False, (
            "the Tesseract binary was not found; install it and set TESSERACT_CMD "
            "if it is not on PATH"
        )
    return True, ""


def ocr_text(pdf_bytes: bytes, *, dpi: int = 300, max_pages: int | None = None) -> str:
    """Rasterise each page and read it with Tesseract. Returns normalised text."""
    import fitz  # noqa: PLC0415
    import pytesseract  # noqa: PLC0415
    from PIL import Image  # noqa: PLC0415

    command = tesseract_cmd()
    if command:
        pytesseract.pytesseract.tesseract_cmd = command

    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:
        raise PdfTextError(f"Could not open PDF for OCR: {exc}") from exc

    # 72 dpi is the PDF user-space unit, so this scales to the requested dpi.
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    chunks: list[str] = []
    with document:
        total = document.page_count if max_pages is None else min(document.page_count, max_pages)
        for index in range(total):
            try:
                pixmap = document.load_page(index).get_pixmap(matrix=matrix)
                image = Image.open(io.BytesIO(pixmap.tobytes("png")))
                chunks.append(pytesseract.image_to_string(image) or "")
            except Exception as exc:
                chunks.append("")
                print(f"  warning: OCR failed on page {index + 1}: {exc}")
    return normalize_text("\n".join(chunks))
