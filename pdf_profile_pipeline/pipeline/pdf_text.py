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


def normalize_text(text: str, *, keep_columns: bool = False) -> str:
    """Tidy whitespace and dashes.

    keep_columns collapses runs of spaces to exactly two instead of one. OCR of
    a multi-column section (a skills grid, say) marks the column break only with
    that run of spaces, so squashing it to one would fuse the columns into a
    single line that can no longer be split back into items.
    """
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace(" ", " ").replace("–", "-").replace("—", "-")
    text = re.sub(r"\t", " ", text)
    text = re.sub(r" {2,}", "  " if keep_columns else " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return "\n".join(line.strip() for line in text.split("\n")).strip()


def looks_empty(text: str, *, min_chars: int = 40) -> bool:
    return len(re.sub(r"\s", "", text or "")) < min_chars


# --- OCR fallback for scanned/image-only PDFs -------------------------------
#
# Only used when the text layer is empty. Pages are rasterised in memory with
# PyMuPDF and read with Tesseract; the PDF itself is still never modified.

_WINDOWS_TESSERACT_DIRS = (
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR"),
    os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR"),
)

_EXE_NAME = "tesseract.exe" if os.name == "nt" else "tesseract"


class TesseractError(RuntimeError):
    """Tesseract is installed but could not be run."""


def _resolve_binary(candidate: str | os.PathLike[str]) -> Path | None:
    """Turn a configured path into the actual executable, or None.

    Accepts the install *directory* as well as the binary. This is the fix for
    `[WinError 5] Access is denied`: a value like
    ``C:\\Program Files\\Tesseract-OCR`` passes ``Path.exists()`` because the
    folder exists, and CreateProcess then refuses to execute a directory.
    """
    path = Path(str(candidate).strip().strip('"').strip("'")).expanduser()
    if not str(path):
        return None
    if path.is_dir():
        path = path / _EXE_NAME
    elif os.name == "nt" and path.suffix == "":
        # "…\tesseract" with no extension is not directly executable.
        with_exe = path.with_suffix(".exe")
        if with_exe.is_file():
            path = with_exe
    return path if path.is_file() else None


def tesseract_cmd() -> str | None:
    """Locate the Tesseract binary: TESSERACT_CMD, then PATH, then Program Files.

    Always returns a path to a real *file*, never a directory, so the value can
    be handed straight to subprocess.
    """
    configured = os.environ.get("TESSERACT_CMD") or ""
    if configured.strip():
        resolved = _resolve_binary(configured)
        # A bad TESSERACT_CMD should not mask a working default install, so we
        # fall through to the usual search instead of returning None here.
        if resolved is not None:
            return str(resolved)

    found = shutil.which("tesseract")
    if found:
        resolved = _resolve_binary(found)
        if resolved is not None:
            return str(resolved)

    for directory in _WINDOWS_TESSERACT_DIRS:
        resolved = _resolve_binary(directory)
        if resolved is not None:
            return str(resolved)
    return None


def tessdata_dir(command: str | None = None) -> str | None:
    """Best guess at the tessdata folder, so TESSDATA_PREFIX can be set for us.

    Tesseract 5 wants TESSDATA_PREFIX to be the tessdata directory itself. An
    unset or stale value is the usual cause of "Error opening data file
    eng.traineddata".
    """
    configured = (os.environ.get("TESSDATA_PREFIX") or "").strip().strip('"')
    if configured:
        base = Path(configured).expanduser()
        for candidate in (base, base / "tessdata"):
            if (candidate / "eng.traineddata").is_file():
                return str(candidate)
    command = command or tesseract_cmd()
    if command:
        candidate = Path(command).parent / "tessdata"
        if (candidate / "eng.traineddata").is_file():
            return str(candidate)
    return None


def _configure_pytesseract():
    """Point pytesseract at the resolved binary and tessdata. Returns the module."""
    import pytesseract  # noqa: PLC0415

    command = tesseract_cmd()
    if not command:
        raise TesseractError(
            "the Tesseract binary was not found; install it or set TESSERACT_CMD "
            f"to {_EXE_NAME} (the executable, not the folder)"
        )
    pytesseract.pytesseract.tesseract_cmd = command

    data_dir = tessdata_dir(command)
    if data_dir:
        # Set it for this process and every subprocess Tesseract spawns.
        os.environ["TESSDATA_PREFIX"] = data_dir
    return pytesseract


def tesseract_version() -> str:
    """Run the binary and return its version. Raises TesseractError if it cannot run."""
    import subprocess  # noqa: PLC0415

    command = tesseract_cmd()
    if not command:
        raise TesseractError(
            "the Tesseract binary was not found; install it or set TESSERACT_CMD "
            f"to {_EXE_NAME} (the executable, not the folder)"
        )
    _configure_pytesseract()
    try:
        completed = subprocess.run(
            [command, "--version"],
            capture_output=True, text=True, timeout=30, check=False,
        )
    except PermissionError as exc:  # WinError 5
        raise TesseractError(
            f"{command} could not be executed ({exc}). Check that it is the "
            f"{_EXE_NAME} file itself and that antivirus or policy is not blocking it."
        ) from exc
    except OSError as exc:
        raise TesseractError(f"{command} could not be executed: {exc}") from exc

    output = (completed.stdout or completed.stderr or "").strip()
    if completed.returncode != 0 and not output:
        raise TesseractError(f"{command} --version exited {completed.returncode}")
    return output.splitlines()[0].strip() if output else "tesseract (version unknown)"


def selftest() -> tuple[bool, str]:
    """Startup probe: confirm Tesseract really runs. Returns (ok, message)."""
    try:
        import fitz  # noqa: F401, PLC0415
    except ImportError:
        return False, "PyMuPDF is not installed (pip install -r requirements.txt)"
    try:
        import pytesseract  # noqa: F401, PLC0415
    except ImportError:
        return False, "pytesseract is not installed (pip install -r requirements.txt)"
    try:
        version = tesseract_version()
    except TesseractError as exc:
        return False, str(exc)

    command = tesseract_cmd()
    data_dir = tessdata_dir(command)
    if not data_dir:
        return False, (
            f"{command} runs but no eng.traineddata was found; set TESSDATA_PREFIX "
            "to the tessdata folder"
        )
    return True, f"{version} [{command}] tessdata={data_dir}"


def ocr_available() -> tuple[bool, str]:
    """Return (usable, reason) so callers can explain a disabled fallback."""
    ok, message = selftest()
    return (True, "") if ok else (False, message)


def _render_page(page, dpi: int):
    """Rasterise one page to a grayscale PIL image."""
    import fitz  # noqa: PLC0415
    from PIL import Image  # noqa: PLC0415

    # 72 dpi is the PDF user-space unit, so this scales to the requested dpi.
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    pixmap = page.get_pixmap(matrix=matrix, colorspace=fitz.csGRAY, alpha=False)
    return Image.open(io.BytesIO(pixmap.tobytes("png")))


def ocr_text(
    pdf_bytes: bytes,
    *,
    dpi: int = 300,
    max_pages: int | None = None,
    lang: str = "eng",
) -> str:
    """Rasterise each page and read it with Tesseract. Returns normalised text.

    Raises TesseractError if Tesseract cannot be run at all; a single bad page
    is logged and skipped so one damaged scan never loses the whole resume.
    """
    import fitz  # noqa: PLC0415

    pytesseract = _configure_pytesseract()

    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:
        raise PdfTextError(f"Could not open PDF for OCR: {exc}") from exc

    # --psm 3 = automatic page segmentation, the right mode for a full page.
    tess_config = os.environ.get("TESSERACT_CONFIG", "--oem 3 --psm 3")
    chunks: list[str] = []
    failures = 0
    with document:
        total = document.page_count if max_pages is None else min(document.page_count, max_pages)
        for index in range(total):
            try:
                image = _render_page(document.load_page(index), dpi)
                chunks.append(pytesseract.image_to_string(image, lang=lang, config=tess_config) or "")
            except PermissionError as exc:  # WinError 5, first page onwards
                raise TesseractError(
                    f"Tesseract at {tesseract_cmd()} could not be executed ({exc}). "
                    "Point TESSERACT_CMD at tesseract.exe, not the install folder."
                ) from exc
            except Exception as exc:
                failures += 1
                chunks.append("")
                print(f"  warning: OCR failed on page {index + 1}: {exc}")
        if failures and failures == total:
            raise TesseractError(f"OCR failed on all {total} page(s); see the warnings above.")
    # Column runs are kept: a scanned skills grid is only separable by them.
    return normalize_text("\n".join(chunks), keep_columns=True)
