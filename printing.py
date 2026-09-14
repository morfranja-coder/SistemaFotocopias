from __future__ import annotations

import math
import os
from pathlib import Path
from typing import Optional

from pypdf import PdfReader


SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".docx", ".txt", ".rtf"}


def count_pdf_pages(file_path: str | os.PathLike) -> int:
    reader = PdfReader(str(file_path))
    return len(reader.pages)


def count_word_pages(file_path: str | os.PathLike) -> Optional[int]:
    """Returns Word's computed page count when Microsoft Word + pywin32 are available."""
    try:
        import win32com.client  # type: ignore
    except Exception:
        return None

    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(Path(file_path).resolve()), ReadOnly=True)
        # wdStatisticPages = 2
        return int(document.ComputeStatistics(2))
    except Exception:
        return None
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def detect_page_count(file_path: str | os.PathLike) -> Optional[int]:
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return count_pdf_pages(path)
    if suffix in {".doc", ".docx"}:
        return count_word_pages(path)
    return None


def calculate_physical_sheets(pages: int, copies: int = 1, duplex: bool = False) -> int:
    if pages <= 0 or copies <= 0:
        raise ValueError("Paginas y copias deben ser mayores que cero")
    per_copy = math.ceil(pages / 2) if duplex else pages
    return per_copy * copies


def send_to_printer(file_path: str | os.PathLike) -> None:
    """Sends a file to Windows' default print handler.

    This launches the registered application's print verb. It only returns when
    Windows accepts the command; it cannot guarantee that paper physically left
    the printer. The caller should register the movement only after this call
    succeeds and offer reversal for mechanical/user errors.
    """
    if os.name != "nt":
        raise RuntimeError("La impresion directa solo esta disponible en Windows")

    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(str(path))
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Formato no soportado para impresion: {path.suffix}")

    try:
        os.startfile(str(path), "print")  # type: ignore[attr-defined]
    except OSError as exc:
        raise RuntimeError(
            "Windows no pudo enviar el archivo a imprimir. Verifique que exista una aplicacion asociada al formato."
        ) from exc
