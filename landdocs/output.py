"""Automatic client filenames; generated documents remain create-only."""
from pathlib import Path
import re
import sys

from .engine import ROOT, generate_docx

OUTPUT_DIR = (Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else ROOT) / "output"


def client_filename(client_name):
    # DevLys glyph codes remain lossless in client data. Only filename-unsafe
    # characters are replaced; no unverified Hindi transliteration is implied.
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", client_name).strip(" .")
    stem = stem.encode("utf-8")[:180].decode("utf-8", errors="ignore").rstrip(" .")
    return f"{stem or 'client'}_data.docx"


def generate_client_docx(client, directory=None):
    directory = Path(directory) if directory is not None else OUTPUT_DIR
    directory.mkdir(parents=True, exist_ok=True)
    filename = Path(client_filename(client.client_name))
    number = 1
    while True:
        name = filename.name if number == 1 else f"{filename.stem}_{number}{filename.suffix}"
        try:
            return generate_docx(client, directory / name)
        except FileExistsError:
            # Includes atomic publish races: an existing DOCX is never replaced.
            if not directory.is_dir():
                raise
            number += 1
