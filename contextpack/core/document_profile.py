"""Build a read-only, processor-neutral profile of an input file."""

from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
from pathlib import Path
from typing import Any

from contextpack.file_types import EXCEL_EXTENSIONS, IMAGE_EXTENSIONS, classify_input


WORD_EXTENSIONS = frozenset({".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm"})
EXCEL_FAMILY_EXTENSIONS = EXCEL_EXTENSIONS | frozenset({".xls", ".xlt"})
MACRO_ENABLED_EXTENSIONS = frozenset({".docm", ".dotm", ".xlsm", ".xltm"})
LEGACY_BINARY_EXTENSIONS = frozenset({".doc", ".dot", ".xls", ".xlt", ".ppt", ".pot"})
DEFAULT_HASH_CHUNK_SIZE = 1024 * 1024


@dataclass(frozen=True, slots=True)
class DocumentProfile:
    """Immutable facts collected without opening the document in an Office app."""

    source_path: Path = field(repr=False)
    source_name: str
    extension: str
    format_family: str
    processor: str
    size_bytes: int
    sha256: str
    safety_signals: tuple[str, ...] = ()
    quality_signals: tuple[str, ...] = ()

    def public_metadata(self) -> dict[str, Any]:
        """Return JSON-ready metadata without exposing the local absolute path."""

        return {
            "source_name": self.source_name,
            "extension": self.extension,
            "format_family": self.format_family,
            "processor": self.processor,
            "size_bytes": self.size_bytes,
            "sha256": self.sha256,
            "safety_signals": list(self.safety_signals),
            "quality_signals": list(self.quality_signals),
        }


def _format_family(extension: str) -> str:
    if extension == ".pdf":
        return "pdf"
    if extension in EXCEL_FAMILY_EXTENSIONS:
        return "excel"
    if extension in WORD_EXTENSIONS:
        return "word"
    if extension in IMAGE_EXTENSIONS:
        return "image"
    return "document"


def _hash_file(path: Path, chunk_size: int) -> str:
    if chunk_size <= 0:
        raise ValueError("Hash chunk size must be positive.")

    digest = sha256()
    with path.open("rb") as source:
        while chunk := source.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest()


def profile_document(path: Path, *, hash_chunk_size: int = DEFAULT_HASH_CHUNK_SIZE) -> DocumentProfile:
    """Return stable file facts while leaving the source completely untouched."""

    requested_path = Path(path)
    if not requested_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {requested_path}")
    if not requested_path.is_file():
        raise ValueError(f"Input must be a file: {requested_path}")

    source_path = requested_path.resolve(strict=True)
    before = source_path.stat()
    extension = source_path.suffix.lower()
    file_hash = _hash_file(source_path, hash_chunk_size)
    after = source_path.stat()
    if before.st_size != after.st_size or before.st_mtime_ns != after.st_mtime_ns:
        raise RuntimeError("Input file changed while its profile was being created.")

    safety_signals: list[str] = []
    if extension in MACRO_ENABLED_EXTENSIONS:
        safety_signals.append("macro_enabled_format")
    if extension in LEGACY_BINARY_EXTENSIONS:
        safety_signals.append("legacy_binary_format")

    quality_signals: list[str] = []
    if before.st_size == 0:
        quality_signals.append("empty_file")

    return DocumentProfile(
        source_path=source_path,
        source_name=source_path.name,
        extension=extension,
        format_family=_format_family(extension),
        processor=classify_input(source_path),
        size_bytes=before.st_size,
        sha256=file_hash,
        safety_signals=tuple(safety_signals),
        quality_signals=tuple(quality_signals),
    )
