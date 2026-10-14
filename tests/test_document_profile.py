from __future__ import annotations

import hashlib
import io
import tempfile
import unittest
from dataclasses import FrozenInstanceError
from pathlib import Path
from unittest.mock import patch

import contextpack.core.document_profile as document_profile_module
from contextpack.core.document_profile import DocumentProfile, profile_document


class CountingReader(io.BytesIO):
    def __init__(self, content: bytes) -> None:
        super().__init__(content)
        self.read_sizes: list[int] = []

    def read(self, size: int = -1) -> bytes:
        self.read_sizes.append(size)
        return super().read(size)


class DocumentProfileTests(unittest.TestCase):
    def test_profile_reuses_processor_routing_and_identifies_format_family(self) -> None:
        cases = (
            ("report.PDF", "pdf", "pdf", ()),
            ("budget.XLSX", "excel", "excel", ()),
            ("budget.XLSM", "excel", "excel", ("macro_enabled_format",)),
            ("legacy.XLS", "excel", "document", ("legacy_binary_format",)),
            ("proposal.DOCX", "word", "document", ()),
            ("proposal.DOCM", "word", "document", ("macro_enabled_format",)),
            ("legacy.DOC", "word", "document", ("legacy_binary_format",)),
            ("scan.WEBP", "image", "image", ()),
            ("notes.TXT", "document", "document", ()),
        )
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for name, family, processor, safety_signals in cases:
                with self.subTest(name=name):
                    source = root / name
                    source.write_bytes(b"fixture")
                    profile = profile_document(source)
                    self.assertEqual(profile.extension, source.suffix.lower())
                    self.assertEqual(profile.format_family, family)
                    self.assertEqual(profile.processor, processor)
                    self.assertEqual(profile.safety_signals, safety_signals)

    def test_profile_hashes_and_describes_source_without_publishing_local_path(self) -> None:
        content = "ქართული document profile".encode()
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "ანგარიში.docx"
            source.write_bytes(content)
            profile = profile_document(source)

            self.assertEqual(profile.source_path, source.resolve())
            self.assertEqual(profile.source_name, source.name)
            self.assertEqual(profile.size_bytes, len(content))
            self.assertEqual(profile.sha256, hashlib.sha256(content).hexdigest())
            self.assertNotIn(str(source.resolve()), repr(profile))
            public = profile.public_metadata()
            self.assertNotIn("source_path", public)
            self.assertNotIn(str(source.parent.resolve()), str(public))
            self.assertEqual(public["safety_signals"], [])

    def test_profile_is_immutable_and_uses_immutable_signal_collections(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "fixture.pdf"
            source.write_bytes(b"pdf fixture")
            profile = profile_document(source)

            self.assertIsInstance(profile, DocumentProfile)
            self.assertIsInstance(profile.safety_signals, tuple)
            self.assertIsInstance(profile.quality_signals, tuple)
            with self.assertRaises(FrozenInstanceError):
                profile.processor = "changed"  # type: ignore[misc]

    def test_empty_file_has_quality_signal_and_source_metadata_is_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "empty.docx"
            source.write_bytes(b"")
            before = source.stat()
            profile = profile_document(source)
            after = source.stat()

            self.assertEqual(profile.quality_signals, ("empty_file",))
            self.assertEqual(profile.sha256, hashlib.sha256(b"").hexdigest())
            self.assertEqual(after.st_size, before.st_size)
            self.assertEqual(after.st_mtime_ns, before.st_mtime_ns)

    def test_hashing_reads_in_bounded_chunks(self) -> None:
        content = b"0123456789abcdef"
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "large.pdf"
            source.write_bytes(content)
            reader = CountingReader(content)
            with patch.object(Path, "open", return_value=reader):
                profile = profile_document(source, hash_chunk_size=4)

            self.assertEqual(profile.sha256, hashlib.sha256(content).hexdigest())
            self.assertGreater(len(reader.read_sizes), 1)
            self.assertTrue(all(size == 4 for size in reader.read_sizes))

    def test_invalid_hash_chunk_size_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "fixture.pdf"
            source.write_bytes(b"fixture")
            with self.assertRaisesRegex(ValueError, "positive"):
                profile_document(source, hash_chunk_size=0)

    def test_source_change_during_hashing_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "changing.pdf"
            source.write_bytes(b"before")
            original_hash_file = document_profile_module._hash_file

            def hash_then_change(path: Path, chunk_size: int) -> str:
                result = original_hash_file(path, chunk_size)
                path.write_bytes(b"changed after hashing")
                return result

            with patch.object(document_profile_module, "_hash_file", side_effect=hash_then_change):
                with self.assertRaisesRegex(RuntimeError, "changed"):
                    profile_document(source)

    def test_missing_path_and_directory_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            with self.assertRaises(FileNotFoundError):
                profile_document(root / "missing.pdf")
            with self.assertRaisesRegex(ValueError, "must be a file"):
                profile_document(root)


if __name__ == "__main__":
    unittest.main()
