"""Durable source-bound review events; extracted text is never edited here."""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import shutil
import uuid

from .package_cache import file_hash, json_hash

STATUSES = {'extracted', 'needs_review', 'source_verified'}
CHECKLIST = """# Source verification checklist

Status: extracted / needs_review / source_verified. Successful conversion is not verification.

1. Identify original path, SHA-256, code, edition/revision and selected page coverage.
2. Open the original page/table with its headings, footnotes and surrounding paragraph.
3. Check digits, decimal comma/point, minus/plus, inequalities, exponents and subscripts.
4. Check units and scale; distinguish a dash/ditto mark from zero or a missing value.
5. For tables, verify the row label, column header, merged-cell spans and applicable notes.
6. For drawings, inspect views, sections, details and their relationships visually; OCR labels alone are insufficient.
7. Record page/clause/table/cell or region with its coordinate system, original extraction, corrected value (or null if unresolved), reviewer and method.
8. Confidence is diagnostic only. Fallback TSV boxes cover only fallback OCR pages and may refer to rotated/deskewed OCR pages. No TSV does not mean verified.
9. Never guess an unclear value. Append a needs_review event and retain the raw extraction.
10. Use source_verified values only within the recorded scope. Engineering applicability and final design decisions need an engineer's review.

Use review-journal.py append with a JSON entry. Events are append-only with a SHA-256 chain; it detects accidental modification, not malicious resealing or authorship.
Raw text/table/TSV snapshots live under raw/<snapshot SHA-256>/. Do not edit them. Later extractions create new snapshots; no correction is automatically applied to OCR.
Keep this sibling review directory when moving a portable package. Review metadata alone does not prove an entry's engineering correctness.
"""


def read_json(path: Path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def safe_directory(path: Path) -> None:
    for ancestor in (path, *path.parents):
        if ancestor.is_symlink() or (hasattr(ancestor, 'is_junction') and ancestor.is_junction()):
            raise ValueError('Review paths must not contain filesystem links')


@contextmanager
def journal_lock(directory: Path):
    safe_directory(directory)
    directory.mkdir(parents=True, exist_ok=True)
    lock = directory / '.writer.lock'
    descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    try:
        os.close(descriptor)
        yield
    finally:
        lock.unlink()


def initialize(package: Path, source: Path, review_root: Path) -> dict:
    """Snapshot only extracted text/metadata, never source PDFs or page images."""
    source_hash = file_hash(source)
    directory = review_root / source_hash
    safe_directory(package)
    inventory = {}
    for path in sorted(package.rglob('*')):
        relative = path.relative_to(package)
        safe_directory(path)
        if relative.parts[0] in {'original', 'assets', '.ocr-work'}:
            continue
        if path.is_file() and path.suffix.lower() in {'.md', '.json', '.tsv', '.txt', '.csv'} and path.name not in {'review.json', 'manifest.json'}:
            inventory[relative.as_posix()] = file_hash(path)
    snapshot = json_hash(inventory)
    with journal_lock(directory):
        if (directory / 'events.jsonl').exists():
            read_events(directory)
        identity = directory / 'source.json'
        if identity.exists():
            if read_json(identity)['sha256'] != source_hash:
                raise ValueError('Review source identity mismatch')
        else:
            identity.write_text(json.dumps({'sha256': source_hash, 'first_path': str(source.resolve()), 'revision': 'not established'}, ensure_ascii=False, indent=2), encoding='utf-8')
        destination = directory / 'raw' / snapshot
        if destination.exists():
            previous = read_json(destination / 'snapshot.json')
            if previous != inventory or any(file_hash(destination / name) != digest for name, digest in inventory.items()):
                raise ValueError('Raw review snapshot was modified; refusing to overwrite')
        else:
            stage = directory / ('snapshot-building-' + uuid.uuid4().hex)
            stage.mkdir()
            try:
                for name, digest in inventory.items():
                    target = stage / name
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(package / name, target)
                    if file_hash(target) != digest:
                        raise ValueError('Extraction changed while snapshotting')
                (stage / 'snapshot.json').write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding='utf-8')
                destination.parent.mkdir(exist_ok=True)
                stage.rename(destination)
            finally:
                if stage.exists():
                    shutil.rmtree(stage)
        checklist = directory / 'CHECKLIST.md'
        if not checklist.exists():
            checklist.write_text(CHECKLIST, encoding='utf-8')
        (directory / 'events.jsonl').touch(exist_ok=True)
    pointer = {'schema_version': 1, 'source_sha256': source_hash, 'initial_status': 'extracted',
               'journal_directory': str(directory.resolve()),
               'relative_to_output_root': directory.relative_to(review_root.parent).as_posix(),
               'raw_snapshot': snapshot, 'statuses': sorted(STATUSES),
               'scope': 'Review entries apply only to their recorded locators, never the entire source by implication.'}
    (package / 'review.json').write_text(json.dumps(pointer, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return pointer


def verify_package_review(package: Path, source_hash: str) -> None:
    pointer = read_json(package / 'review.json')
    relative = Path(pointer['relative_to_output_root'])
    if relative.is_absolute() or '..' in relative.parts or relative.parts[:1] != ('.contextpack-reviews',):
        raise ValueError('Invalid review path')
    directory = package.parent / relative
    safe_directory(directory)
    if pointer['source_sha256'] != source_hash or read_json(directory / 'source.json')['sha256'] != source_hash:
        raise ValueError('Review source mismatch')
    snapshot = pointer['raw_snapshot']
    if len(snapshot) != 64 or any(c not in '0123456789abcdef' for c in snapshot):
        raise ValueError('Invalid snapshot identity')
    raw = directory / 'raw' / snapshot
    inventory = read_json(raw / 'snapshot.json')
    if json_hash(inventory) != snapshot:
        raise ValueError('Snapshot inventory changed')
    for name, digest in inventory.items():
        relative_file = Path(name)
        if relative_file.is_absolute() or '..' in relative_file.parts:
            raise ValueError('Invalid snapshot file path')
        safe_directory(raw / relative_file)
        if file_hash(raw / relative_file) != digest:
            raise ValueError('Raw extraction changed')
    read_events(directory)


def read_events(directory: Path) -> list[dict]:
    safe_directory(directory)
    events = []
    previous = None
    for line in (directory / 'events.jsonl').read_text(encoding='utf-8').splitlines():
        event = json.loads(line)
        digest = event.pop('event_sha256')
        if event.get('previous_sha256') != previous or json_hash(event) != digest:
            raise ValueError('Review journal chain is damaged')
        event['event_sha256'] = digest
        events.append(event)
        previous = digest
    return events


def append_event(directory: Path, entry: dict) -> dict:
    with journal_lock(directory):
        if any(key in entry for key in ('event_id', 'event_sha256', 'previous_sha256', 'at_utc')):
            raise ValueError('Event identity and chain fields are generated by the journal')
        identity = read_json(directory / 'source.json')
        for key in ('record_id', 'source_sha256', 'source_revision', 'locator', 'original', 'corrected', 'status', 'reviewer', 'method', 'raw_snapshot'):
            if key not in entry:
                raise ValueError(f'Missing review field: {key}')
        if entry['source_sha256'] != identity['sha256']:
            raise ValueError('Entry belongs to a different source revision/hash')
        for key in ('record_id', 'source_revision', 'reviewer', 'method'):
            if not isinstance(entry[key], str) or not entry[key].strip():
                raise ValueError(f'Empty or invalid review field: {key}')
        if entry['status'] not in STATUSES:
            raise ValueError('Unknown review status')
        if not isinstance(entry['locator'], dict) or not entry['locator'] or not any(k in entry['locator'] for k in ('page', 'clause', 'table', 'cell', 'region', 'anchor')):
            raise ValueError('Provide a precise source locator')
        if 'region' in entry['locator'] and not entry['locator'].get('coordinate_system'):
            raise ValueError('Region requires a coordinate system')
        if not isinstance(entry['original'], str) or (entry['corrected'] is not None and not isinstance(entry['corrected'], str)):
            raise ValueError('Values must remain strings (or null for unresolved correction)')
        if entry['status'] != 'source_verified' and entry['corrected'] is not None:
            raise ValueError('Do not apply unverified corrections')
        if entry['status'] == 'source_verified' and entry['corrected'] is None:
            raise ValueError('Verified entries require the checked value (unchanged values are allowed)')
        snapshot = str(entry['raw_snapshot'])
        if len(snapshot) != 64 or any(c not in '0123456789abcdef' for c in snapshot) or not (directory / 'raw' / snapshot / 'snapshot.json').is_file():
            raise ValueError('Unknown raw extraction snapshot')
        events = read_events(directory)
        # Each new event supersedes only earlier events with this record_id; all remain readable.
        event = dict(entry)
        event.update({'event_id': uuid.uuid4().hex, 'at_utc': datetime.now(timezone.utc).isoformat(),
                      'previous_sha256': events[-1]['event_sha256'] if events else None})
        event['event_sha256'] = json_hash(event)
        with (directory / 'events.jsonl').open('a', encoding='utf-8', newline='\n') as stream:
            stream.write(json.dumps(event, ensure_ascii=False) + '\n')
            stream.flush()
            os.fsync(stream.fileno())
        return event


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='action', required=True)
    init = sub.add_parser('init')
    init.add_argument('package', type=Path)
    init.add_argument('source', type=Path)
    init.add_argument('review_root', type=Path)
    append = sub.add_parser('append')
    append.add_argument('directory', type=Path)
    append.add_argument('entry_json', type=Path)
    check = sub.add_parser('check')
    check.add_argument('directory', type=Path)
    args = parser.parse_args()
    if args.action == 'init':
        result = initialize(args.package, args.source, args.review_root)
    elif args.action == 'append':
        result = append_event(args.directory, read_json(args.entry_json))
    else:
        events = read_events(args.directory)
        result = {'events': len(events), 'chain_valid': True,
                  'latest_by_record': {e['record_id']: e for e in events}}
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
