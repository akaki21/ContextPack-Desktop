"""Content-addressed reuse checks for generated packages, never a trust signal."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
from pathlib import Path
import sys


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def json_hash(value: object) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf-8')).hexdigest()


def pipeline_signature(root: Path, settings: dict) -> str:
    paths = list(root.glob('*.py')) + list(root.glob('*.ps1'))
    paths += list((root / 'contextpack').rglob('*.py')) + list((root / 'powershell').rglob('*.ps1'))
    paths.append(root / 'requirements.txt')
    versions = {}
    for name in ('markitdown', 'ocrmypdf', 'pypdfium2', 'Pillow', 'beautifulsoup4', 'html5lib', 'markdownify'):
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = None
    return json_hash({'schema': 1, 'settings': settings, 'python': list(sys.version_info[:3]), 'dependencies': versions,
                      'code': {p.relative_to(root).as_posix(): hashlib.sha256(p.read_text(encoding='utf-8-sig').encode('utf-8')).hexdigest() for p in sorted(set(paths))}})


def package_files(package: Path) -> dict[str, str]:
    result = {}
    for path in sorted(package.rglob('*')):
        if path.is_symlink() or (hasattr(path, 'is_junction') and path.is_junction()):
            raise ValueError('Generated packages must not contain filesystem links')
        if path.is_file() and path != package / 'manifest.json':
            result[path.relative_to(package).as_posix()] = file_hash(path)
    return result


def validate_cache(package: Path, signature: str, source_hash: str) -> tuple[bool, str]:
    try:
        if package.is_symlink() or (hasattr(package, 'is_junction') and package.is_junction()):
            return False, 'package_is_link'
        manifest = json.loads((package / 'manifest.json').read_text(encoding='utf-8-sig'))
        cache = manifest.pop('cache', {})
        if manifest['source']['sha256'] != source_hash or cache.get('signature') != signature:
            return False, 'source_or_pipeline_changed'
        if cache.get('manifest_sha256') != json_hash(manifest):
            return False, 'manifest_changed'
        if not cache.get('files') or cache['files'] != package_files(package):
            return False, 'outputs_changed'
        return True, 'verified'
    except (OSError, ValueError, KeyError, TypeError):
        return False, 'unreadable_or_unsealed'


def seal_cache(package: Path, signature: str) -> None:
    path = package / 'manifest.json'
    manifest = json.loads(path.read_text(encoding='utf-8-sig'))
    manifest.pop('cache', None)
    manifest['cache'] = {'signature': signature, 'manifest_sha256': json_hash(manifest), 'files': package_files(package)}
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['signature', 'probe', 'seal'])
    parser.add_argument('path', type=Path)
    parser.add_argument('--signature', default='')
    parser.add_argument('--source-hash', default='')
    args = parser.parse_args()
    if args.action == 'signature':
        # PowerShell sends UTF-8 bytes; Windows' legacy stdin locale may differ.
        print(pipeline_signature(args.path, json.loads(sys.stdin.buffer.read())))
    elif args.action == 'probe':
        valid, reason = validate_cache(args.path, args.signature, args.source_hash)
        print(json.dumps({'valid': valid, 'reason': reason}))
    else:
        seal_cache(args.path, args.signature)


if __name__ == '__main__':
    main()
