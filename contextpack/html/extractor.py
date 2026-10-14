from __future__ import annotations

import argparse
from html import escape
import json
from pathlib import Path
import re
import shutil
from urllib.parse import unquote, urlsplit

from bs4 import BeautifulSoup, NavigableString, Tag
from markitdown import MarkItDown

from contextpack.core.package_cache import file_hash, json_hash


MAX_GRID_SLOTS = 1_000_000
SAFE_INLINE_TAGS = {'a', 'p', 'br', 'b', 'strong', 'i', 'em', 'u', 'sub', 'sup', 'span', 'code', 'img', 'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption'}


def safe_inline(node: Tag | NavigableString) -> str:
    if isinstance(node, NavigableString):
        return escape(str(node))
    if node.name in {'script', 'style', 'iframe', 'object', 'embed'}:
        return ''
    content = ''.join(safe_inline(child) for child in node.children)
    if node.name not in SAFE_INLINE_TAGS:
        return content
    attributes = []
    for key in ('id', 'name', 'href', 'src', 'alt', 'rowspan', 'colspan'):
        value = node.get(key)
        if not isinstance(value, str):
            continue
        if key in {'href', 'src'} and urlsplit(value.strip()).scheme.lower() not in {'', 'http', 'https'}:
            continue
        attributes.append(f' {key}="{escape(value, quote=True)}"')
    tag = node.name
    return f'<{tag}{"".join(attributes)}>' + ('' if tag in {'br', 'img'} else content + f'</{tag}>')


def positive_span(cell: Tag, name: str, remaining_rows: int, warnings: list[str]) -> int:
    raw = str(cell.get(name, '1'))
    if name == 'rowspan' and raw == '0':
        return remaining_rows
    try:
        value = int(raw)
    except ValueError:
        value = 1
        warnings.append(f'Invalid {name}={raw!r}; inspect the original table.')
    limit = remaining_rows if name == 'rowspan' else 1000
    if not 1 <= value <= limit:
        warnings.append(f'{name}={raw!r} was bounded to the table limits; inspect the original.')
    return max(1, min(limit, value))


def extract_table(table: Tag, number: int) -> dict:
    """Store strings unchanged and represent merged positions by cell IDs."""
    table_id = f'contextpack-table-{number:04d}'
    rows = [row for row in table.find_all('tr') if row.find_parent('table') is table]
    occupied: dict[tuple[int, int], str] = {}
    cells, warnings = [], []
    for row_index, row in enumerate(rows):
        column = 0
        for cell in row.find_all(['td', 'th'], recursive=False):
            while (row_index, column) in occupied:
                column += 1
            group_remaining = sum(r.parent is row.parent for r in rows[row_index:])
            row_span = positive_span(cell, 'rowspan', group_remaining, warnings)
            col_span = positive_span(cell, 'colspan', len(rows) - row_index, warnings)
            if len(occupied) + row_span * col_span > MAX_GRID_SLOTS:
                raise ValueError('HTML table exceeds the supported grid size; original has been preserved.')
            cell_id = f'{table_id}-r{row_index + 1}-c{column + 1}'
            for rr in range(row_index, row_index + row_span):
                for cc in range(column, column + col_span):
                    if (rr, cc) in occupied:
                        warnings.append(f'Overlapping spans at row {rr + 1}, column {cc + 1}; original requires review.')
                    else:
                        occupied[rr, cc] = cell_id
            cells.append({'id': cell_id, 'row': row_index + 1, 'column': column + 1,
                          'rowspan': row_span, 'colspan': col_span, 'tag': cell.name,
                          'text': cell.get_text(' ', strip=True),
                          'html': ''.join(safe_inline(child) for child in cell.children),
                          'source_id': cell.get('id'),
                          'source_ids': [t.get('id') or t.get('name') for t in [cell, *cell.find_all(['a', 'span', 'p'])] if t.get('id') or t.get('name')]})
            column += col_span
    columns = max((column + 1 for _, column in occupied), default=0)
    if len(rows) * columns > MAX_GRID_SLOTS:
        raise ValueError('HTML table has excessive dimensions; original has been preserved.')
    preceding = [p.get_text(' ', strip=True) for p in table.find_all_previous(['p', 'h1', 'h2', 'h3'], limit=4) if p.find_parent('table') is not table]
    following = []
    for sibling in table.next_siblings:
        if isinstance(sibling, Tag) and (sibling.name == 'table' or sibling.find('table')):
            break
        text = sibling.get_text(' ', strip=True) if isinstance(sibling, Tag) else str(sibling).strip()
        if text:
            following.append(text)
        if sum(map(len, following)) >= 1200:
            break
    return {'id': table_id, 'source_id': table.get('id'), 'row_source_ids': [row.get('id') for row in rows], 'rows': len(rows), 'columns': columns,
            'caption': table.caption.get_text(' ', strip=True) if table.caption else None,
            'preceding_context': '\n'.join(reversed(preceding))[-1600:],
            'following_context': '\n'.join(following)[:1600],
            'cells': cells, 'grid': [[occupied.get((r, c)) for c in range(columns)] for r in range(len(rows))],
            'warnings': warnings, 'review_status': 'unverified'}


def table_html(table: dict) -> str:
    lines = [f'<a id="{table["id"]}"></a>', '<table>']
    if table['source_id']:
        lines[1] = f'<table id="{escape(table["source_id"], quote=True)}">'
    if table['caption']:
        lines.append(f'<caption>{escape(table["caption"])}</caption>')
    for row in range(1, table['rows'] + 1):
        row_id = table['row_source_ids'][row - 1]
        lines.append(f'<tr id="{escape(row_id, quote=True)}">' if row_id else '<tr>')
        for cell in table['cells']:
            if cell['row'] == row:
                attrs = f' rowspan="{cell["rowspan"]}" colspan="{cell["colspan"]}"'
                if cell['source_id']:
                    attrs += f' id="{escape(cell["source_id"], quote=True)}"'
                lines.append(f'<{cell["tag"]}{attrs}>{cell["html"]}</{cell["tag"]}>')
        lines.append('</tr>')
    return '\n'.join(lines + ['</table>'])


def local_image(source: Path, src: str) -> Path | None:
    parsed = urlsplit(src)
    if parsed.scheme or parsed.netloc:
        return None
    candidate = (source.parent / unquote(parsed.path)).resolve()
    allowed = candidate.is_relative_to(source.parent.resolve()) and candidate.suffix.lower() in {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.webp'}
    return candidate if allowed and candidate.is_file() and candidate.stat().st_size <= 20 * 1024 * 1024 else None


def asset_signature(source: Path) -> str:
    soup = BeautifulSoup(source.read_bytes(), 'html5lib')
    assets = []
    for image in soup.find_all('img'):
        src = str(image.get('src', ''))
        candidate = local_image(source, src)
        assets.append({'src': src, 'sha256': file_hash(candidate) if candidate else None})
    return json_hash(assets)


def prepare_images(soup: BeautifulSoup, source: Path, output: Path, warnings: list[str]) -> list[dict]:
    assets = []
    for image in soup.find_all('img'):
        src = str(image.get('src', ''))
        parsed = urlsplit(src)
        if parsed.scheme or parsed.netloc:
            warnings.append('An external image was not downloaded; see assets.json and the original.')
            assets.append({'source': src, 'status': 'external_not_downloaded'})
            image.replace_with(soup.new_string(f'[External image: {image.get("alt", "")}]'))
            continue
        candidate = local_image(source, src)
        if candidate:
            digest = file_hash(candidate)
            relative = Path('original') / candidate.relative_to(source.parent.resolve()) if (output / 'original').is_dir() else Path('assets') / (digest[:16] + candidate.suffix.lower())
            (output / relative).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(candidate, output / relative)
            if file_hash(output / relative) != digest:
                raise RuntimeError('A referenced image changed while it was being copied')
            image['src'] = relative.as_posix()
            assets.append({'source': src, 'stored_path': relative.as_posix(), 'sha256': digest, 'status': 'copied'})
        else:
            warnings.append('A local image was unavailable or outside the document folder; see assets.json and the original.')
            assets.append({'source': src, 'status': 'not_copied'})
            image.replace_with(soup.new_string(f'[Image requiring original: {image.get("alt", src)}]'))
    return assets


def extract_html(source: Path, output: Path) -> dict:
    output.mkdir(parents=True, exist_ok=True)
    soup = BeautifulSoup(source.read_bytes(), 'html5lib')
    warnings: list[str] = []
    for node in soup.find_all(['script', 'style', 'iframe', 'object', 'embed', 'form', 'link', 'base']):
        node.decompose()
    assets = prepare_images(soup, source, output, warnings)
    anchors = []
    seen = set()
    for node in soup.find_all(True):
        for key in ('id', 'name') if node.name == 'a' else ('id',):
            value = node.get(key)
            if value:
                if value not in seen:
                    anchors.append({'id': value, 'tag': node.name})
                    seen.add(value)
                elif key == 'id' or node.get('id') != value:
                    warnings.append(f'Duplicate source anchor: {value}')
    fragment_links = [unquote(str(a['href'])[1:]) for a in soup.select('a[href^="#"]')]
    missing_links = sorted(set(fragment_links) - seen)
    warnings.extend(f'Source link has no target: #{target}' for target in missing_links)

    # Single-cell layout wrappers must not turn whole chapters into one Markdown row.
    for table in list(soup.find_all('table')):
        rows = [r for r in table.find_all('tr') if r.find_parent('table') is table]
        direct_cells = [c for r in rows for c in r.find_all(['td', 'th'], recursive=False)]
        if len(rows) == 1 and len(direct_cells) == 1 and not table.get('id'):
            cell = direct_cells[0]
            if not cell.get('id') and not rows[0].get('id') and not cell.get('rowspan') and not cell.get('colspan'):
                cell.unwrap()
                rows[0].unwrap()
                for group in list(table.find_all(['tbody', 'thead', 'tfoot'], recursive=False)):
                    group.unwrap()
                table.unwrap()

    tables = [extract_table(table, i) for i, table in enumerate(soup.find_all('table'), 1)]
    warnings.extend(warning for table in tables for warning in table['warnings'])
    replacements: dict[str, str] = {}
    # Replace outer tables first. Nested tables remain represented in tables.json;
    # their contents are preserved as inline text in the outer table's cells.
    for tag, table in zip(list(soup.find_all('table')), tables):
        if tag.parent is None:
            continue
        token = f'CONTEXTPACKTABLETOKEN{len(replacements):08d}END'
        replacements[token] = '\n\n' + table_html(table) + '\n\n'
        tag.replace_with(soup.new_string(token))

    body = soup.body or soup
    for node in list(body.find_all(True)):
        for key in ('id', 'name') if node.name == 'a' else ('id',):
            value = node.get(key)
            if value:
                token = f'CONTEXTPACKANCHORTOKEN{len(replacements):08d}END'
                replacements[token] = f'<a id="{escape(value, quote=True)}"></a>'
                node.insert_before(soup.new_string(token))
                del node[key]
        # Avoid executable link schemes in generated Markdown.
        if node.name == 'a' and node.get('href'):
            if urlsplit(str(node['href']).strip()).scheme.lower() not in {'', 'http', 'https'}:
                del node['href']
    converter = MarkItDown(enable_plugins=False)
    import io
    from markitdown import StreamInfo
    result = converter.convert_stream(io.BytesIO(str(body).encode('utf-8')), stream_info=StreamInfo(extension='.html', charset='utf-8'))
    markdown = result.text_content
    for token, value in replacements.items():
        markdown = markdown.replace(token, value)

    # HTML table cells retain anchors; any targets on table rows/groups are
    # preserved separately so original fragment links always have a destination.
    emitted = set(re.findall(r'<[a-z]+\b[^>]*\b(?:id|name)="([^"]+)"', markdown))
    recovered = [a['id'] for a in anchors if escape(a['id'], quote=True) not in emitted]
    if recovered:
        markdown += '\n\n<!-- Source anchors on table structure; consult anchors.json and the original. -->\n'
        markdown += '\n'.join(f'<a id="{escape(value, quote=True)}"></a>' for value in recovered)
        warnings.append('Some anchors on table structure were preserved at the end of the Markdown; use the original for their exact position.')
    (output / 'source.md').write_text(markdown + '\n', encoding='utf-8')
    for name, data in [('tables.json', {'schema_version': 1, 'source': source.name, 'tables': tables}), ('anchors.json', {'anchors': anchors, 'missing_source_targets': missing_links, 'relocated_targets': recovered}), ('assets.json', assets)]:
        (output / name).write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    metrics = {'table_count': len(tables), 'cell_count': sum(len(t['cells']) for t in tables), 'source_anchor_count': len(anchors),
               'missing_source_targets': missing_links, 'encoding': soup.original_encoding, 'warnings': list(dict.fromkeys(warnings)), 'review_status': 'unverified'}
    (output / 'html-metrics.json').write_text(json.dumps(metrics, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return metrics


def main() -> None:
    parser = argparse.ArgumentParser(description='Preserve HTML table structure, source anchors and local images.')
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path, nargs='?')
    parser.add_argument('--asset-signature', action='store_true')
    args = parser.parse_args()
    if args.asset_signature:
        print(asset_signature(args.source.resolve()))
        return
    if args.output is None:
        parser.error('An output directory is required')
    metrics = extract_html(args.source.resolve(), args.output.resolve())
    print(f'Extracted {metrics["table_count"]} HTML tables and {metrics["source_anchor_count"]} source anchors')


if __name__ == '__main__':
    main()
