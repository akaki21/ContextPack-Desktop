from __future__ import annotations

import argparse
from contextlib import closing
import csv
import json
import math
from pathlib import Path
import subprocess
import tempfile
import unicodedata

import pypdfium2 as pdfium


def normalize_text(text: str) -> str:
    return text.replace('\r\n', '\n').replace('\r', '\n').strip()


def text_needs_ocr(text: str) -> bool:
    text = text.strip()
    if len(text) < 40:
        return True
    damaged = sum(c == '\ufffd' or unicodedata.category(c) == 'Co' for c in text)
    return damaged / len(text) > 0.01


def render_page(page, destination: Path, dpi: int, max_pixels: int = 24_000_000) -> int:
    width, height = page.get_size()
    scale = min(dpi / 72, math.sqrt(max_pixels / max(1, width * height)))
    bitmap = page.render(scale=scale)
    try:
        image = bitmap.to_pil().convert('RGB')
        try:
            image.save(destination, format='PNG', optimize=True)
        finally:
            image.close()
    finally:
        bitmap.close()
    return max(1, round(scale * 72))


def ocr_page(page, *, output: Path, number: int, tesseract: str, tessdata: str,
             languages: str, timeout: int = 120) -> tuple[str, dict]:
    output.mkdir(exist_ok=True)
    base = output / f'page-{number:03d}'
    info: dict = {'attempted': True, 'status': 'failed', 'languages': languages,
                  'method': 'rendered_page_tesseract', 'review_required': True}
    with tempfile.TemporaryDirectory(prefix='.ocr-page-', dir=output.parent) as temporary:
        png = Path(temporary) / 'page.png'
        dpi = render_page(page, png, 300)
        info['dpi'] = dpi
        command = [tesseract, str(png), str(base), '--tessdata-dir', tessdata, '-l', languages, '--dpi', str(dpi), 'txt', 'tsv']
        try:
            result = subprocess.run(command, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=timeout)
        except (OSError, subprocess.TimeoutExpired) as error:
            info['error'] = str(error)
            return '', info
        if result.returncode:
            info['error'] = result.stderr.strip()[-1500:] or f'Tesseract exit code {result.returncode}'
            return '', info
    text_file = base.with_suffix('.txt')
    text = normalize_text(text_file.read_text(encoding='utf-8')) if text_file.exists() else ''
    info.update({'status': 'extracted' if text else 'empty', 'text_file': text_file.name if text_file.exists() else None})
    tsv = base.with_suffix('.tsv')
    if tsv.exists():
        with tsv.open(encoding='utf-8', newline='') as stream:
            words = [r for r in csv.DictReader(stream, delimiter='\t') if r.get('level') == '5' and r.get('text', '').strip()]
        confidences = [float(w['conf']) for w in words if w.get('conf') is not None and float(w['conf']) >= 0]
        info.update({'word_boxes': tsv.name, 'mean_word_confidence': round(sum(confidences) / len(confidences), 2) if confidences else None,
                     'low_confidence_words': sum(c < 80 for c in confidences)})
    return text, info


def extract_pdf(source: Path, destination: Path, *, dpi: int = 180, page_text: Path | None = None,
                metrics_path: Path | None = None, max_pages: int = 0, render_mode: str = 'All',
                max_render_pages: int = 0, tesseract: str | None = None, tessdata: str = '',
                languages: str = 'kat+eng', cancel_file: Path | None = None) -> dict:
    if not 72 <= dpi <= 600 or max_pages < 0 or max_render_pages < 0:
        raise ValueError('Invalid DPI or page limit')
    if render_mode not in {'All', 'Sparse', 'None'}:
        raise ValueError('Unknown PDF render mode')
    destination.mkdir(parents=True, exist_ok=True)
    parts = [f'# Page-aware text — {source.name}\n']
    metrics = []
    rendered = 0
    with pdfium.PdfDocument(str(source)) as document:
        count = len(document)
        safety_skip = bool(max_pages and count > max_pages)
        digits = max(3, len(str(count)))
        for index in range(count):
            if cancel_file and cancel_file.exists():
                raise InterruptedError('The job was cancelled between PDF pages.')
            with closing(document[index]) as page:
                with closing(page.get_textpage()) as text_page:
                    text = normalize_text(text_page.get_text_range())
                original_count = len(text)
                needs_ocr = text_needs_ocr(text)
                ocr = {'attempted': False, 'status': 'not_requested'}
                if tesseract and needs_ocr:
                    recovered, ocr = ocr_page(page, output=destination.parent / 'ocr-pages', number=index + 1,
                                              tesseract=tesseract, tessdata=tessdata, languages=languages)
                    if recovered and (len(recovered) >= len(text) or not text_needs_ocr(recovered)):
                        text = recovered
                        ocr['used_for_page_text'] = True
                    else:
                        ocr['used_for_page_text'] = False
                selected = render_mode == 'All' or (render_mode == 'Sparse' and needs_ocr)
                image_name = None
                rendered_dpi = None
                if selected and not safety_skip and (max_render_pages == 0 or rendered < max_render_pages):
                    image_name = f'page-{index + 1:0{digits}d}.png'
                    rendered_dpi = render_page(page, destination / image_name, dpi)
                    rendered += 1
                sparse = text_needs_ocr(text)
                metrics.append({'page': index + 1, 'characters': len(text), 'input_text_characters': original_count,
                                'words': len(text.split()), 'sparse_text': sparse, 'ocr': ocr,
                                'image': image_name, 'image_dpi': rendered_dpi, 'review_required': sparse or ocr['attempted']})
                parts.extend([f'<!-- source-page: {index + 1} -->', f'## Page {index + 1}', '',
                              text if text else '_No extractable text on this page. Review the original._', ''])
                print(f'Processed PDF page {index + 1}/{count}; characters={len(text)}; OCR={ocr["status"]}', flush=True)
    if page_text:
        page_text.write_text('\n'.join(parts), encoding='utf-8')
    payload = {'source': source.name, 'page_count': count, 'rendered_pages': rendered,
               'render_skipped': rendered == 0 and count > 0, 'render_mode': render_mode,
               'max_pages': max_pages or None, 'max_render_pages': max_render_pages or None,
               'total_characters': sum(p['characters'] for p in metrics),
               'sparse_pages': [p['page'] for p in metrics if p['sparse_text']],
               'fallback_ocr_pages': [p['page'] for p in metrics if p['ocr']['attempted']],
               'ocr_failed_pages': [p['page'] for p in metrics if p['ocr']['status'] in {'empty', 'failed'}],
               'pages': metrics, 'review_status': 'unverified'}
    if safety_skip:
        payload['reason'] = f'PDF exceeds the legacy {max_pages}-page PNG safety limit; all page text was still extracted.'
    if metrics_path:
        metrics_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Rendered {rendered} pages at up to {dpi} DPI', flush=True)
    return payload


def main() -> None:
    parser = argparse.ArgumentParser(description='Extract every PDF page; bound optional images and recover sparse pages.')
    parser.add_argument('input_pdf', type=Path)
    parser.add_argument('output_dir', type=Path)
    parser.add_argument('--dpi', type=int, default=180)
    parser.add_argument('--page-text', type=Path)
    parser.add_argument('--metrics', type=Path)
    parser.add_argument('--max-pages', type=int, default=0, help='Legacy all-or-none PNG safety limit; text is never skipped')
    parser.add_argument('--max-render-pages', type=int, default=0, help='Maximum saved PNGs, 0 for unlimited')
    parser.add_argument('--render-mode', choices=['All', 'Sparse', 'None'], default='All')
    parser.add_argument('--tesseract')
    parser.add_argument('--tessdata', default='')
    parser.add_argument('--languages', default='kat+eng')
    parser.add_argument('--cancel-file', type=Path)
    args = parser.parse_args()
    try:
        extract_pdf(args.input_pdf.resolve(), args.output_dir.resolve(), dpi=args.dpi, page_text=args.page_text,
                    metrics_path=args.metrics, max_pages=args.max_pages, render_mode=args.render_mode,
                    max_render_pages=args.max_render_pages, tesseract=args.tesseract, tessdata=args.tessdata,
                    languages=args.languages, cancel_file=args.cancel_file)
    except InterruptedError as error:
        print(str(error))
        raise SystemExit(2)


if __name__ == '__main__':
    main()
