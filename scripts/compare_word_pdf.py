"""Compare a native Word PDF export with its LaTeX reference, page by page.

This is a strict visual comparison, independent of the structural Word checks.
No registration, cropping, or resizing is applied to hide layout differences.
"""
import argparse
from difflib import SequenceMatcher
import hashlib
import json
from pathlib import Path
import statistics
import unicodedata

import pymupdf as fitz
from PIL import Image, ImageChops, ImageStat


def characters(page):
    """Keep non-space Unicode characters and their actual PDF baselines."""
    result = []
    for block in page.get_text('rawdict')['blocks']:
        for line in block.get('lines', []):
            for span in line['spans']:
                for char in span['chars']:
                    for value in unicodedata.normalize('NFKC', char['c']):
                        if not value.isspace():
                            result.append((value, char['origin']))
    return result


def text_geometry(reference, rendered):
    a, b = characters(reference), characters(rendered)
    matcher = SequenceMatcher(None, ''.join(c[0] for c in a), ''.join(c[0] for c in b), autojunk=False)
    offsets = []
    for block in matcher.get_matching_blocks():
        for i in range(block.size):
            x, y = a[block.a+i][1]
            xx, yy = b[block.b+i][1]
            offsets.append((xx-x, yy-y))
    return {
        'reference_characters': len(a), 'rendered_characters': len(b),
        'matched_characters': len(offsets),
        'text_sequence_equal': [c[0] for c in a] == [c[0] for c in b],
        'median_absolute_dx_pt': statistics.median(abs(x) for x, y in offsets) if offsets else None,
        'median_absolute_dy_pt': statistics.median(abs(y) for x, y in offsets) if offsets else None,
        'characters_within_half_point': sum(abs(x) <= .5 and abs(y) <= .5 for x, y in offsets),
        'note': 'Sequence-matched baseline diagnostics; raster equality remains the pass criterion.',
    }


def compare(reference, rendered, output, dpi=144):
    reference, rendered, output = map(Path, (reference, rendered, output))
    output.mkdir(parents=True, exist_ok=True)
    report = {
        'reference': str(reference.resolve()),
        'rendered': str(rendered.resolve()),
        'reference_sha256': hashlib.sha256(reference.read_bytes()).hexdigest(),
        'rendered_sha256': hashlib.sha256(rendered.read_bytes()).hexdigest(),
        'dpi': dpi,
        'scope': 'Exact raster equality at the stated DPI; no alignment correction.',
        'pages': [],
    }
    with fitz.open(reference) as source, fitz.open(rendered) as actual:
        report['page_counts'] = [len(source), len(actual)]
        for index in range(max(len(source), len(actual))):
            record = {'page': index + 1, 'equal': False}
            report['pages'].append(record)
            if index >= min(len(source), len(actual)):
                record['missing_in'] = 'reference' if index >= len(source) else 'rendered'
                continue
            a, b = source[index], actual[index]
            record['page_sizes_pt'] = [list(a.rect), list(b.rect)]
            record['text_geometry'] = text_geometry(a, b)
            images = []
            for label, page in [('reference', a), ('word', b)]:
                pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False)
                img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                img.save(output / f'page-{index+1:02d}-{label}.png')
                images.append(img)
            first, second = images
            if first.size != second.size:
                record['raster_size_mismatch'] = [first.size, second.size]
                continue
            difference = ImageChops.difference(first, second)
            record['equal'] = difference.getbbox() is None and a.rect == b.rect
            record['mean_absolute_channel_error'] = sum(ImageStat.Stat(difference).mean) / 3
            red, green, blue = difference.split()
            mask = ImageChops.lighter(ImageChops.lighter(red, green), blue).point(lambda value: 255 if value else 0)
            record['changed_pixel_fraction'] = mask.histogram()[255] / (first.width * first.height)
            # Red marks all differing pixels at their original page positions.
            overlay = first.convert('L').convert('RGB')
            overlay.paste((255, 0, 0), mask=mask)
            overlay.save(output / f'page-{index+1:02d}-difference.png')
            pair = Image.new('RGB', (first.width * 2, first.height), 'white')
            pair.paste(first, (0, 0)); pair.paste(second, (first.width, 0))
            pair.save(output / f'page-{index+1:02d}-pair.png')
    report['passed'] = all(p['equal'] for p in report['pages']) and bool(report['pages'])
    (output / 'comparison.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('reference', type=Path)
    parser.add_argument('rendered', type=Path)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--dpi', type=int, default=144)
    args = parser.parse_args()
    if not 72 <= args.dpi <= 600:
        parser.error('--dpi must be between 72 and 600')
    result = compare(args.reference, args.rendered, args.output, args.dpi)
    print(json.dumps({'passed': result['passed'], 'page_counts': result['page_counts'],
                      'different_pages': [p['page'] for p in result['pages'] if not p['equal']]}))
    raise SystemExit(0 if result['passed'] else 1)
