"""Record numeric spacing measurements for the bundled LaTeX example.

Run explicitly after intentionally changing the template's default layout.
Recording reads both PDFs and writes a profile; it never edits either PDF.
"""
import argparse
from difflib import SequenceMatcher
import json
from pathlib import Path

import fitz

from pdf_microtype import glyphs, layout_signature, sha256, source_hashes

ROOT = Path(__file__).resolve().parents[1]


def rows(items):
    grouped = []
    for glyph in sorted(items, key=lambda g: (g['y'], g['x'])):
        if grouped and abs(glyph['y']-grouped[-1][0]['y']) < .3:
            grouped[-1].append(glyph)
        else:
            grouped.append([glyph])
    return [sorted(row, key=lambda g: g['x']) for row in grouped]


def align(reference, native):
    reference_rows, native_rows = rows(reference), rows(native)
    mapped = {}
    for actual_row in native_rows:
        reference_row = min(reference_rows, key=lambda row: abs(row[0]['y']-actual_row[0]['y']))
        if abs(reference_row[0]['y']-actual_row[0]['y']) > 1:
            continue
        matcher = SequenceMatcher(None, [g['c'] for g in reference_row], [g['c'] for g in actual_row], autojunk=False)
        for block in matcher.get_matching_blocks():
            for i in range(block.size):
                r, a = reference_row[block.a+i], actual_row[block.b+i]
                if (abs(a['x']-r['x']) <= 2 and abs(a['size']-r['size']) < .02
                        and a['font'] == r['font'] and a['color'] == r['color']):
                    mapped[id(a)] = r
    if len(mapped) != len(native):
        raise ValueError(f'Only {len(mapped)}/{len(native)} glyphs match; fix text/style/reflow before recording.')
    return mapped


def measured_grid(page):
    # bp measurements of the supplied Word declaration and salary table.
    if page == 3:
        return {'native_strokes': 1, 'columns': [119.06, 490.18], 'gaps': [370.63],
                'bands': [[770.52, .48], [145.94, .47998]], 'vertical': [[146.42, 624.1]]}
    if page == 15:
        return {'native_strokes': 25, 'columns': [92.304, 151.70, 223.73, 367.27, 475.78],
                'gaps': [58.92, 71.544, 143.06, 108.02],
                'bands': [[322.01, .47998], [301.97, .48004], [262.46, .47998],
                          [222.98, .48004], [183.38, .47998]],
                'vertical': [[302.45, 19.56], [262.94, 39.024], [223.46, 39], [183.86, 39.12]]}
    return None


def record(reference, native, output):
    report = {'version': 1, 'units': 'PDF bp (1/72 inch)',
              'scope': 'Unchanged bundled guide example; native text and graphic objects.',
              'reference_sha256': sha256(reference), 'source_sha256': source_hashes(ROOT),
              'pages': [], 'native_equal_pages': []}
    with fitz.open(reference) as ref, fitz.open(native) as tex:
        if len(ref) != len(tex):
            raise ValueError('Page counts differ; fix pagination before recording.')
        report['native_layout_sha256'] = layout_signature(tex)
        report['page_count'] = len(tex)
        for i, (rpage, tpage) in enumerate(zip(ref, tex)):
            if rpage.rect != tpage.rect:
                raise ValueError(f'Page {i+1} has a different size.')
            first, second = (p.get_pixmap(dpi=144, alpha=False) for p in (rpage, tpage))
            if first.samples == second.samples:
                report['native_equal_pages'].append(i+1)
            actual = glyphs(tpage)
            if not actual:
                continue
            mapped = align(glyphs(rpage, cover_reference=i == 0), actual)
            offsets = [[j, round(mapped[id(g)]['x']-g['x'], 12), round(mapped[id(g)]['y']-g['y'], 12),
                        mapped[id(g)]['group'] if i == 0 else 0] for j, g in enumerate(actual)]
            spec = {'page': i+1, 'glyph_offsets': offsets, 'image_matrices': []}
            for info in rpage.get_image_info():
                a, b, c, d, e, f = info['transform']
                target = (a, -b, -c, d, e+c, float(tpage.rect.height)-f-d)
                spec['image_matrices'].append([round(v, 5) if j != 1 else v for j, v in enumerate(target)])
            if i == 0:
                spec['cover_text_boxes'] = {'split_y': 500, 'upper': [77.664, 634.54, 462.1, 177.74],
                                            'lower': [74.904, 163.1, 414.07, 101.16]}
            grid = measured_grid(i+1)
            if grid:
                spec['table_grid'] = grid
            report['pages'].append(spec)
            print(f'Page {i+1}: {len(offsets)} matched glyphs', flush=True)
    # One compact data line per page keeps the measured offsets inspectable.
    pages = report.pop('pages')
    header = json.dumps(report, ensure_ascii=False, indent=2).rstrip()[:-1]
    body = ',\n'.join('    '+json.dumps(page, ensure_ascii=False, separators=(',', ':')) for page in pages)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(header.rstrip()+',\n  "pages": [\n'+body+'\n  ]\n}\n', encoding='utf-8')
    print('Recorded numeric profile: '+str(output))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--reference', required=True, type=Path)
    parser.add_argument('--native', default=ROOT/'build/main.native.pdf', type=Path)
    parser.add_argument('--output', default=ROOT/'assets/guide-layout.json', type=Path)
    args = parser.parse_args()
    record(args.reference.resolve(), args.native.resolve(), args.output.resolve())
