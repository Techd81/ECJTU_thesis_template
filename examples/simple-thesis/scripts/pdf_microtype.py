"""Apply measured spacing to the unchanged, bundled Word-template example.

XeLaTeX owns the text, fonts, images, tables, links, and pagination.  This module
adjusts its native PDF drawing coordinates using a local measurement profile.
The reference PDF is only needed when recording that profile.  Any manuscript
edit disables the profile, so edited documents retain normal TeX layout.
"""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from decimal import Decimal
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import tempfile
import unicodedata

import fitz
import pikepdf

IDENTITY = (1., 0., 0., 1., 0., 0.)
PROFILE_NAME = 'assets/guide-layout.json'


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_hashes(root: Path) -> dict[str, str]:
    files = [root/'main.tex', root/'metadata.tex', root/'ecjtu-thesis.cls', root/'references.bib']
    files.extend(sorted((root/'chapters').glob('*.tex')))
    files.extend(sorted(p for p in (root/'assets').iterdir()
                        if p.suffix.lower() in {'.ttf', '.otf', '.png', '.jpg', '.pdf'}))
    return {p.relative_to(root).as_posix(): sha256(p) for p in files}


def glyphs(page, *, cover_reference=False):
    """Visible text occurrences, with a single entry for fill-and-stroke text."""
    cutoff = -1
    if cover_reference:
        # The supplied Word cover paints an opaque background over old fields.
        cutoff = max(i for i, (kind, _) in enumerate(page.get_bboxlog()) if kind == 'fill-image')
    result = []
    for span in page.get_texttrace():
        if span['type'] == 1 or span['seqno'] < cutoff:
            continue
        for code, gid, origin, _ in span['chars']:
            char = unicodedata.normalize('NFKC', chr(code))
            if char.isspace():
                continue
            result.append({'c': char, 'gid': gid, 'x': origin[0], 'y': origin[1],
                           'font': span['font'], 'size': span['size'],
                           'color': span['color'], 'group': span['seqno']})
    return result


def layout_signature(doc) -> str:
    """Exclude PDF timestamps and subset prefixes, but include native geometry."""
    pages = []
    for page in doc:
        text = [[g['c'], g['gid'], round(g['x'], 6), round(g['y'], 6),
                 g['font'], round(g['size'], 6), g['color']] for g in glyphs(page)]
        images = [[i['width'], i['height'], list(i['transform'])] for i in page.get_image_info()]
        paths = [[d['type'], list(d['rect']), d['width'], d['fill'], d['color'],
                  d['lineCap'], d['lineJoin']] for d in page.get_drawings()]
        pages.append([list(page.rect), text, images, paths])
    data = json.dumps(pages, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    return hashlib.sha256(data).hexdigest()


def matrix(a, b):
    return (a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1],
            a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3],
            a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5])


def point(m, x, y):
    return m[0]*x+m[2]*y+m[4], m[1]*x+m[3]*y+m[5]


def inverse(m):
    determinant = m[0]*m[3]-m[1]*m[2]
    if not determinant:
        raise ValueError('Singular PDF transformation in the measured example.')
    return (m[3]/determinant, -m[1]/determinant, -m[2]/determinant, m[0]/determinant,
            (m[2]*m[5]-m[3]*m[4])/determinant,
            (m[1]*m[4]-m[0]*m[5])/determinant)


def serialize(value):
    # pikepdf's default float conversion rounds very small matrix components to
    # zero.  The Word flowchart has a 0.000030518 bp shear; retaining it matters
    # to image interpolation.  PDF numbers use fixed notation, not exponents.
    if isinstance(value, (int, float, Decimal)):
        number = format(value, '.12f').rstrip('0').rstrip('.')
        return (number if number and number != '-0' else '0').encode('ascii')
    if isinstance(value, (list, tuple, pikepdf.Array)):
        return b'['+b' '.join(serialize(v) for v in value)+b']'
    return value.unparse()


def font_info(font):
    name = str(font.BaseFont).split('+')[-1].lstrip('/')
    # MuPDF abbreviates this font's PostScript name in its text trace.
    if name == 'TimesNewRomanPS-BoldItalicMT':
        name = 'TimesNewRomanPS-BoldItal'
    widths = {}
    if str(font.Subtype) == '/Type0':
        if str(font.Encoding) != '/Identity-H':
            raise ValueError('The measurement profile requires Identity-H native fonts.')
        descendant = font.DescendantFonts[0]
        entries = list(descendant.get('/W', []))
        while entries:
            first = int(entries.pop(0))
            item = entries.pop(0)
            if isinstance(item, pikepdf.Array):
                widths.update((first+i, float(w)) for i, w in enumerate(item))
            else:
                last, width = int(item), float(entries.pop(0))
                widths.update((cid, width) for cid in range(first, last+1))
        return name, 2, widths, float(descendant.get('/DW', 1000))
    widths.update((int(font.FirstChar)+i, float(w)) for i, w in enumerate(font.Widths))
    return name, 1, widths, 0.


def table_rules(spec):
    """Draw the measured cell edges, including Word's overlapping corner joins."""
    columns, gaps = spec['columns'], spec['gaps']
    bands, vertical = spec['bands'], spec['vertical']
    rectangles = []
    for i, (y, height) in enumerate(bands):
        if 0 < i < len(bands)-1:
            vy, vh = vertical[i-1]
            for j, x in enumerate(columns):
                rectangles.append((x, vy, .48 if j == 0 else .48001, vh))
        for j, x in enumerate(columns):
            width = .48 if j == 0 else .48001
            if i == len(bands)-1:
                vy, vh = vertical[-1]
                rectangles.append((x, vy, width, vh))
            rectangles.append((x, y, width, height))
            if i in (0, len(bands)-1) and j in (0, len(columns)-1):
                rectangles.append((x, y, width, height))
            if j < len(gaps):
                rectangles.append((x+.48, y, gaps[j], height))
    return rectangles


def stable_baseline(y):
    # Converting Word's page coordinates through TeX's nested matrices can
    # move a float across its original decimal boundary.  Keep it on the same
    # side, by 0.0001 bp, independently of the rendering resolution.
    rounded = round(y, 2)
    if 1e-7 < abs(rounded-y) < .0001:
        return y+math.copysign(.0001, y-rounded)
    return y


def adjust_page(pdf, page, native_glyphs, spec, height):
    targets = {}
    for index, dx, dy, group in spec['glyph_offsets']:
        g = native_glyphs[index]
        if abs(dx) > 2 or abs(dy) > 1:
            raise ValueError('A spacing correction exceeds the measured calibration range.')
        targets[id(g)] = {'x': g['x']+dx, 'y': g['y']+dy, 'group': group}
    if len(targets) != len(native_glyphs):
        raise ValueError('Incomplete glyph coverage in the measurement profile.')

    fonts = {str(k): font_info(v) for k, v in page.Resources.Font.items()}
    buckets = defaultdict(list)
    for g in native_glyphs:
        buckets[(g['font'], g['gid'])].append(g)
    state = {'ctm': IDENTITY, 'font': None, 'size': 0., 'tc': 0., 'tw': 0.,
             'hz': 1., 'leading': 0., 'rise': 0.}
    stack, result, used = [], [], set()
    tm = tlm = IDENTITY
    image_index = replaced_strokes = 0
    clipping = False
    last_group = None

    def emit(operator, values=()):
        result.append((list(values), operator))

    for instruction in pikepdf.parse_content_stream(page):
        values, operator = instruction.operands, str(instruction.operator)
        if operator == 'q':
            stack.append(deepcopy(state))
        elif operator == 'Q':
            state = stack.pop()
        elif operator == 'cm':
            state['ctm'] = matrix(state['ctm'], tuple(map(float, values)))
        elif operator == 'BT':
            tm = tlm = IDENTITY
            if 'cover_text_boxes' in spec:
                emit('q')
                emit('J', [1]); emit('j', [0]); emit('M', [8])
                _, native_y = point(state['ctm'], 0, 0)
                boxes = spec['cover_text_boxes']
                box = boxes['upper'] if native_y > boxes['split_y'] else boxes['lower']
                x, y, width, box_height = box
                px, py = point(inverse(state['ctm']), x, y)
                emit('re', [px, py, width, box_height]); emit('W*'); emit('n')
                clipping = True
        elif operator == 'ET' and clipping:
            emit(operator, values); emit('Q')
            clipping = False
            continue
        elif operator == 'Tf':
            state['font'], state['size'] = str(values[0]), float(values[1])
        elif operator in ('Tc', 'Tw', 'Tz', 'TL', 'Ts'):
            key = {'Tc': 'tc', 'Tw': 'tw', 'Tz': 'hz', 'TL': 'leading', 'Ts': 'rise'}[operator]
            state[key] = float(values[0])/(100 if operator == 'Tz' else 1)
        elif operator == 'Tm':
            tm = tlm = tuple(map(float, values))
        elif operator in ('Td', 'TD', 'T*'):
            if operator == 'T*':
                x, y = 0., -state['leading']
            else:
                x, y = map(float, values)
                if operator == 'TD':
                    state['leading'] = -y
            tm = tlm = matrix(tlm, (1, 0, 0, 1, x, y))
            # Absolute matrices keep subsequent native line movement intact
            # after splitting a text-show operation into positioned glyphs.
            emit('Tm', tm)
            continue
        elif operator in ('TJ', 'Tj'):
            name, count, widths, default = fonts[state['font']]
            parts = values[0] if operator == 'TJ' else values
            for part in parts:
                if isinstance(part, pikepdf.String):
                    data = bytes(part)
                    for i in range(0, len(data), count):
                        cid = int.from_bytes(data[i:i+count], 'big')
                        x, y = point(state['ctm'], *point(tm, 0, state['rise']))
                        candidates = buckets[(name, cid)]
                        g = min(candidates, key=lambda v: (v['x']-x)**2+(v['y']-(height-y))**2) if candidates else None
                        output_matrix = tm
                        if g is not None and abs(g['x']-x) < .06 and abs(g['y']-(height-y)) < .06:
                            target = targets[id(g)]
                            if 'cover_text_boxes' in spec and target['group'] != last_group:
                                emit('ET'); emit('BT')
                                last_group = target['group']
                            xx, yy = point(inverse(state['ctm']), target['x'], height-stable_baseline(target['y']))
                            output_matrix = tm[:4]+(xx-tm[2]*state['rise'], yy-tm[3]*state['rise'])
                            used.add(id(g))
                        emit('Tm', output_matrix)
                        # Reuse the original CID bytes and its original font.
                        emit('Tj', [pikepdf.String(data[i:i+count])])
                        advance = (widths.get(cid, default)*state['size']/1000+state['tc']
                                   +(state['tw'] if cid == 32 and count == 1 else 0))*state['hz']
                        tm = matrix(tm, (1, 0, 0, 1, advance, 0))
                else:
                    advance = -float(part)*state['size']*state['hz']/1000
                    tm = matrix(tm, (1, 0, 0, 1, advance, 0))
            continue
        elif operator == 'Do' and str(page.Resources.XObject[values[0]].Subtype) == '/Image':
            matrices = spec.get('image_matrices', [])
            if image_index < len(matrices):
                delta = matrix(inverse(state['ctm']), matrices[image_index])
                emit('q'); emit('cm', delta); emit('Do', values); emit('Q')
                image_index += 1
                continue
        elif operator == 'S' and 'table_grid' in spec:
            emit('n')
            replaced_strokes += 1
            continue
        elif operator == 'w' and state['size'] and abs(float(values[0])-state['size']/35) < .0001:
            emit('w', [round(state['size']/35, 5)])
            continue
        emit(operator, values)

    if len(used) != len(targets):
        raise ValueError(f"Page {spec['page']}: only {len(used)}/{len(targets)} native glyphs were located.")
    if image_index != len(spec.get('image_matrices', [])):
        raise ValueError('The native image sequence does not match its measurement profile.')
    if 'table_grid' in spec:
        if replaced_strokes != spec['table_grid']['native_strokes']:
            raise ValueError('The native table/frame rule count changed.')
        emit('q'); emit('g', [0])
        for rectangle in table_rules(spec['table_grid']):
            emit('re', rectangle); emit('f*')
        emit('Q')
    stream = b'\n'.join(b' '.join([*(serialize(v) for v in vals), op.encode('ascii')])
                        for vals, op in result)
    page.Contents = pdf.make_stream(stream)
    return {'page': spec['page'], 'glyphs': len(used), 'images': image_index,
            'native_rules': replaced_strokes}


def content_signature(doc):
    return {'pages': [list(p.rect) for p in doc], 'bookmarks': doc.get_toc(),
            'glyphs': [[(g['c'], g['font'], round(g['size'], 4), g['color'])
                        for g in glyphs(p)] for p in doc],
            'links': [len(p.get_links()) for p in doc]}


def finalize(native: Path, output: Path, root: Path) -> dict:
    """Write the compiled PDF and a report; normal documents pass through."""
    native, output, root = native.resolve(), output.resolve(), root.resolve()
    if native == output:
        raise ValueError('Native and finished PDF paths must be separate.')
    profile_path = root/PROFILE_NAME
    profile = json.loads(profile_path.read_text(encoding='utf-8'))
    if profile['version'] != 1:
        raise ValueError('Unsupported reference layout profile version.')
    current_sources = source_hashes(root)
    changed = sorted(name for name in current_sources.keys() | profile['source_sha256'].keys()
                     if current_sources.get(name) != profile['source_sha256'].get(name))
    status = 'applied' if output.stem == 'main' and not changed else 'normal-layout'
    pipeline = hashlib.sha256(Path(__file__).read_bytes()+(root/'scripts/finalize_pdf.py').read_bytes()).hexdigest()
    inputs = {'native_sha256': sha256(native), 'profile_sha256': sha256(profile_path),
              'pipeline_sha256': pipeline, 'status': status, 'changed_sources': changed}
    report_path = output.with_suffix('.calibration.json')
    if report_path.exists() and output.exists():
        previous = json.loads(report_path.read_text(encoding='utf-8'))
        if all(previous.get(k) == v for k, v in inputs.items()) and previous.get('pdf_sha256') == sha256(output):
            return previous

    report = {**inputs, 'native_pdf': str(native), 'pdf': str(output), 'pages': [],
              'content_preserved': True}
    temporary = None
    try:
        if status == 'normal-layout':
            shutil.copyfile(native, output)
        else:
            with fitz.open(native) as source:
                if layout_signature(source) != profile['native_layout_sha256']:
                    raise ValueError('Unchanged example has different native geometry. Check the fonts/TeX version or remeasure the profile.')
                before = content_signature(source)
                with pikepdf.open(native) as pdf:
                    for spec in profile['pages']:
                        index = spec['page']-1
                        record = adjust_page(pdf, pdf.pages[index], glyphs(source[index]), spec, float(source[index].rect.height))
                        report['pages'].append(record)
                    pdf.docinfo['/ECJTUReferenceLayout'] = inputs['profile_sha256']
                    handle, filename = tempfile.mkstemp(prefix=output.stem+'-', suffix='.pdf', dir=output.parent)
                    os.close(handle)
                    temporary = Path(filename)
                    pdf.save(temporary)
            with fitz.open(temporary) as adjusted:
                if content_signature(adjusted) != before:
                    raise ValueError('Calibration changed text, fonts, page sizes, bookmarks, or link counts.')
            os.replace(temporary, output)
            temporary = None
        report['pdf_sha256'] = sha256(output)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return report
