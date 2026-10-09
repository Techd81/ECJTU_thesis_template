"""PDF 布局解析适配：保留词间空格、图形和可扩展的 Word 正文。"""
from __future__ import annotations

from contextlib import contextmanager
from copy import deepcopy

import pymupdf as fitz
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pdf2docx.page.RawPageFactory import RawPageFactory
from pdf2docx.page.RawPageFitz import RawPageFitz


def inside(bbox, rect):
    box = fitz.Rect(bbox)
    return fitz.Point((box.x0 + box.x1)/2, (box.y0 + box.y1)/2) in fitz.Rect(rect)


def figure_regions(page, include_images=False):
    """将含曲线/斜线的连通图形整体保留；横竖网格仍交给表格解析器。"""
    drawings = page.get_drawings()
    regions = []
    for rect in page.cluster_drawings(drawings=drawings):
        if rect.width < 40 or rect.height < 40:
            continue
        if rect.get_area() > page.rect.get_area() * .7:
            continue
        complex_path = any(
            item[0] == 'c' or (item[0] == 'l'
                and abs(item[1].x-item[2].x) > .5
                and abs(item[1].y-item[2].y) > .5)
            for drawing in drawings if inside(drawing['rect'], rect)
            for item in drawing['items'])
        if complex_path:
            regions.append(tuple((rect + (-1, -1, 1, 1)) & page.rect))
    if include_images:
        regions.extend(tuple(info['bbox']) for info in page.get_image_info()
                       if 40 < info['bbox'][2]-info['bbox'][0] < page.rect.width-40
                       and 40 < info['bbox'][3]-info['bbox'][1] < page.rect.height-40)
    return regions


def page_text(page, regions=()):
    """读取文字字符；图形内文字单独记录在核验清单中。"""
    parts = []
    for block in page.get_text('rawdict', flags=fitz.TEXTFLAGS_RAWDICT)['blocks']:
        for line in block.get('lines', []):
            for span in line['spans']:
                parts.extend(c['c'] for c in span['chars']
                             if not any(inside(c['bbox'], r) for r in regions))
    return ''.join(parts)


class EditableRawPage(RawPageFitz):
    def _preprocess_text(self, **settings):
        blocks = super()._preprocess_text(**settings)
        regions = settings['word_figures'].get(self.page_engine.number, [])
        for block in blocks:
            lines = []
            for line in block['lines']:
                spans = []
                for span in line['spans']:
                    span['chars'] = [c for c in span['chars']
                                     if not any(inside(c['bbox'], r) for r in regions)]
                    if not span['chars']:
                        continue
                    # pdf2docx 0.5 丢弃独立的空白 span。XeLaTeX 的词间空格
                    # 常以这种形式编码；并回前一 span，保持字符和实际字宽。
                    if spans and all(c['c'].isspace() for c in span['chars']):
                        previous = spans[-1]
                        previous['chars'].extend(span['chars'])
                        previous['bbox'] = tuple(fitz.Rect(previous['bbox']) | fitz.Rect(span['bbox']))
                    else:
                        spans.append(span)
                if spans:
                    line['spans'] = spans
                    lines.append(line)
            block['lines'] = lines
        return [b for b in blocks if b['lines']]

    def extract_raw_dict(self, **settings):
        raw = super().extract_raw_dict(**settings)
        regions = settings['word_figures'].get(self.page_engine.number, [])
        for rect in regions:
            raw['shapes'] = [s for s in raw['shapes'] if not inside(
                s['bbox'] if 'bbox' in s else (*s['start'], *s['end']), rect)]
            raw['blocks'] = [b for b in raw['blocks']
                             if b.get('type') != 1 or not inside(b['bbox'], rect)]
            pix = self.page_engine.get_pixmap(clip=fitz.Rect(rect),
                matrix=fitz.Matrix(settings['clip_image_res_ratio'], settings['clip_image_res_ratio']),
                alpha=False)
            raw['blocks'].append({'type': 1, 'bbox': rect, 'ext': 'png',
                'width': pix.width, 'height': pix.height, 'image': pix.tobytes('png')})
        return raw


@contextmanager
def editable_pdf_parser():
    # pdf2docx 提供的后端注册点；转换结束后恢复，避免污染后续调用。
    previous = RawPageFactory.MAP['PYMUPDF']
    RawPageFactory.MAP['PYMUPDF'] = EditableRawPage
    try:
        yield
    finally:
        RawPageFactory.MAP['PYMUPDF'] = previous


def normalize_editing(doc):
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.25
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), 'SimSun')
    for height in doc.element.xpath('//w:trHeight'):
        height.set(qn('w:hRule'), 'atLeast')
    for p in doc.element.xpath('//w:p'):
        text_runs = p.xpath('./w:r[w:t]')
        if not text_runs:
            continue
        props = p.get_or_add_pPr()
        # Enter 新段落和插入文字继承正常字号、颜色和中英文字体。
        mark = props.find(qn('w:rPr'))
        if mark is not None:
            props.remove(mark)
        mark = deepcopy(text_runs[-1].get_or_add_rPr())
        for child in list(mark):
            if child.tag in (qn('w:spacing'), qn('w:w')):
                mark.remove(child)
        fonts = mark.find(qn('w:rFonts'))
        if fonts is not None:
            fonts.set(qn('w:eastAsia'), 'SimSun')
        props.append(mark)
    for r in doc.element.xpath('//w:r[w:t]'):
        pr = r.get_or_add_rPr()
        fonts = pr.find(qn('w:rFonts'))
        if fonts is not None and fonts.get(qn('w:eastAsia'), '').startswith('Times'):
            fonts.set(qn('w:eastAsia'), 'SimSun')
    compat = doc.settings.element.find(qn('w:compat'))
    for item in list(compat):
        if item.get(qn('w:name')) == 'compatibilityMode':
            compat.remove(item)
    item = OxmlElement('w:compatSetting')
    for key, value in {'name': 'compatibilityMode',
                       'uri': 'http://schemas.microsoft.com/office/word', 'val': '15'}.items():
        item.set(qn('w:' + key), value)
    compat.append(item)


def restore_floating_image_sizes(doc, source):
    """保留 PDF 中背景图的实际缩放比例，避免表单横线相对文字上移。"""
    page_index = 0
    if len(doc.sections) != len(source):
        return
    for child in doc.element.body:
        for anchor in child.xpath('.//wp:anchor'):
            horizontal = anchor.find(qn('wp:positionH'))
            vertical = anchor.find(qn('wp:positionV'))
            if horizontal is None or vertical is None:
                continue
            if horizontal.get('relativeFrom') != 'page' or vertical.get('relativeFrom') != 'page':
                continue
            x = int(horizontal.find(qn('wp:posOffset')).text) / 12700
            y = int(vertical.find(qn('wp:posOffset')).text) / 12700
            extent = anchor.find(qn('wp:extent'))
            width = int(extent.get('cx')) / 12700
            matches = [fitz.Rect(info['bbox']) for info in source[page_index].get_image_info()
                       if abs(info['bbox'][0]-x) < 1 and abs(info['bbox'][1]-y) < 1
                       and abs(info['bbox'][2]-info['bbox'][0]-width) < 1]
            if len(matches) == 1:
                rect = matches[0]
                for node in [extent, *anchor.xpath('.//a:xfrm/a:ext')]:
                    node.set('cx', str(round(rect.width * 12700)))
                    node.set('cy', str(round(rect.height * 12700)))
        if child.tag == qn('w:sectPr') or child.xpath('./w:pPr/w:sectPr'):
            page_index += 1


def restore_page_footers(doc, source):
    """将 PDF 底部页码恢复为 Word 页脚，避免正文前增高后挤出空白页。"""
    groups, group = [], []
    for child in doc.element.body:
        group.append(child)
        if child.tag == qn('w:sectPr') or child.xpath('./w:pPr/w:sectPr'):
            groups.append(group)
            group = []
    if len(groups) != len(source) or len(doc.sections) != len(source):
        return
    started = False
    for section, group, page in zip(doc.sections, groups, source):
        section.footer.is_linked_to_previous = False
        candidates = [b for b in page.get_text('blocks')
                      if b[1] > page.rect.height * .9 and b[4].strip().isdigit()]
        if len(candidates) != 1:
            continue
        block = candidates[0]
        number = block[4].strip()
        candidates = [p for p in group if p.tag == qn('w:p') and p.xpath('.//w:t')]
        if not candidates or ''.join(candidates[-1].xpath('.//w:t/text()')).strip() != number:
            continue
        original = candidates[-1]
        # 定位间隔转成段前间隔会累积字体度量误差；正文保留 75% 的段间空白。
        for p in group:
            if p.tag != qn('w:p') or not p.xpath('./w:r/w:t'):
                continue
            for spacing in p.xpath('./w:pPr/w:spacing'):
                for key in ('before', 'after'):
                    attr = qn('w:' + key)
                    if spacing.get(attr):
                        spacing.set(attr, str(round(int(spacing.get(attr)) * .75)))
        footer_p = section.footer.paragraphs[0]._p
        for child in list(footer_p):
            footer_p.remove(child)
        for child in original:
            footer_p.append(deepcopy(child))
        props = footer_p.get_or_add_pPr()
        spacing = props.find(qn('w:spacing'))
        if spacing is not None:
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
        section.footer_distance = Pt(page.rect.height - block[3])
        if not started:
            num = OxmlElement('w:pgNumType')
            num.set(qn('w:start'), number)
            section._sectPr.append(num)
            started = True
        # 缓存页码保留在 w:t 中；Word 排版时计算 PAGE 字段。
        first = footer_p.xpath('./w:r[w:t]')[0]
        begin = OxmlElement('w:r')
        field = OxmlElement('w:fldChar'); field.set(qn('w:fldCharType'), 'begin'); begin.append(field)
        instruction = OxmlElement('w:r')
        code = OxmlElement('w:instrText'); code.text = ' PAGE '; instruction.append(code)
        separate = OxmlElement('w:r')
        field = OxmlElement('w:fldChar'); field.set(qn('w:fldCharType'), 'separate'); separate.append(field)
        for item in (begin, instruction, separate):
            first.addprevious(item)
        end = OxmlElement('w:r')
        field = OxmlElement('w:fldChar'); field.set(qn('w:fldCharType'), 'end'); end.append(field)
        footer_p.append(end)
        original.getparent().remove(original)
