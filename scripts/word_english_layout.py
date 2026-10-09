"""Add editable discretionary hyphens at LaTeX's English line breaks."""
import pymupdf as fitz
from docx.oxml import OxmlElement
from docx.shared import Pt

from word_line_layout import normalized, source_body_characters


def restore_english_hyphens(doc, pdf_path, right_indent_pt=0):
    with fitz.open(pdf_path) as pdf:
        raw = source_body_characters(pdf, 2)
    chars = []
    hyphens = set()
    for index, char in enumerate(raw):
        if (char[0] == '-' and index + 1 < len(raw)
                and raw[index + 1][2] > char[2] + 5
                and raw[index + 1][3] < char[3]
                and raw[index + 1][0].isalpha()):
            hyphens.add(len(chars))
            continue
        chars.append(char)
    source = ''.join(c[0] for c in chars)
    cursor = 0
    records = []
    for paragraph in doc.paragraphs:
        if paragraph.style.name != 'English Body':
            continue
        value = normalized(paragraph.text)
        start = source.find(value, cursor)
        if start < 0 or paragraph._p.xpath('.//m:oMath | .//w:fldChar | .//w:hyperlink'):
            continue
        cursor = start + len(value)
        positions = {i - start for i in hyphens if start < i < cursor}
        if not positions:
            continue
        offset = 0
        for node in list(paragraph._p.xpath('.//w:t')):
            original = node.text or ''
            chunks = []
            current = ''
            for char in original:
                size = len(normalized(char))
                if size and offset in positions:
                    chunks.append(current)
                    current = ''
                current += char
                offset += size
            chunks.append(current)
            if len(chunks) == 1:
                continue
            node.text = chunks[0]
            previous = node
            for chunk in chunks[1:]:
                hyphen = OxmlElement('w:softHyphen')
                previous.addnext(hyphen)
                replacement = OxmlElement('w:t')
                replacement.text = chunk
                replacement.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                hyphen.addnext(replacement)
                previous = replacement
        if right_indent_pt:
            paragraph.paragraph_format.right_indent = Pt(right_indent_pt)
        records.append({'prefix': paragraph.text[:35], 'hyphens': len(positions)})
    return records
