"""检查 Word 的可编辑文字、内容完整性及其对应的 PDF。"""
from __future__ import annotations
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import posixpath
import sys
import unicodedata
from zipfile import ZipFile

from lxml import etree
import pymupdf as fitz
from pdf_word_layout import page_text, figure_regions

ROOT = Path(__file__).resolve().parents[1]
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def characters(text):
    return Counter(c for c in unicodedata.normalize('NFKC', text) if not c.isspace())


def verify(path, mapping=None, write_report=True):
    path = Path(path).resolve()
    if mapping is None:
        mapping = json.loads(path.with_suffix('.map.json').read_text(encoding='utf-8'))
    if mapping.get('schema')==3:
        from check_thesis_word import verify_source_word
        return verify_source_word(path,mapping,write_report)
    checks = []
    def check(name, value):
        checks.append({'name': name, 'passed': bool(value)})
    check('可编辑导出格式', mapping.get('schema') == 2 and mapping.get('mode') == 'editable')
    if not checks[-1]['passed']:
        print('旧版固定页面稿，请使用新版 export_word.py 重新导出。')
        return 1
    check('Word 文件哈希', hashlib.sha256(path.read_bytes()).hexdigest() == mapping['docx_sha256'])
    source = Path(mapping['pdf'])
    check('源 PDF 哈希', hashlib.sha256(source.read_bytes()).hexdigest() == mapping['pdf_sha256'])
    with ZipFile(path) as archive, fitz.open(source) as pdf:
        document = etree.fromstring(archive.read('word/document.xml'))
        settings = etree.fromstring(archive.read('word/settings.xml'))
        styles = etree.fromstring(archive.read('word/styles.xml'))
        parts = [document] + [etree.fromstring(archive.read(n)) for n in archive.namelist()
                            if n.startswith(('word/header', 'word/footer')) and n.endswith('.xml')]
        check('正文允许编辑', not settings.xpath('//w:documentProtection[not(@w:enforcement="0" or @w:enforcement="false")]', namespaces=NS))
        check('正文为流式段落', not document.xpath('//w:framePr | //w:txbxContent', namespaces=NS))
        check('表格行高可以扩展', not document.xpath('//w:trHeight[@w:hRule="exact"]', namespaces=NS))
        check('默认文字字号正常', styles.xpath('//w:style[@w:styleId="Normal"]/w:rPr/w:sz/@w:val', namespaces=NS) == ['24'])
        applied_styles = styles.xpath('//w:docDefaults | //w:style[@w:styleId="Normal"]', namespaces=NS)
        for root in [*parts, *applied_styles]:
            check('文字可见', not root.xpath('.//w14:textFill | .//w:vanish[not(@w:val="0" or @w:val="false")] | .//w:color[@w:val="FFFFFF" or @w:val="ffffff"]', namespaces=NS))
        check('源 PDF 页数记录', len(pdf) == len(mapping['pages']))
        expected = ''
        for page, record in zip(pdf, mapping['pages']):
            detected = figure_regions(page)
            check(f'第 {page.number+1} 页图形范围', [list(r) for r in detected] == [list(r) for r in record['figures']])
            expected += page_text(page, detected)
        actual = ''.join(''.join(root.xpath('//w:t/text()', namespaces=NS)) for root in parts)
        check('全部正文字符无缺失重复', characters(actual) == characters(expected))
        paragraphs = document.xpath('//w:body//w:p[w:r/w:t]', namespaces=NS)
        check('段落数量', len(paragraphs) == len(mapping['paragraphs']))
        for p, record in zip(paragraphs, mapping['paragraphs']):
            check(record['bookmark']+' 连续原文', ''.join(p.xpath('.//w:t/text()', namespaces=NS))
                  == record['text'].replace('\n', '').replace('\t', ''))
            check(record['bookmark']+' 定位书签', p.xpath('./w:bookmarkStart/@w:name', namespaces=NS) == [record['bookmark']])
        media = {n for n in archive.namelist() if n.startswith('word/media/')}
        check('图像清单', media == set(mapping['media_sha256']))
        for name, digest in mapping['media_sha256'].items():
            check(name+' 图像完整', name in media and hashlib.sha256(archive.read(name)).hexdigest() == digest)
        rels = etree.fromstring(archive.read('word/_rels/document.xml.rels'))
        targets = {r.get('Id'): posixpath.normpath('word/'+r.get('Target', '')) for r in rels}
        for rid in document.xpath('//a:blip/@r:embed', namespaces=NS):
            check('嵌入图像关系', targets.get(rid) in media)
    result = {'passed': all(c['passed'] for c in checks), 'checks': checks,
              'scope': '可编辑文档结构和内容检查；Word 实际输入、保存及渲染由 verify_word.py 验证。'}
    if write_report:
        path.with_suffix('.check.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    failures = [c['name'] for c in checks if not c['passed']]
    print(f'Word 结构检查：{len(checks)-len(failures)}/{len(checks)} 通过', flush=True)
    for name in failures:
        print('失败：'+name)
    return 0 if result['passed'] else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx', nargs='?', type=Path, default=ROOT/'build/main-review.docx')
    args = parser.parse_args()
    try:
        return verify(args.docx)
    except (OSError, ValueError, KeyError) as error:
        print(f'Word 检查失败：{error}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
