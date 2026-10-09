"""从论文源码生成学校样式 Word；也可显式转换独立 PDF。"""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
from zipfile import ZipFile, BadZipFile

import pymupdf as fitz
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from pdf2docx import Converter
from pdf_word_layout import (editable_pdf_parser, figure_regions, normalize_editing,
                             restore_floating_image_sizes, restore_page_footers)

ROOT = Path(__file__).resolve().parents[1]


def protect_existing(output):
    if not output.exists():
        return
    try:
        with ZipFile(output) as existing:
            if 'word/comments.xml' in existing.namelist():
                if len(etree.fromstring(existing.read('word/comments.xml'))):
                    raise ValueError('输出文件已有批注，请用 --output 指定新文件名。')
            root = etree.fromstring(existing.read('word/document.xml'))
            if next(root.iter(qn('w:ins')), None) is not None or next(root.iter(qn('w:del')), None) is not None:
                raise ValueError('输出文件已有修订，请用 --output 指定新文件名。')
        mapping = json.loads(output.with_suffix('.map.json').read_text(encoding='utf-8'))
        if mapping.get('docx_sha256') == hashlib.sha256(output.read_bytes()).hexdigest():
            return
    except BadZipFile as error:
        raise ValueError('已有输出不是有效的 Word 文件，请指定新文件名。') from error
    except (OSError, KeyError, json.JSONDecodeError):
        pass
    raise ValueError('输出文件已修改或缺少导出记录，请用 --output 指定新文件名，保留现有内容。')


def export(pdf_path, output, dpi=300):
    pdf_path, output = Path(pdf_path).resolve(), Path(output).resolve()
    if pdf_path == output or output.suffix.lower() != '.docx':
        raise ValueError('输出应为独立的 .docx 文件。')
    if not 150 <= dpi <= 600:
        raise ValueError('--dpi 应在 150 至 600 之间')
    protect_existing(output)
    pdf_bytes = pdf_path.read_bytes()
    report = {'schema': 2, 'mode': 'editable',
              'created_at': datetime.now(timezone.utc).isoformat(),
              'pdf': str(pdf_path), 'pdf_sha256': hashlib.sha256(pdf_bytes).hexdigest(),
              'docx': str(output), 'figure_dpi': dpi, 'pages': [], 'paragraphs': []}
    with fitz.open(stream=pdf_bytes, filetype='pdf') as source:
        if source.needs_pass or not len(source):
            raise ValueError('PDF 为空或需要密码。')
        regions = {}
        for index, page in enumerate(source):
            text = page.get_text()
            if not text.strip():
                raise ValueError(f'第 {index+1} 页没有可提取的文字，请先对扫描件进行 OCR。')
            if '\ufffd' in text:
                raise ValueError(f'第 {index+1} 页含无法解码的文字，请检查 PDF 字体编码。')
            regions[index] = figure_regions(page)
            report['pages'].append({'number': index+1, 'width': page.rect.width,
                                   'height': page.rect.height, 'figures': regions[index]})
        output.parent.mkdir(parents=True, exist_ok=True)
        # 完成转换和检查后再替换目标，解析失败不影响已有文件。
        with tempfile.TemporaryDirectory(prefix='word-export-', dir=output.parent) as work:
            temporary = Path(work)/'converted.docx'
            converter = Converter(stream=pdf_bytes)
            try:
                with editable_pdf_parser():
                    converter.convert(str(temporary), ignore_page_error=False,
                        raw_exceptions=True, word_figures=regions, clip_image_res_ratio=dpi/72)
            finally:
                converter.close()
            doc = Document(temporary)
            normalize_editing(doc)
            restore_floating_image_sizes(doc, source)
            restore_page_footers(doc, source)
            doc.core_properties.title = source.metadata.get('title') or pdf_path.stem
            doc.core_properties.author = source.metadata.get('author') or ''
            doc.core_properties.subject = '论文 Word 稿'
            doc.core_properties.comments = ''
            doc.core_properties.created = doc.core_properties.modified = datetime.now(timezone.utc)
            for p in doc.element.xpath('//w:body//w:p[w:r/w:t]'):
                index = len(report['paragraphs'])
                name = f'paragraph_{index}'
                start = OxmlElement('w:bookmarkStart')
                start.set(qn('w:id'), str(index))
                start.set(qn('w:name'), name)
                end = OxmlElement('w:bookmarkEnd')
                end.set(qn('w:id'), str(index))
                p.insert(1 if p.find(qn('w:pPr')) is not None else 0, start)
                p.append(end)
                text = ''.join(n.text or '' if n.tag == qn('w:t') else '\t' if n.tag == qn('w:tab') else '\n'
                               for n in p.iter() if n.tag in (qn('w:t'), qn('w:br'), qn('w:tab')))
                report['paragraphs'].append({'bookmark': name, 'text': text,
                                              'in_table': bool(p.xpath('ancestor::w:tc'))})
            doc.save(temporary)
            report['docx_sha256'] = hashlib.sha256(temporary.read_bytes()).hexdigest()
            with ZipFile(temporary) as archive:
                report['media_sha256'] = {name: hashlib.sha256(archive.read(name)).hexdigest()
                    for name in archive.namelist() if name.startswith('word/media/')}
            from check_word import verify
            if verify(temporary, mapping=report, write_report=False):
                raise ValueError('Word 内容或可编辑性检查失败，未替换目标文件。')
            temporary.replace(output)
    output.with_suffix('.map.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(f'可编辑 Word 已生成：{output}\n源 PDF {len(report["pages"])} 页，{len(report["paragraphs"])} 个文字段落。', flush=True)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pdf', type=Path, help='转换指定 PDF；省略时先构建当前论文')
    parser.add_argument('--source', type=Path, help='与 --pdf 同用时，从指定 LaTeX 源码恢复段落与学校样式')
    parser.add_argument('--output', type=Path, default=ROOT/'build/main-review.docx')
    parser.add_argument('--dpi', type=int, default=300, help='复杂图形保留为图片时的分辨率，默认 300')
    args = parser.parse_args()
    if args.source is not None and args.pdf is None:
        parser.error('--source 需要与已编译的 --pdf 一同指定；直接运行脚本会构建并导出 main.tex。')
    if args.pdf is None:
        result = subprocess.call([sys.executable, str(ROOT/'scripts/build.py')], cwd=ROOT)
        if result:
            return result
    try:
        if args.source is not None or args.pdf is None:
            from export_thesis_word import export_source
            export_source(args.source or ROOT/'main.tex',args.pdf or ROOT/'build/main.pdf',args.output)
        else:
            export(args.pdf, args.output, args.dpi)
        from check_word import verify
        return verify(args.output)
    except (OSError, ValueError, ImportError, RuntimeError, subprocess.CalledProcessError, fitz.FileDataError) as error:
        print(f'Word 导出失败：{error}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
