"""在 Windows Microsoft Word 中核验实际输入、删除、换行、批注和保存重开。"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ET

import pymupdf as fitz

ROOT = Path(__file__).resolve().parents[1]


def compact(text):
    return ''.join(unicodedata.normalize('NFKC', text).split())


def range_content_text(word_range):
    """原生公式按 OMML 读取；Range.Text 会另行改写数学字母、减号和换行。"""
    if not word_range.OMaths.Count:
        return word_range.Text
    root = ET.fromstring(word_range.WordOpenXML)
    word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    math_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    body = root.find(f'.//{{{word_ns}}}body')
    if body is None:
        raise ValueError('Word 公式范围没有可核验的正文 XML。')
    text_tags = {f'{{{word_ns}}}t', f'{{{math_ns}}}t'}
    return ''.join(node.text or '' for node in body.iter() if node.tag in text_tags)


def visible_text(pdf):
    """按可见字形提取，去掉 Word 仿粗体重复的填充/描边指令。"""
    result=[]
    for page in pdf:
        seen=set()
        for span in page.get_texttrace():
            if span['type']==3 or span.get('opacity',1)<=.9 or min(span['color'])>=.8:
                continue
            for c in span['chars']:
                key=(c[0],round(c[2][0],3),round(c[2][1],3))
                if c[0]>=0 and key not in seen:
                    result.append(chr(c[0]));seen.add(key)
    return compact(''.join(result))


def verify(docx_path, output):
    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError('需要 Windows、Microsoft Word 和 pywin32。') from error
    docx_path = Path(docx_path).resolve()
    output = Path(output).resolve()
    mapping = json.loads(docx_path.with_suffix('.map.json').read_text(encoding='utf-8'))
    if mapping.get('schema') not in (2,3):
        raise ValueError('请先使用新版 export_word.py 生成可编辑文档。')
    if hashlib.sha256(docx_path.read_bytes()).hexdigest() != mapping['docx_sha256']:
        raise ValueError('Word 文件已修改；请核验刚导出的原始文件。')
    source_path = Path(mapping['pdf'])
    if hashlib.sha256(source_path.read_bytes()).hexdigest() != mapping['pdf_sha256']:
        raise ValueError('源 PDF 已变更，请重新导出。')
    output.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix='editable-', dir=output))
    copy = work/'editing-test.docx'
    shutil.copy2(docx_path, copy)
    checks, markers = [], []
    def check(name, passed, detail=None):
        checks.append({'name': name, 'passed': bool(passed), 'detail': detail})
    app = win32com.client.DispatchEx('Word.Application')
    app.Visible = False
    app.DisplayAlerts = 0
    document = None
    removed = None
    try:
        version = app.Version + ' (' + app.Build + ')'
        document = app.Documents.Open(str(copy), ReadOnly=False, AddToRecentFiles=False)
        document.Repaginate()
        page_count = document.ComputeStatistics(2)
        check('初始页数与导出记录相同', page_count == len(mapping['pages']), page_count)
        check('文档允许编辑', document.ProtectionType == -1)
        check('初始批注为空', document.Comments.Count == 0)
        if mapping.get('schema')==3:
            check('Word 识别原生目录',document.TablesOfContents.Count==1)
            check('Word 识别真实表格',document.Tables.Count==mapping['table_count'])
        body_start_page = 3
        if mapping.get('schema')==3:
            first_heading = compact(mapping['headings'][0])
            first_record = next(r for r in mapping['paragraphs'] if compact(r['text'])==first_heading)
            body_start_page = document.Bookmarks(first_record['bookmark']).Range.Information(3)
        page_samples = {}
        for record in mapping['paragraphs']:
            r = document.Bookmarks(record['bookmark']).Range
            actual = range_content_text(r)
            complete = compact(actual) == compact(record['text'])
            check(record['bookmark']+' 文字完整', complete,
                  None if complete else {'expected': record['text'], 'actual': actual})
            page = r.Information(3)
            if len(record['text']) > len(page_samples.get(page, {}).get('text', '')):
                page_samples[page] = record
        document.ExportAsFixedFormat(str(work/'word.pdf'), 17, Item=0, DocStructureTags=False)
        # 逐页在已有字符中间键入，完全继承光标处的原始格式。
        for page, record in sorted(page_samples.items()):
            r = document.Bookmarks(record['bookmark']).Range
            offset = len(record['text'][:len(record['text'])//2].encode('utf-16-le'))//2
            start = min(r.Start+offset, r.End)
            document.Range(start, start).Select()
            marker = f'输入验证{page:02d}ABC'
            app.Selection.TypeText(marker)
            markers.append(marker)
            inserted = document.Range(start, start+len(marker))
            check(f'第 {page} 页输入可见格式', inserted.Font.Hidden == 0 and inserted.Font.Size >= 8)
        # 用真实替换验证旧正文消失；仅选择在全文唯一出现的片段。
        for record in sorted(mapping['paragraphs'], key=lambda r: len(r['text']), reverse=True):
            if record['in_table']:
                continue
            r = document.Bookmarks(record['bookmark']).Range
            if r.Information(3) < body_start_page or r.Paragraphs(1).Style.NameLocal == 'Source Code':
                continue
            candidate = r.Text[:14]
            if len(candidate) == 14 and '\r' not in candidate and '\x0b' not in candidate and document.Content.Text.count(candidate) == 1:
                removed = candidate
                start = r.Start
                document.Range(start, start+len(candidate.encode('utf-16-le'))//2).Select()
                marker = '正文替换验证XYZ'
                app.Selection.TypeText(marker)
                markers.append(marker)
                check('删除和替换正文', removed not in document.Content.Text)
                break
        check('覆盖实际替换场景', removed is not None)
        # Enter 后输入新段落，再添加足够长的内容验证折行。
        record = next(r for r in mapping['paragraphs'] if not r['in_table'] and len(r['text']) > 60
                      and document.Bookmarks(r['bookmark']).Range.Information(3)>=body_start_page
                      and document.Bookmarks(r['bookmark']).Range.Paragraphs(1).Style.NameLocal!='Source Code')
        r = document.Bookmarks(record['bookmark']).Range
        document.Range(r.End, r.End).Select()
        app.Selection.TypeParagraph()
        marker = '新段落换行验证END'
        app.Selection.TypeText(marker + ' 新增正文应当正常显示并自动换行。'*12)
        markers.append(marker)
        # 在实际表格单元格中插入多行文字，验证行高允许增长。
        if document.Tables.Count:
            table = max((document.Tables(i) for i in range(1, document.Tables.Count+1)),
                        key=lambda t: t.Range.Cells.Count)
            cell = table.Range.Cells(table.Range.Cells.Count)
            r = cell.Range
            old_end = document.Range(r.End-2, r.End-1)
            old_bottom = (old_end.Information(3), old_end.Information(6))
            document.Range(r.End-1, r.End-1).Select()
            marker = '表格输入验证CELL'
            tail = '表格末尾验证TAIL'
            app.Selection.TypeText(marker + ' 单元格新增内容应当完整显示。'*15 + tail)
            markers.extend([marker, tail])
            document.Repaginate()
            r = cell.Range
            new_end = document.Range(r.End-2, r.End-1)
            new_bottom = (new_end.Information(3), new_end.Information(6))
            check('表格文字自动折行', new_bottom[0] > old_bottom[0] or
                  new_bottom[1] > old_bottom[1]+8, [old_bottom, new_bottom])
        if document.OMaths.Count:
            math_count = document.OMaths.Count
            r = document.OMaths(1).Range.Duplicate
            marker = 'x+98765'
            r.Text = marker
            if r.OMaths.Count:
                edited_math = r.OMaths(1)
            else:
                edited_math = document.OMaths.Add(r).OMaths(1)
            edited_math.BuildUp()
            markers.append(marker)
            check('公式修改后保留原生数学对象',
                  document.OMaths.Count == math_count and
                  compact(marker) in compact(edited_math.Range.Text))
        comment_range = document.Bookmarks(record['bookmark']).Range
        comment_range.End = min(comment_range.Start+10, comment_range.End)
        document.Comments.Add(comment_range, 'Word 编辑与批注保存测试')
        if mapping.get('schema')==3:
            toc=document.TablesOfContents(1)
            toc.Update();document.Repaginate();toc.UpdatePageNumbers()
            entries=[line for line in toc.Range.Text.split('\r') if line.strip()]
            check('编辑后目录条目完整',len(entries)==len(mapping['headings']))
            records={compact(r['text']):r for r in mapping['paragraphs']}
            for i,(line,title) in enumerate(zip(entries,mapping['headings']),1):
                record=records[compact(title)]
                page=document.Bookmarks(record['bookmark']).Range.Information(1)
                check(f'编辑后目录第 {i} 项页码',line.rsplit('\t',1)[-1].strip()==str(page))
        document.Save()
        document.Close(0)
        document = None
        document = app.Documents.Open(str(copy), ReadOnly=True, AddToRecentFiles=False)
        for marker in markers:
            check(marker+' 保存后存在', marker in document.Content.Text)
        check('替换保存后保留', removed is not None and removed not in document.Content.Text)
        check('批注保存后存在', document.Comments.Count == 1 and
              document.Comments(1).Range.Text == 'Word 编辑与批注保存测试')
        document.ExportAsFixedFormat(str(work/'edited.pdf'), 17, Item=0, DocStructureTags=False)
    finally:
        if document is not None:
            document.Close(0)
        app.Quit()
    # PDF 文字和绘制指令联合核验：透明文字虽可提取，却不会出现在可见绘制中。
    with fitz.open(work/'edited.pdf') as edited:
        visible = visible_text(edited)
        for marker in markers:
            check(marker+' 实际渲染可见', compact(marker) in visible)
        check('替换原文从页面消失', removed is not None and compact(removed) not in visible)
        for i, page in enumerate(edited):
            page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False).save(work/f'edited-page-{i+1:02d}.png')
    with fitz.open(work/'word.pdf') as rendered:
        check('原稿渲染页数', len(rendered) == len(mapping['pages']))
        for i, page in enumerate(rendered):
            page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False).save(work/f'page-{i+1:02d}.png')
    result = {'passed': all(c['passed'] for c in checks), 'word_version': version,
              'docx_sha256': mapping['docx_sha256'], 'work': str(work),
              'checks': checks, 'input_markers': markers}
    (output/'word-verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    failures = [c for c in checks if not c['passed']]
    print(f'Word 编辑核验：{len(checks)-len(failures)}/{len(checks)} 通过；报告：{output/"word-verification.json"}')
    for failure in failures:
        print(failure['name'], failure['detail'])
    return 0 if result['passed'] else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx', nargs='?', type=Path, default=ROOT/'build/main-review.docx')
    parser.add_argument('--output', type=Path, default=ROOT/'tmp/word-qa')
    args = parser.parse_args()
    try:
        return verify(args.docx, args.output)
    except (RuntimeError, OSError, ValueError) as error:
        print(f'Word 核验失败：{error}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
