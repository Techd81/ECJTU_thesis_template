"""核验编译日志、TeX 实际尺寸与 PDF 排版；默认核验随附模板的完整结构。"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
import math
from pathlib import Path
import re
import sys
import unicodedata

import fitz
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
MM = 72 / 25.4
PT = 72 / 72.27


def compact(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', text))


def rows(page):
    """按实际基线合并被 PDF 拆开的章号、标题及页码。"""
    result = []
    for block in page.get_text('dict')['blocks']:
        for line in block.get('lines', []):
            for span in line['spans']:
                if not span['text'].strip():
                    continue
                baseline = span['origin'][1]
                row = next((r for r in result if abs(r['y'] - baseline) < .5), None)
                if row is None:
                    row = {'y': baseline, 'spans': []}
                    result.append(row)
                row['spans'].append(span)
    for row in result:
        row['spans'].sort(key=lambda s: s['bbox'][0])
        row['text'] = ''.join(s['text'] for s in row['spans'])
    return sorted(result, key=lambda r: r['y'])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pdf', default='build/main.pdf')
    parser.add_argument('--report', default='build/format-report.json')
    parser.add_argument('--document', action='store_true',
                        help='核验改写后的论文，允许章节名称与模板不同')
    args = parser.parse_args()
    pdf_path = (ROOT / args.pdf).resolve()
    report_path = ROOT / args.report
    checks = []

    def check(name, ok, detail=''):
        checks.append({'name': name, 'passed': bool(ok), 'detail': detail})

    required = [pdf_path, pdf_path.with_suffix('.log'),
                pdf_path.with_suffix('.toc'), pdf_path.with_suffix('.layout')]
    for path in required:
        check('构建产物 ' + path.name, path.is_file(), str(path))
    if not all(c['passed'] for c in checks):
        return finish(checks, report_path, pdf_path)

    log = pdf_path.with_suffix('.log').read_text(encoding='utf-8', errors='replace')
    diagnostics = [line for line in log.splitlines() if re.search(
        r'Warning:|^!|Overfull \\[hv]box|Missing character:|'
        r'undefined references|undefined citations|LaTeX Error|Package .+ Error', line)]
    underfull = [line for line in log.splitlines() if re.search(r'Underfull \\[hv]box', line)]
    check('编译错误、警告、缺字与溢出盒', not diagnostics, '\n'.join(diagnostics))
    class_options = re.search(r'\\documentclass(?:\[([^]]*)\])?\{ecjtu-thesis\}',
                              (ROOT/'main.tex').read_text(encoding='utf-8'))
    is_final = class_options is not None and 'final' in (class_options.group(1) or '').split(',')
    blg_path = pdf_path.with_suffix('.blg')
    blg = blg_path.read_text(encoding='utf-8', errors='replace') if blg_path.exists() else ''
    check('BibTeX 著录完成或模板示例适用', not is_final or
          (bool(blg) and not re.search(r'Warning--|error message|I couldn.t open', blg)))

    # 只检查本次编译实际读入的项目源文件，不让陈旧 PDF 冒充检查通过。
    fls = pdf_path.with_suffix('.fls')
    inputs = []
    if fls.exists():
        for line in fls.read_text(encoding='utf-8', errors='replace').splitlines():
            if line.startswith('INPUT '):
                path = (ROOT / line[6:]).resolve()
                if path.is_relative_to(ROOT) and path.suffix in {'.tex', '.cls', '.png', '.jpg', '.pdf'}:
                    if path != pdf_path and path.is_file():
                        inputs.append(path)
    inputs += list(ROOT.glob('*.bib'))
    calibration_path = pdf_path.with_suffix('.calibration.json')
    calibration = json.loads(calibration_path.read_text(encoding='utf-8')) if calibration_path.exists() else {}
    if calibration:
        inputs += [ROOT/'assets/guide-layout.json', ROOT/'scripts/pdf_microtype.py',
                   ROOT/'scripts/finalize_pdf.py']
        check('PDF 校准记录对应当前产物', calibration.get('pdf_sha256') == hashlib.sha256(pdf_path.read_bytes()).hexdigest()
              and calibration.get('content_preserved') is True)
    stale = sorted({str(p.relative_to(ROOT)) for p in inputs
                    if p.stat().st_mtime > pdf_path.stat().st_mtime + 1})
    check('PDF 对应当前源文件', bool(inputs) and not stale, ', '.join(stale))

    metrics = dict(re.findall(r'^(\w+)=([\d.]+)\s*pt$',
                             pdf_path.with_suffix('.layout').read_text(), re.M))
    # 指定 Word PDF 的实际尺寸含有 Word 的单位取整；与名义 mm 值分开核验。
    # 此处是固定的参考测量值，不能从待检 class 或 PDF 反推期望值。
    expected = {'paperwidth': 595.32, 'paperheight': 841.92,
                'left': 70.944, 'evenleft': 70.944, 'right': 56.676,
                'top': 70.9, 'bottom': 70.9, 'header': 42.55,
                'footer': 49.6402, 'bodyfont': 12, 'baseline': 19.5}
    for key, target in expected.items():
        actual = float(metrics.get(key, 'nan')) * PT
        check('实际尺寸 ' + key, abs(actual-target) < .03,
              f'{actual:.4f} bp / 要求 {target:.4f} bp')

    doc = fitz.open(pdf_path)
    page_rows = [rows(p) for p in doc]
    bookmarks = doc.get_toc()
    check('三级 PDF 书签', {1, 2, 3}.issubset({b[0] for b in bookmarks}))
    chapter_one = next((b for b in bookmarks if compact(b[1]).startswith('第1章')), None)
    check('正文第一章存在', chapter_one is not None)
    if chapter_one is None:
        return finish(checks, report_path, pdf_path)
    body_start = chapter_one[2] - 1
    def first_page(predicate):
        return next((i for i, pr in enumerate(page_rows[:body_start]) if predicate(pr)), None)

    declaration_index = first_page(lambda pr: any('诚信声明' in compact(r['text']) for r in pr))
    cn_index = first_page(lambda pr: any(compact(r['text']) == '摘要' for r in pr))
    en_index = first_page(lambda pr: any(compact(r['text']).upper() == 'ABSTRACT' for r in pr))
    toc_index = first_page(lambda pr: any(compact(r['text']) == '目录' for r in pr))
    check('封面、声明、两种摘要、目录顺序',
          all(i is not None for i in (declaration_index, cn_index, en_index, toc_index))
          and 0 < declaration_index < cn_index < en_index < toc_index < body_start)
    if any(i is None for i in (declaration_index, cn_index, en_index, toc_index)):
        return finish(checks, report_path, pdf_path)

    allowed_colors = {0} if is_final else {0, 0x0000FF, 0xFF0000, 0x4F81BD}

    violations, colors, fonts, sizes = [], [], Counter(), Counter()
    for i, (page, pr) in enumerate(zip(doc, page_rows)):
        check(f'第 {i+1} 页参考 A4 尺寸', abs(page.rect.width-595.32) < .03
              and abs(page.rect.height-841.92) < .03)
        check(f'第 {i+1} 页有效内容', bool(pr) or i in {1, 3})
        for row in pr:
            for s in row['spans']:
                fonts[s['font']] += len(s['text'])
                sizes[round(s['size'], 2)] += len(s['text'])
                if s['color'] not in allowed_colors:
                    colors.append(f'{i+1}: {s["text"][:30]}')
        if i >= 2:
            # xeCJK 压缩标点的字面框含空白边，字面框可越界而墨迹仍在版心。
            # 直接检查 144 dpi 渲染的墨迹；0.8 bp 容差用于像素取整与笔画边缘。
            pix = page.get_pixmap(matrix=fitz.Matrix(2,2), colorspace=fitz.csGRAY, alpha=False)
            im = Image.frombytes('L',(pix.width,pix.height),pix.samples)
            left = math.floor((25*MM-.8)*2)
            right_margin = 19 if i == en_index else 20
            right = math.ceil(((210-right_margin)*MM+.8)*2)
            top_margin = (15 if toc_index < i < body_start else
                          17.2 if i in (cn_index, toc_index) else
                          17.6 if i == en_index else 25)
            top = math.floor((top_margin*MM-.8)*2)
            bottom = math.ceil((272*MM+.8)*2)
            strips = {'左边距':(0,120,left,1550), '右边距':(right,120,pix.width,1550),
                      '上边距':(left,120 if i>=body_start else 0,right,top),
                      '下边距':(left,bottom,right,1550 if i>=body_start else pix.height)}
            for label, box in strips.items():
                histogram = im.crop(box).histogram()
                # RGB 透明度组中的斜体末笔可能只留下一个低覆盖率抗锯齿像素。
                # 参考 PDF 第 15 页在 (1079, 862) 也有灰度 178 的同一像素。
                # 保留 0.8 bp 边界；单个浅色边缘像素不构成可见文字越界。
                if sum(histogram[:180]) > 1 or sum(histogram[:128]):
                    violations.append(f'{i+1} 页{label}存在墨迹')
        for f in page.get_fonts():
            check(f'第 {i+1} 页字体嵌入 {f[3]}', bool(doc.extract_font(f[0])[3]))
    check('摘要、目录、正文文字位于版心', not violations, '\n'.join(violations))
    check('文字颜色符合文档模式', not colors, '\n'.join(colors))
    check('宋体、黑体、楷体及 Times New Roman 已使用',
          all(any(f.startswith(n) for f in fonts) for n in
              ['SimSun', 'SimHei', 'KaiTi', 'TimesNewRomanPS']))

    center = (25*MM+190*MM)/2
    for i in range(body_start, len(doc)):
        page_no = i-body_start+1
        header = [r for r in page_rows[i] if r['y'] < 60]
        footer = [r for r in page_rows[i] if r['y'] > 775]
        expected_header = ('华东交通大学毕业设计' if page_no % 2 else
                           '学生姓名：毕业论文题目（请将这里修改为自己的姓名和题目，并将括号内容删除。）'
                           if not is_final else doc.metadata['author']+'：'+doc.metadata['title'])
        check(f'正文 {page_no} 页眉内容及奇偶规则',
              len(header) == 1 and compact(header[0]['text']) == compact(expected_header))
        check(f'正文 {page_no} 连续阿拉伯页码',
              len(footer) == 1 and footer[0]['text'].strip() == str(page_no))
        for kind, group, family in [('页眉', header, 'SimSun'), ('页脚', footer, 'TimesNewRoman')]:
            if len(group) == 1:
                spans = group[0]['spans']
                check(f'正文 {page_no} {kind}字号与居中',
                      all(abs(s['size']-9)<.03 for s in spans)
                      and abs((spans[0]['bbox'][0]+spans[-1]['bbox'][2])/2-center)<1)
                check(f'正文 {page_no} {kind}字体',
                      all(s['font'].startswith(family if re.search(r'[\u4e00-\u9fff]', s['text'])
                                                else 'TimesNewRoman') for s in spans))
    check('前置部分无页码页眉', all(not any((r['y'] < 60 and
                                                   ('华东交通大学毕业设计' in r['text'] or
                                                    r['text'].startswith('学生姓名：'))) or
                                                  (r['y'] > 775 and r['text'].strip().isdigit())
                                                  for r in page_rows[i])
                                    for i in range(2, body_start)))

    # 书签、目录显示页码和实际目标页互相校对。
    toc_rows = [r for pr in page_rows[toc_index:body_start] for r in pr]
    toc_links = [l for p in list(doc)[toc_index:body_start] for l in p.get_links() if l.get('page',-1)>=0]
    for level, title, target in bookmarks:
        heading = compact(title)
        matches = [r for r in toc_rows if compact(r['text']).startswith(heading)]
        check('目录页码 ' + title, any(compact(r['text']).endswith(str(target-body_start)) for r in matches))
        check('目录跳转 ' + title, any(l['page'] == target-1 and
              any(l['from'].y0-.5 < r['y'] < l['from'].y1+.5 for r in matches) for l in toc_links))
        found = [r for r in page_rows[target-1] if heading in compact(r['text']) and 60<r['y']<775]
        check('正文标题 ' + title, bool(found))
        target_size = {1:15.96, 2:14.04, 3:12}[level]
        check('标题字号 ' + title, bool(found) and
              all(abs(s['size']-target_size)<.03 for s in found[0]['spans']))
        if found:
            check('标题中文黑体 ' + title, all(s['font'].startswith('SimHei') for s in found[0]['spans']
                                           if re.search(r'[\u4e00-\u9fff]', s['text'])))
        if matches:
            check('目录字号 ' + title, all(abs(s['size']-12)<.03 for s in matches[0]['spans']))
            x0 = min(s['bbox'][0] for s in matches[0]['spans'])
            check('目录逐级缩进 ' + title, abs(x0-{1:70.944, 2:81.504, 3:91.944}[level]) < .8)

    if not args.document:
        source_map = json.loads((ROOT/'docs/source-map.json').read_text(encoding='utf-8'))
        expected_titles = [compact(s['title']) for s in source_map]
        check('原 Word 的全部章节结构', len(bookmarks) == len(source_map)+4 and
              all(any(compact(b[1]).endswith(t) for b in bookmarks) for t in expected_titles),
              f'{len(source_map)} 个章节标题 + 致谢、参考文献、两个附录')
        check('Word 后置部分顺序', [compact(b[1]) for b in bookmarks if b[0] == 1][-4:] ==
              ['致谢', '参考文献', '附录A软件使用说明书', '附录B主要源代码'])

    for idx, label, heading, family in [(cn_index,'中文','摘要','SimHei'),
                                         (en_index,'英文','ABSTRACT','TimesNewRoman')]:
        pr = page_rows[idx]
        abstract_heading = [r for r in pr if compact(r['text']).upper() == heading.upper()]
        check(label+'摘要三号标题', len(abstract_heading)==1 and
              all(abs(s['size']-15.96)<.03 and s['font'].startswith(family) for s in abstract_heading[0]['spans']))
        keywords = [r for r in pr if compact(r['text']).startswith('关键词' if idx==cn_index else 'Keywords:')]
        check(label+'关键词独立成行、顶格及小四字号', len(keywords)==1 and
              abs(keywords[0]['spans'][0]['bbox'][0]-25*MM)<.8 and
              all(abs(s['size']-12)<.03 for s in keywords[0]['spans']))
        if keywords:
            check(label+'关键词分号及数量', 2 <= compact(keywords[0]['text']).count(';') <= 4
                  and not compact(keywords[0]['text']).endswith((';', '.', '。', ',')))
    keyword_y = next((r['y'] for r in page_rows[cn_index] if '关键词' in r['text']), None)
    cnbody = [s for r in page_rows[cn_index] for s in r['spans']
              if keyword_y is not None and 150<r['y']<keyword_y and re.search(r'[\u4e00-\u9fff]',s['text'])]
    check('中文正文小四宋体', bool(cnbody) and all(s['font']=='SimSun' and abs(s['size']-12)<.03 for s in cnbody))
    # 长段落的连续基线直接从 PDF 测量，避免仅凭配置值宣称行距正确。
    baselines = [r['y'] for r in page_rows[cn_index] if keyword_y is not None and 150<r['y']<keyword_y]
    gaps = [b-a for a,b in zip(baselines, baselines[1:])]
    if calibration.get('status') == 'applied':
        # 指定 Word PDF 的中文摘要实测基线；Word 按打印网格取整，
        # 相邻行交替为 19.44/19.56 bp，跨段另有 0.02～0.03 bp 取整。
        # 使用独立参考值及 0.03 bp 容差，不从待检 PDF/校准参数反推目标。
        word_baselines = [156.62, 176.18, 195.62, 215.18, 234.65, 254.21,
                          273.65, 293.21, 312.65, 332.21, 351.65, 371.21,
                          390.65, 410.23, 429.67, 449.23, 468.67]
        line_spacing_ok = len(baselines) == len(word_baselines) and all(
            abs(actual-target) < .03 for actual, target in zip(baselines, word_baselines))
        spacing_detail = '按参考 Word 的 17 条基线核验，容差 0.03 bp'
    else:
        line_spacing_ok = len(gaps)>5 and all(abs(g-19.5)<.05 for g in gaps)
        spacing_detail = '正常 LaTeX 正文基线间距 19.5 bp，容差 0.05 bp'
    check('PDF 正文行距 19.5 bp', line_spacing_ok, spacing_detail)
    toc_heading = [r for r in page_rows[toc_index] if compact(r['text'])=='目录']
    check('目录标题三号黑体', len(toc_heading)==1 and
          all(s['font'].startswith('SimHei') and abs(s['size']-15.96)<.03 for s in toc_heading[0]['spans']))
    ref_page = next((b[2]-1 for b in bookmarks if compact(b[1])=='参考文献'), None)
    if ref_page is not None:
        ref_spans = [s for r in page_rows[ref_page] if 100<r['y']<775 for s in r['spans']]
        check('参考文献五号字体', bool(ref_spans) and all(abs(s['size']-10.56)<.03 for s in ref_spans))

    if not args.document:
        figure_page = next((i for i,p in enumerate(doc) if '工资制作流程图' in p.get_text() and i>=body_start), None)
        check('图表示例存在', figure_page is not None)
        if figure_page is not None:
            pr = page_rows[figure_page]
            fig = next((r for r in pr if compact(r['text'])=='图3-1工资制作流程图'), None)
            tab = next((r for r in pr if compact(r['text'])=='表3-1工资管理问题点'), None)
            images = doc[figure_page].get_image_info()
            flow = next((image for image in images if image['width']==649 and image['height']==534), None)
            head = next((r for r in pr if '涉及点' in r['text']), None)
            check('图题在下、表题在上及章内编号', all([fig,tab,flow,head])
                  and flow['bbox'][3]<fig['y']<tab['y']<head['y'])
            for label, row in [('图题',fig),('表题',tab)]:
                check(label+'五号黑体居中', row is not None and
                      all(abs(s['size']-10.56)<.03 for s in row['spans']) and
                      all(s['font'].startswith('SimHei') for s in row['spans'] if re.search(r'[\u4e00-\u9fff]',s['text'])) and
                      abs((row['spans'][0]['bbox'][0]+row['spans'][-1]['bbox'][2])/2-center)<1)
            check('工资流程图保留原图尺寸及版位', flow is not None
                  and abs(flow['bbox'][0]-120.25)<1 and abs(flow['bbox'][1]-75.1)<1
                  and abs(flow['bbox'][2]-489.25)<1 and abs(flow['bbox'][3]-378.7)<1)

    return finish(checks, report_path, pdf_path, {'pages':len(doc), 'body_pages':len(doc)-body_start,
                  'bookmarks':len(bookmarks), 'fonts':dict(fonts), 'font_sizes_bp':dict(sizes),
                  'underfull_boxes':underfull})


def finish(checks, report_path, pdf_path, details=None):
    failures = [c for c in checks if not c['passed']]
    report = {'generated_at':datetime.now(timezone.utc).isoformat(),
              'pdf':str(pdf_path),
              'pdf_sha256':hashlib.sha256(pdf_path.read_bytes()).hexdigest() if pdf_path.exists() else None,
              'passed':not failures, 'check_count':len(checks), 'failure_count':len(failures),
              'details':details or {}, 'checks':checks,
              'scope':'模板排版的本地自动核验；逐页视觉检查另行进行。论文内容质量及校方外部检测不在本脚本范围。'}
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(f'格式检查：{len(checks)-len(failures)}/{len(checks)} 通过，{len(failures)} 项失败。')
    for c in failures:
        print('FAIL '+c['name']+': '+c['detail'])
    print('报告：'+str(report_path))
    return int(bool(failures))


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
