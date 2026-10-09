"""在 tmp/ 中独立验证首次编译、final 选项和检查器拒绝错误版式的能力。"""
from pathlib import Path
import json
import re
import shutil
import subprocess
import sys
import tempfile

import fitz

ROOT = Path(__file__).resolve().parents[1]


def main():
    (ROOT/'tmp').mkdir(exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix='verify-', dir=ROOT/'tmp')).resolve()
    assert work.is_relative_to((ROOT/'tmp').resolve())
    for name in ['main.tex','metadata.tex','ecjtu-thesis.cls','references.bib','.latexmkrc']:
        shutil.copy2(ROOT/name, work/name)
    for name in ['assets','chapters','scripts','docs']:
        shutil.copytree(ROOT/name, work/name, ignore=shutil.ignore_patterns('__pycache__'))
    results = []

    def run(label, command, expect=0):
        proc = subprocess.run(command, cwd=work, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        text = proc.stdout.decode('utf-8',errors='replace')
        (work/(label+'.log')).write_text(text,encoding='utf-8')
        passed = proc.returncode == expect
        results.append({'test':label,'passed':passed,'exit_code':proc.returncode})
        print(label+': '+('PASS' if passed else 'FAIL'),flush=True)
        if not passed:
            print(text[-5000:])
            raise RuntimeError(label)

    run('clean-build', [sys.executable,'scripts/build.py','--check'])
    calibration = json.loads((work/'build/main.calibration.json').read_text(encoding='utf-8'))
    assert calibration['status'] == 'applied' and calibration['content_preserved']
    # The temporary copy has no reference PDF: its stored numeric profile must
    # reproduce the checked output using only its own LaTeX source and assets.
    with fitz.open(ROOT/'build/main.pdf') as baseline, fitz.open(work/'build/main.pdf') as fresh:
        identical = len(baseline) == len(fresh) and all(
            a.get_pixmap(dpi=144).samples == b.get_pixmap(dpi=144).samples for a,b in zip(baseline,fresh))
    assert identical
    results.append({'test':'clean-build-reproduces-checked-pages','passed':identical})
    shutil.copy2(work/'build/main.calibration.json', work/'clean-build.calibration.json')

    # Editing an example must keep the new prose and use ordinary TeX flow.
    chapter = work/'chapters/chapter07.tex'
    original_chapter = chapter.read_bytes()
    chapter.write_text(chapter.read_text(encoding='utf-8')+'\n用户新增正文用于核验正常排版。\n',encoding='utf-8')
    run('edited-manuscript-build', [sys.executable,'scripts/build.py'])
    edited_calibration = json.loads((work/'build/main.calibration.json').read_text(encoding='utf-8'))
    with fitz.open(work/'build/main.pdf') as edited:
        retained = '用户新增正文用于核验正常排版。' in ''.join(p.get_text() for p in edited)
    assert edited_calibration['status'] == 'normal-layout' and retained
    results.append({'test':'edited-manuscript-preserves-new-prose','passed':retained})
    chapter.write_bytes(original_chapter)
    main_path = work/'main.tex'
    main_path.write_text(main_path.read_text(encoding='utf-8').replace(
        r'\documentclass{ecjtu-thesis}', r'\documentclass[final]{ecjtu-thesis}'),encoding='utf-8')
    run('final-build', [sys.executable,'scripts/build.py','--check'])
    final_calibration = json.loads((work/'build/main.calibration.json').read_text(encoding='utf-8'))
    assert final_calibration['status'] == 'normal-layout'
    pdf = fitz.open(work/'build/main.pdf')
    text = ''.join(p.get_text() for p in pdf)
    passed = '在此撰写本节内容' not in text and '本系统采用B/S架构' not in text
    results.append({'test':'final-hides-guidance','passed':passed})
    assert passed
    pdf.close()

    # 改动源码而不编译，必须拒绝陈旧 PDF。
    main_path.write_text(main_path.read_text(encoding='utf-8')+'\n% freshness probe\n',encoding='utf-8')
    import os
    stamp = (work/'build/main.pdf').stat().st_mtime+3
    os.utime(main_path,(stamp,stamp))
    run('reject-stale-pdf', [sys.executable,'scripts/check_format.py'],expect=1)
    stale = json.loads((work/'build/format-report.json').read_text(encoding='utf-8'))
    assert any(c['name']=='PDF 对应当前源文件' and not c['passed'] for c in stale['checks'])
    os.utime(main_path,None)

    cls = work/'ecjtu-thesis.cls'
    class_source = cls.read_text(encoding='utf-8')
    wrong_margin_source, changes = re.subn(
        r'\bleft=\d+(?:\.\d+)?(?:mm|bp)(?=,right=)',
        'left=30mm', class_source, count=1)
    assert changes == 1, '主版心的左边距配置未找到，不能执行错误页边距验证。'
    cls.write_text(wrong_margin_source, encoding='utf-8')
    run('wrong-margin-build', [sys.executable,'scripts/build.py'])
    run('reject-wrong-margin', [sys.executable,'scripts/check_format.py'],expect=1)
    bad = json.loads((work/'build/format-report.json').read_text(encoding='utf-8'))
    assert any(c['name']=='实际尺寸 left' and not c['passed'] for c in bad['checks'])

    # 编译说明文档中的原样代码，续表补足 80 行以覆盖跨页表头。
    shutil.copy2(ROOT/'ecjtu-thesis.cls', cls)
    examples = re.findall(r'```latex\n(.*?)\n```',
                          (work/'docs/补充排版示例.md').read_text(encoding='utf-8'),re.S)
    long_table = examples[1].replace('% 在这里添加其他测试用例。',
        '\n'.join(f'UC-{i:02d} & 访问测试接口 & 返回预期结果'+r'\\' for i in range(3,81)))
    (work/'assets/Example.java').write_text('public class Example {\n    public int value() { return 1; }\n}\n')
    (work/'examples.tex').write_text(r'\documentclass{ecjtu-thesis}'+'\n'+examples[2]+'\n'+
        r'\begin{document}\thesismainmatter\chapter{排版示例}'+'\n'+examples[0]+'\n'+long_table+'\n'+
        examples[3]+'\n'+examples[4]+'\n'+r'\end{document}',encoding='utf-8')
    run('documentation-examples', ['latexmk','-xelatex','examples.tex'])
    log = (work/'build/examples.log').read_text(encoding='utf-8',errors='replace')
    assert not re.search(r'Warning:|Overfull|Underfull|Missing character:|^!',log,re.M)
    example_pdf = fitz.open(work/'build/examples.pdf')
    continuation_pages = [p for p in example_pdf if '续表' in p.get_text()]
    assert continuation_pages and all('预期结果' in p.get_text() for p in continuation_pages)
    example_pdf.close()

    report = {'passed':all(r['passed'] for r in results),'workspace':str(work),'tests':results}
    (ROOT/'build/verification-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('All verification cases passed. Reports retained in '+str(work))
    return 0


if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
