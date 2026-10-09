"""使用 XeLaTeX / BibTeX 构建，所有中间文件保存在 build/。"""
from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys

from pdf_microtype import finalize

ROOT = Path(__file__).resolve().parents[1]

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='构建后执行格式检查')
    parser.add_argument('--word', action='store_true', help='同时生成可编辑、可批注的 Word 文档')
    args = parser.parse_args()
    if not shutil.which('latexmk') or not shutil.which('xelatex'):
        print('需要 TeX Live 及 latexmk、xelatex 命令。', file=sys.stderr)
        return 2
    target = ROOT/'build'
    target.mkdir(exist_ok=True)
    # \include 的子目录需预先存在，适用于不同版本的 latexmk。
    (target/'chapters').mkdir(exist_ok=True)
    env = os.environ.copy()
    if os.name == 'nt':
        env['LC_ALL'] = 'C'
        env['LANG'] = 'C'
    proc = subprocess.run(['latexmk','-xelatex','main.tex'],cwd=ROOT,env=env,
                          stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
    output = proc.stdout.decode('utf-8',errors='replace')
    (target/'compile-console.log').write_text(output,encoding='utf-8')
    if proc.returncode:
        print('\n'.join(output.splitlines()[-45:]))
        return proc.returncode
    if not (target/'main.pdf').is_file():
        print('构建未产生 build/main.pdf。',file=sys.stderr)
        return 1
    # Refresh a measurement-profile/script change even when the TeX inputs did
    # not require another latexmk run.  The driver retains the native PDF.
    native = target/'main.native.pdf'
    if native.exists():
        finalize(native, target/'main.pdf', ROOT)
    else:
        result = subprocess.call([sys.executable, str(ROOT/'scripts/finalize_pdf.py'),
                                  '-E', '-o', str(target/'main.pdf'), str(target/'main.xdv')], cwd=ROOT)
        if result:
            return result
    print('PDF 已生成：'+str(target/'main.pdf'), flush=True)
    if args.check:
        result = subprocess.call([sys.executable,str(ROOT/'scripts/check_format.py')],cwd=ROOT)
        if result:
            return result
    if args.word:
        return subprocess.call([sys.executable,str(ROOT/'scripts/export_word.py'),
                                '--source',str(ROOT/'main.tex'),'--pdf',str(target/'main.pdf')],cwd=ROOT)
    return 0

if __name__ == '__main__':
    if hasattr(sys.stdout,'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
