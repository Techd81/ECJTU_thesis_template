"""latexmk's xdvipdfmx driver, with the bundled example's spacing calibration."""
from pathlib import Path
import subprocess
import sys

from pdf_microtype import finalize

ROOT = Path(__file__).resolve().parents[1]


def main():
    arguments = sys.argv[1:]
    if '-o' not in arguments or arguments.index('-o') == len(arguments)-1:
        raise SystemExit('Usage: finalize_pdf.py [xdvipdfmx options] -o output.pdf input.xdv')
    index = arguments.index('-o')+1
    output = Path(arguments[index]).resolve()
    native = output.with_suffix('.native.pdf')
    arguments[index] = str(native)
    result = subprocess.run(['xdvipdfmx', *arguments])
    if result.returncode:
        return result.returncode
    report = finalize(native, output, ROOT)
    print('PDF spacing: '+report['status'], flush=True)
    return 0


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
