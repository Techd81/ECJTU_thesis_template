$pdf_mode = 5;
$xelatex = 'xelatex -interaction=nonstopmode -halt-on-error -file-line-error -synctex=1 %O %S';
$xdvipdfmx = 'python -X utf8 scripts/finalize_pdf.py %O -o %D %S';
$out_dir = 'build';
$max_repeat = 6;
$bibtex_use = 2;
