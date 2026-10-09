"""Calibrate uniform paragraph baseline offsets against native PDF glyphs."""
import statistics

import pymupdf as fitz
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from compare_word_pdf import characters
from word_line_layout import normalized


def calibrate_paragraph_baselines(doc,source_pdf,word_pdf):
    def collect(pdf):
        rows=[];starts=[]
        for page in pdf:
            starts.append(len(rows))
            rows.extend((char,page.number,x,y) for char,(x,y) in characters(page))
        return ''.join(row[0] for row in rows),rows,starts
    with fitz.open(source_pdf) as source,fitz.open(word_pdf) as word:
        first_body=min(row[2] for row in source.get_toc())-1
        a,aa,ap=collect(source);b,bb,bp=collect(word)
    left=right=0;body=False;records=[]
    for p in doc.paragraphs:
        if p.style.name.lower().startswith('toc'):continue
        if not body and p.style.name=='Heading 1':
            left,right=ap[first_body],bp[first_body];body=True
        text=normalized(p.text)
        if not text:continue
        start=a.find(text,left);actual=b.find(text,right)
        if start<0 or actual<0:continue
        left,right=start+len(text),actual+len(text)
        source_chars,word_chars=aa[start:left],bb[actual:right]
        if any(s[1]!=w[1] for s,w in zip(source_chars,word_chars)):continue
        offsets=[w[3]-s[3] for s,w in zip(source_chars,word_chars)]
        # Reflow, mixed baselines, and page changes need line-level handling.
        # A uniform offset can be corrected without changing paragraph flow.
        if max(offsets)-min(offsets)>.4:continue
        delta=statistics.median(offsets)
        if abs(delta)<.26 or abs(delta)>12:continue
        step=round(delta*2)
        for run in p._p.xpath('.//w:r'):
            pr=run.get_or_add_rPr()
            old=pr.find(qn('w:position'))
            value=int(old.get(qn('w:val'))) if old is not None else 0
            if old is not None:pr.remove(old)
            setting=OxmlElement('w:position');setting.set(qn('w:val'),str(value+step));pr.append(setting)
        records.append({'prefix':p.text[:35],'page':source_chars[0][1]+1,'raise_pt':step/2})
    return records
