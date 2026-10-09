"""Apply the LaTeX TOC number-box geometry to Word's native TOC result."""
import re
import unicodedata

import pymupdf as fitz

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER


def format_toc_entries(doc):
    count=0
    for p in doc.paragraphs:
        match=re.fullmatch(r'toc ([123])',p.style.name,re.I)
        if not match:continue
        level=int(match.group(1))
        # Number labels occupy 48, 30, or 42 pt; the latter two also have
        # a 12/24 pt left indent. Tabs remain inside the native TOC field.
        stop={1:48,2:42,3:66}[level]
        tabs=p.paragraph_format.tab_stops
        texts=p._p.xpath('.//w:t')
        label=re.match(r'^(?:第\s*\d+\s*章|附录\s*[A-Z]|\d+(?:\.\d+){1,2})([\u2002\u3000 ]+)',
                       ''.join(t.text or '' for t in texts))
        if label:
            tabs.add_tab_stop(Pt(stop))
        tabs.add_tab_stop(Pt(165*72/25.4),WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
        if label:
            start,end=label.span(1)
            offset=0
            for text in texts:
                value=text.text or ''
                left,right=max(0,start-offset),min(len(value),end-offset)
                if left<right:
                    text.text=value[:left]
                    if offset+left==start:
                        text.addnext(OxmlElement('w:tab'))
                    if value[right:]:
                        tail=OxmlElement('w:t');tail.text=value[right:]
                        text.getparent().append(tail)
                offset+=len(value)
        leader=p._p.xpath('.//w:r[w:tab]')[-1]
        properties=leader.get_or_add_rPr()
        old=properties.find(qn('w:spacing'))
        if old is not None:properties.remove(old)
        spacing=OxmlElement('w:spacing');spacing.set(qn('w:val'),'60');properties.append(spacing)
        count+=1
    return count


def calibrate_toc_baselines(doc,source_pdf,word_pdf):
    from compare_word_pdf import characters
    def compact(text):return ''.join(unicodedata.normalize('NFKC',text).split())
    def locate(pages,text):
        matches=[]
        for index,chars in enumerate(pages):
            joined=''.join(c[0] for c in chars)
            start=joined.find(text)
            if start>=0:matches.append((index,chars[start][1][1]))
        return matches
    changes=[]
    with fitz.open(source_pdf) as source,fitz.open(word_pdf) as word:
        first_body=min(row[2] for row in source.get_toc())-1
        sources=[characters(p) for p in list(source)[:first_body]]
        actual=[characters(p) for p in list(word)[:first_body]]
        for p in doc.paragraphs:
            if p.style.name=='TOC Heading':title=compact(p.text)
            elif re.fullmatch(r'toc [123]',p.style.name,re.I):title=compact(p.text.rsplit('\t',1)[0])
            else:continue
            a,b=locate(sources,title),locate(actual,title)
            if len(a)!=1 or len(b)!=1 or a[0][0]!=b[0][0]:continue
            raise_pt=b[0][1]-a[0][1]
            if abs(raise_pt)>12:continue
            for run in p._p.xpath('.//w:r'):
                props=run.get_or_add_rPr()
                old=props.find(qn('w:position'))
                if old is not None:props.remove(old)
                position=OxmlElement('w:position')
                position.set(qn('w:val'),str(round(raise_pt*2)))
                props.append(position)
            changes.append({'title':title,'page':a[0][0]+1,'raise_pt':round(raise_pt*2)/2})
    return changes
