"""Recover source PDF line boundaries inside editable Word paragraphs."""
import unicodedata
from copy import deepcopy

import pymupdf as fitz
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def normalized(text):
    return ''.join(unicodedata.normalize('NFKC',text).split())


def source_body_characters(pdf,first=None):
    if first is None:first=min(row[2] for row in pdf.get_toc())-1
    result=[]
    for page_index in range(first,len(pdf)):
        for block in pdf[page_index].get_text('rawdict')['blocks']:
            for line in block.get('lines',[]):
                for span in line['spans']:
                    for char in span['chars']:
                        x,y=char['origin']
                        if not 70<y<775:continue
                        for c in normalized(char['c']):
                            result.append((c,page_index,round(y,2),x))
    return result


def line_boundaries(chars):
    return {i for i in range(1,len(chars)) if chars[i][1]!=chars[i-1][1]
            or (chars[i][2]>chars[i-1][2]+5 and chars[i][3]<chars[i-1][3])}


def restore_line_breaks(doc,pdf_path,tracking_twips=0,word_pdf=None):
    with fitz.open(pdf_path) as pdf:
        first_body=min(row[2] for row in pdf.get_toc())-1
        # The template's first two sections are the cover and declaration;
        # both abstracts precede the native TOC and academic body.
        first=min(2,first_body)
        source=source_body_characters(pdf,first)
    actual=[]
    if word_pdf is not None:
        with fitz.open(word_pdf) as pdf:actual=source_body_characters(pdf,first)
    actual_text=''.join(c[0] for c in actual)
    actual_cursor=0
    text=''.join(c[0] for c in source)
    cursor=0
    started=False;body=False
    records=[]
    for p in doc.paragraphs:
        if p.style.name in ('Abstract Title','English Title','Heading 1'):started=True
        if p.style.name.lower().startswith('toc'):continue
        if p.style.name=='Heading 1' and not body:
            body=True
            cursor=next((i for i,c in enumerate(source) if c[1]>=first_body),len(source))
            actual_cursor=next((i for i,c in enumerate(actual) if c[1]>=first_body),len(actual))
        if not started or not p.text.strip():continue
        value=normalized(p.text)
        start=text.find(value,cursor)
        if start<0:continue
        cursor=start+len(value)
        if p._p.xpath('.//m:oMath | .//w:fldChar | .//w:drawing | .//w:br'):continue
        chars=source[start:cursor]
        # Superscript citations change the baseline within the same line.
        # A real new line also returns toward the left edge.
        boundaries=line_boundaries(chars)
        if word_pdf is not None:
            at=actual_text.find(value,actual_cursor)
            if at<0:continue
            actual_cursor=at+len(value)
            if line_boundaries(actual[at:actual_cursor])==boundaries:continue
        elif not boundaries:continue
        # Word keeps the paragraph semantic boundary and all existing runs;
        # each source line ending becomes a soft line break inside a run.
        offset=0
        for node in list(p._p.xpath('.//w:t')):
            value=node.text or ''
            parts=[];current=''
            for char in value:
                size=len(normalized(char))
                if size and offset in boundaries:
                    parts.append(current.rstrip());current=''
                current+=char;offset+=size
            parts.append(current)
            if len(parts)==1:continue
            node.text=parts[0]
            tail=node
            for part in parts[1:]:
                br=OxmlElement('w:br');tail.addnext(br)
                replacement=OxmlElement('w:t')
                replacement.text=part
                replacement.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
                br.addnext(replacement);tail=replacement
        records.append({'prefix':p.text[:35],'page':chars[0][1]+1,'lines':len(boundaries)+1})
        if tracking_twips:
            line_index=0
            for run in list(p._p.xpath('./w:r')):
                props=run.rPr
                groups=[[]]
                for child in run:
                    if child.tag==qn('w:rPr'):continue
                    groups[-1].append(deepcopy(child))
                    if child.tag==qn('w:br'):groups.append([])
                for group in groups:
                    if not group:continue
                    replacement=OxmlElement('w:r')
                    if props is not None:replacement.append(deepcopy(props))
                    if line_index<len(boundaries) or chars[-1][3]>520:
                        pr=replacement.get_or_add_rPr()
                        old=pr.find(qn('w:spacing'))
                        if old is not None:pr.remove(old)
                        setting=OxmlElement('w:spacing');setting.set(qn('w:val'),str(tracking_twips));pr.append(setting)
                    for child in group:replacement.append(child)
                    run.addprevious(replacement)
                    line_index+=sum(c.tag==qn('w:br') for c in group)
                run.getparent().remove(run)
    return records
