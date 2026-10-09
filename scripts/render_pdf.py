"""为逐页版面检查生成 PNG 与联系表。"""
import argparse
from pathlib import Path
import fitz
from PIL import Image, ImageDraw

parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('pdf',nargs='?',default='build/main.pdf')
parser.add_argument('--output',default='tmp/qa')
args=parser.parse_args()
out=Path(args.output)
out.mkdir(parents=True,exist_ok=True)
doc=fitz.open(args.pdf)
thumbs=[]
for i,p in enumerate(doc):
    pix=p.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
    pix.save(out/f'page-{i+1:02d}.png')
    im=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
    im.thumbnail((298,422))
    tile=Image.new('RGB',(318,455),'#dddddd')
    tile.paste(im,((318-im.width)//2,24))
    ImageDraw.Draw(tile).text((8,6),str(i+1),fill='black')
    thumbs.append(tile)
for start in range(0,len(thumbs),9):
    sheet=Image.new('RGB',(954,1365),'white')
    for j,tile in enumerate(thumbs[start:start+9]):
        sheet.paste(tile,((j%3)*318,(j//3)*455))
    sheet.save(out/f'contact-{start//9+1}.jpg')
print(f'{len(doc)} pages rendered to {out}')
