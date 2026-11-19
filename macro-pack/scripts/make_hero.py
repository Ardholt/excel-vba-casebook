"""Render the complete delivered letter, using its file contents verbatim."""
from pathlib import Path
import textwrap
from PIL import Image, ImageDraw, ImageFont
letter=Path('output/letters/C001.txt').read_text('ascii')
font_dir=Path('/usr/share/fonts/truetype/dejavu')
body=ImageFont.truetype(str(font_dir/'DejaVuSans.ttf'),30)
label=ImageFont.truetype(str(font_dir/'DejaVuSans-Bold.ttf'),24)
lines=[]
for line in letter.splitlines(): lines.extend(textwrap.wrap(line,width=39) or [''])
im=Image.new('RGB',(720,100+len(lines)*44),'#f8f6f1'); d=ImageDraw.Draw(im)
d.rectangle((0,0,719,60),fill='#233f40')
d.text((28,16),'FINISHED LETTER  /  C001.txt',font=label,fill='white')
for i,line in enumerate(lines): d.text((28,82+i*44),line,font=body,fill='#172f30')
Path('media').mkdir(exist_ok=True); im.save('media/hero.png')
im.save('images/letter-hero.png')
