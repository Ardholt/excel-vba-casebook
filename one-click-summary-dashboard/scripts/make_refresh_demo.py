"""Replay actual saved worksheet renders from the recorded button activations."""
from PIL import Image,ImageDraw,ImageFont
from pathlib import Path
import hashlib,json
run=json.loads(Path('output/vba-run.json').read_text())
assert run['result']=='PASS'
paths=['images/before.png','images/dashboard.png','images/repeat.png']
labels=['1  Press: original sales list','2  Add sales, then press','3  Press again: same total']
images=[Image.open(p).convert('RGB') for p in paths]
width=max(i.width for i in images);height=max(i.height for i in images)
frames=[]
font=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',26)
for image,label in zip(images,labels):
    frame=Image.new('RGB',(width,height+58),'white')
    frame.paste(image,((width-image.width)//2,58))
    draw=ImageDraw.Draw(frame);draw.rectangle((0,0,width,55),fill='#142F43')
    draw.text((20,12),label,font=font,fill='white')
    frames.append(frame)
frames[0].save('media/refresh.gif',save_all=True,append_images=frames[1:],duration=[2200,3000,2600],loop=0,disposal=2)
Path('evidence/refresh-demo.json').write_text(json.dumps({
    'file':'media/refresh.gif','sha256':hashlib.sha256(Path('media/refresh.gif').read_bytes()).hexdigest(),
    'method':'Saved-view replay, not a desktop screen recording; captions identify the recorded accessibility button actions.',
    'button_events':'output/vba-run.json',
    'frames':[{'image':p,'source':r['saved_output'],'saved_sha256':r['saved_sha256'],'caption':label} for p,r,label in zip(paths,run['runs'],labels)]},indent=2)+'\n')
