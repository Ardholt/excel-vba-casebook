from pathlib import Path
from PIL import Image
for action in ['split','merge','rename','letters']:
    frames=[]
    for state in ['before','after']:
        matches=[Path('images/sheets')/('01-'+action+'-'+state+'.png')]
        assert len(matches)==1, matches
        im=Image.open(matches[0]).convert('RGB'); im.thumbnail((780,580)); frames.append(im)
    width=max(i.width for i in frames); height=max(i.height for i in frames)
    canvases=[]
    for im in frames:
        canvas=Image.new('RGB',(width,height),'white'); canvas.paste(im,(0,0)); canvases.append(canvas)
    canvases[0].save('images/'+action+'.gif',save_all=True,append_images=canvases[1:],duration=[2200,3800],loop=0,optimize=True)
