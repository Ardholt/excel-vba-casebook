#!/usr/bin/python3
import argparse,json,pathlib,re,subprocess,sys,tempfile,time
from PIL import Image,ImageChops
from lo_session import Office,prop,url

def main():
    start=time.monotonic()
    render=pathlib.Path(sys.argv[0]).name=='xlrender'
    p=argparse.ArgumentParser(description='Calc recalculation and per-sheet PDF/PNG rendering (macros disabled).')
    p.add_argument('input');p.add_argument('output')
    if render:p.add_argument('--sheet')
    args=p.parse_args()
    with Office() as office:
        doc=office.load(args.input);doc.calculateAll()
        if not render:
            out=pathlib.Path(args.output).resolve()
            if out.suffix.lower()!='.xlsx':p.error('recalculation output must be .xlsx (macros are not preserved)')
            out.parent.mkdir(parents=True,exist_ok=True)
            doc.storeAsURL(url(out),(prop('FilterName','Calc MS Excel 2007 XML'),prop('Overwrite',True)))
            print(f'xlrecalc: {out} ({time.monotonic()-start:.3f}s)')
        else:
            out=pathlib.Path(args.output).resolve();out.mkdir(parents=True,exist_ok=True)
            names=list(doc.Sheets.ElementNames)
            if args.sheet and args.sheet not in names:p.error(f'unknown sheet: {args.sheet}')
            results=[]
            for idx,name in enumerate(names,1):
                if args.sheet and name!=args.sheet:continue
                sheet=doc.Sheets.getByName(name)
                # Export only this sheet, preserving the workbook's print range if it has one.
                sheet.IsVisible=True
                for other in names:
                    if other!=name:doc.Sheets.getByName(other).IsVisible=False
                controller=doc.getCurrentController();controller.setActiveSheet(sheet)
                style=doc.StyleFamilies.getByName('PageStyles').getByName(sheet.PageStyle)
                style.ScaleToPages=0;style.ScaleToPagesX=1;style.ScaleToPagesY=0
                stem=f'{idx:02d}-'+(re.sub(r'[^\w.-]+','-',name).strip('-') or 'sheet')
                pdf=out/(stem+'.pdf')
                doc.storeToURL(url(pdf),(prop('FilterName','calc_pdf_Export'),prop('Overwrite',True),
                    prop('FilterData',(prop('SinglePageSheets',False),))))
                with tempfile.TemporaryDirectory(prefix='xl-png-') as temp:
                    prefix=str(pathlib.Path(temp)/'page')
                    subprocess.run(['pdftoppm','-png','-scale-to','2400',str(pdf),prefix],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE)
                    pages=[]
                    for file in sorted(pathlib.Path(temp).glob('page-*.png')):
                        im=Image.open(file).convert('RGB')
                        box=ImageChops.difference(im,Image.new('RGB',im.size,'white')).getbbox()
                        if box:im=im.crop((max(0,box[0]-12),max(0,box[1]-12),min(im.width,box[2]+12),min(im.height,box[3]+12)))
                        pages.append(im.copy())
                    if not pages:raise RuntimeError(f'No PDF pages for {name}')
                    canvas=Image.new('RGB',(max(i.width for i in pages),sum(i.height for i in pages)+16*(len(pages)-1)),'white')
                    y=0
                    for im in pages:canvas.paste(im,(0,y));y+=im.height+16
                    # 2x 390px phone viewport, retain source aspect ratio.
                    canvas=canvas.resize((780,max(1,round(canvas.height*780/canvas.width))),Image.Resampling.LANCZOS)
                    canvas.save(out/(stem+'.png'),optimize=True)
                    results.append({'sheet':name,'png':stem+'.png','pdf':stem+'.pdf','size':list(canvas.size)})
            (out/'render-manifest.json').write_text(json.dumps(results,indent=2)+'\n')
            print(f'xlrender: {len(results)} sheet(s) ({time.monotonic()-start:.3f}s)')
        doc.close(True)
if __name__=='__main__':main()
