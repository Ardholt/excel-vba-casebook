from pathlib import Path
import json, time, subprocess, tempfile, shutil, socket, hashlib, platform
import uno
from com.sun.star.beans import PropertyValue
import openpyxl, xlsxwriter, pytest, PIL
from zipfile import is_zipfile
from reference import table, REQUIRED, render_letter

def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p

def normalized(s):
    return [l.strip() for l in s.replace('\r','').splitlines() if l.strip() and not l.strip().startswith(('Attribute ','Option VBASupport','Rem Attribute VBA_ModuleType='))]

def run():
    start=time.perf_counter(); dest=Path('evidence'); dest.mkdir(exist_ok=True); Path('output/vba').mkdir(parents=True,exist_ok=True)
    profile=Path(tempfile.mkdtemp(prefix='macro-profile-')); (profile/'user').mkdir()
    (profile/'user/registrymodifications.xcu').write_text('''<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Filter/Microsoft/VBA/Excel"><prop oor:name="Load" oor:op="fuse"><value>true</value></prop><prop oor:name="Executable" oor:op="fuse"><value>true</value></prop><prop oor:name="Save" oor:op="fuse"><value>true</value></prop></item></oor:items>''')
    with socket.socket() as sock: sock.bind(('127.0.0.1',0)); port=sock.getsockname()[1]
    process=subprocess.Popen(['libreoffice','-env:UserInstallation='+profile.as_uri(),'--headless',f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext','--norestore','--nodefault','--nofirststartwizard'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    ctx=uno.getComponentContext(); resolver=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
    try:
        for _ in range(150):
            try: remote=resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'); break
            except Exception: time.sleep(.1)
        else: raise RuntimeError('Office listener did not start')
        desktop=remote.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',remote)
        work=Path('.build/vba-run')
        if work.exists(): shutil.rmtree(work)
        work.mkdir(parents=True)
        shutil.copyfile('macro-pack.xlsm',work/'macro-pack.xlsm')
        doc=desktop.loadComponentFromURL((work/'macro-pack.xlsm').resolve().as_uri(),'_blank',0,(prop('Hidden',True),prop('MacroExecutionMode',4),prop('UpdateDocMode',3)))
        assert doc
        libs=doc.BasicLibraries; print('Imported VBA libraries:',libs.getElementNames())
        library='VBAProject' if libs.hasByName('VBAProject') else libs.getElementNames()[0]
        libs.loadLibrary(library); imported=libs.getByName(library).getByName('Module1'); expected=Path('.build/embedded.bas').read_text()
        if normalized(imported)!=normalized(expected):
            import difflib
            print('\n'.join(list(difflib.unified_diff(normalized(expected),normalized(imported)))[:35]))
        assert normalized(imported)==normalized(expected), 'Embedded module differs from readable source'
        ref=openpyxl.load_workbook('output/after.xlsx'); source=openpyxl.load_workbook('input/input.xlsx'); actions=[]
        def values(name):
            s=doc.Sheets.getByName(name); cursor=s.createCursor(); cursor.gotoEndOfUsedArea(True); end=cursor.RangeAddress
            vals=s.getCellRangeByPosition(0,0,end.EndColumn,end.EndRow).getDataArray()
            return [list(r[:3]) for r in vals if any(v!='' for v in r)]
        def invoke(action):
            tick=time.perf_counter(); result=doc.getScriptProvider().getScript(f'vnd.sun.star.script:{library}.Module1.Run{action}?language=Basic&location=document').invoke((),(),())[0]
            print(action,result); assert str(result).startswith('OK'), result
            return round(time.perf_counter()-tick,3)
        def snapshot(name):
            doc.storeToURL(Path('output/vba',name+'.xlsx').resolve().as_uri(),(prop('FilterName','Calc MS Excel 2007 XML'),prop('Overwrite',True)))
        def compare_split():
            manifest=values('__MP_Split')[1:]
            assert {p.name for p in (work/'output/split').glob('group-*.xlsx')}=={r[1] for r in manifest}
            for group,name,count in manifest:
                path=work/'output/split'/name
                assert is_zipfile(path), 'Macro must write real XLSX, not renamed OLE'
                book=openpyxl.load_workbook(path)
                assert book.sheetnames==['Data']
                rows=table(book['Data'],REQUIRED['Data'])
                expected=[r for r in values('Data')[1:] if str(r[1])==str(group)]
                assert rows==expected and len(rows)==count
        def compare(name):
            actual=values(name); want=[list(r[:3]) for r in ref[name].values]
            if name=='__MP_Letters':
                for row in actual[1:]: row[1]=f'{float(row[1]):.2f}'
            assert actual==want, (name, len(actual),len(want), [(i,a,w) for i,(a,w) in enumerate(zip(actual,want)) if a!=w][:4])
        for action,sheet in [('Split','__MP_Split'),('Merge','__MP_Merged'),('Rename','__MP_Rename'),('Letters','__MP_Letters')]:
            snapshot(action.lower()+'-before')
            if action=='Letters':
                for path in Path('.build/vba-run/output/letters').glob('*.txt'): path.unlink()
            elapsed=invoke(action); compare(sheet)
            if action=='Split':
                compare_split()
                shutil.copytree(work/'output/split','output/vba/split',dirs_exist_ok=True)
            if action=='Rename':
                for old,new,status in list(ref['__MP_Rename'].values)[1:]:
                    if status=='renamed': assert values(new)==[list(r) for r in source[old].values]
                    elif old in source.sheetnames: assert values(old)==[list(r) for r in source[old].values]
            if action=='Letters':
                for rid,name,amount in list(source['Recipients'].values)[1:]:
                    assert Path('.build/vba-run/output/letters',rid+'.txt').read_text()==render_letter(Path('input/template.txt').read_text(),[rid,name,amount])+'\n'
            snapshot(action.lower()+'-after')
            actions.append({'action':action,'seconds':elapsed,'matches_reference':True,'snapshot':f'../output/vba/{action.lower()}-after.xlsx'})
        # Repeated actions preserve business data. The rename log accurately changes to missing source.
        names=list(doc.Sheets.getElementNames())
        for action,sheet in [('Split','__MP_Split'),('Merge','__MP_Merged'),('Letters','__MP_Letters')]: invoke(action); compare(sheet)
        compare_split()
        invoke('Rename'); assert set(names)==set(doc.Sheets.getElementNames())
        for source_name in REQUIRED:
            assert values(source_name)==[list(r) for r in source[source_name].values], source_name
        # A new group requires a new workbook, with no prepared group tab.
        data_sheet=doc.Sheets.getByName('Data')
        data_sheet.getCellRangeByName('A1202:C1203').setDataArray((('NEW','A/B',-7.5),('NEW','A/B',-7.5)))
        invoke('Split'); compare_split()
        assert len(values('__MP_Split'))==8
        assert (work/'output/split/group-007.xlsx').exists()
        (work/'output/split/retain.me').write_text('unrelated file')
        data_sheet.getCellRangeByName('A1202:C1203').clearContents(1023)
        invoke('Split'); compare_split(); compare('__MP_Split')
        assert not (work/'output/split/group-007.xlsx').exists()
        assert (work/'output/split/retain.me').read_text()=='unrelated file'
        saved_data=data_sheet.getCellRangeByName('A2:C1201').getDataArray()
        data_sheet.getCellRangeByName('A2:C1201').clearContents(1023)
        invoke('Split'); compare_split(); assert len(values('__MP_Split'))==1
        data_sheet.getCellRangeByName('A2:C1201').setDataArray(saved_data)
        invoke('Split'); compare_split(); compare('__MP_Split')
        # Removing a recipient removes only the prior generated letter, not a stray unrelated file.
        recipients=doc.Sheets.getByName('Recipients'); last=recipients.getCellRangeByName('A81:C81'); saved=last.getDataArray()
        Path('.build/vba-run/output/letters/retain.me').write_text('unrelated file')
        last.clearContents(1023); invoke('Letters')
        assert not Path('.build/vba-run/output/letters/C080.txt').exists()
        assert Path('.build/vba-run/output/letters/retain.me').read_text()=='unrelated file'
        last.setDataArray(saved); invoke('Letters'); compare('__MP_Letters'); Path('.build/vba-run/output/letters/retain.me').unlink()
        # Invalid identifiers must stop before mutating the prior letters.
        first=recipients.getCellRangeByName('A2'); original=first.getString(); first.setString('../outside')
        rejected=doc.getScriptProvider().getScript(f'vnd.sun.star.script:{library}.Module1.RunLetters?language=Basic&location=document').invoke((),(),())[0]
        assert str(rejected).startswith('ERROR'); assert len(list(Path('.build/vba-run/output/letters').glob('*.txt')))==80
        first.setString(original)
        doc.close(True); desktop.terminate()
        runtime={'system':platform.freedesktop_os_release()['PRETTY_NAME'],'python':platform.python_version(),'libreoffice':subprocess.check_output(['libreoffice','--version'],text=True).strip(),'openpyxl':openpyxl.__version__,'xlsxwriter':xlsxwriter.__version__,'pytest':pytest.__version__,'pillow':PIL.__version__,'vba_source_matches_import':True,'vba_source_sha256':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(Path('vba').glob('*.bas'))},'actions':actions,'repeat_preserves_business_data':True,'input_tables_preserved':True,'removed_recipient_letter_removed':True,'unrelated_file_preserved':True,'unsafe_recipient_rejected_before_write':True,'elapsed_seconds':round(time.perf_counter()-start,3),'split_files_are_ooxml':True,'split_new_group_and_duplicates_checked':True,'split_removed_group_and_empty_input_checked':True,'split_unrelated_file_preserved':True,'export_note':'LibreOffice executed the imported macros. SaveGroup uses its LibreOffice branch with the native OOXML filter; the Excel SaveAs branch is separate. Snapshots are UNO exports after macro execution. GIFs show sheet snapshots, not desktop recordings.'}
        (dest/'runtime.json').write_text(json.dumps(runtime,indent=2)+'\n')
        print('PASS actual VBA outputs match reference')
    finally:
        process.terminate()
        try: process.wait(timeout=10)
        except subprocess.TimeoutExpired: process.kill()
        shutil.rmtree(profile)
if __name__=='__main__': run()
