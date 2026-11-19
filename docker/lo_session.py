"""Private, disposable LibreOffice profile and local UNO pipe, never TCP."""
import os, pathlib, subprocess, tempfile, time, uuid
import uno
from com.sun.star.beans import PropertyValue

def prop(name,value):
    p=PropertyValue(); p.Name=name; p.Value=value; return p

def url(path): return pathlib.Path(path).resolve().as_uri()

class Office:
    def __enter__(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='xl-lo-')
        self.pipe='xl_'+uuid.uuid4().hex
        self.process=subprocess.Popen(['libreoffice','--headless','--nologo','--nodefault',
            '--nofirststartwizard','--norestore',f'-env:UserInstallation={url(self.tmp.name)}',
            f'--accept=pipe,name={self.pipe};urp;StarOffice.ComponentContext'],stdout=subprocess.DEVNULL)
        local=uno.getComponentContext()
        resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        for _ in range(200):
            try:
                self.ctx=resolver.resolve(f'uno:pipe,name={self.pipe};urp;StarOffice.ComponentContext');break
            except Exception:
                if self.process.poll() is not None: raise RuntimeError('LibreOffice exited during startup')
                time.sleep(.1)
        else: raise TimeoutError('LibreOffice UNO startup exceeded 20s')
        self.desktop=self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',self.ctx)
        return self

    def load(self,path,macros=False):
        d=self.desktop.loadComponentFromURL(url(path),'_blank',0,(
            prop('Hidden',True),prop('ReadOnly',False),prop('UpdateDocMode',3),
            prop('MacroExecutionMode',4 if macros else 0)))
        if d is None: raise RuntimeError(f'Cannot load {path}')
        return d

    def __exit__(self,*exc):
        try:self.desktop.terminate()
        except Exception:pass
        try:self.process.wait(timeout=5)
        except subprocess.TimeoutExpired:self.process.kill();self.process.wait()
        self.tmp.cleanup()
