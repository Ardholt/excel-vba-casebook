"""Render the unchanged README with an offline response for its remote CI badge."""
import hashlib,json,runpy,shutil,sys,os,shlex
from pathlib import Path
try:
    from playwright.sync_api import Page
except ModuleNotFoundError:
    if os.environ.get('CASE_PREVIEW_REEXEC'):
        raise
    interpreter=shlex.split(Path(shutil.which('pagerender')).read_text().splitlines()[0][2:])
    os.environ['CASE_PREVIEW_REEXEC']='1'
    os.execv(interpreter[0],interpreter+[str(Path(__file__).resolve())])
badge='https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg'
svg='<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"></svg>'
original_goto,original_set=Page.goto,Page.set_content
def install(page):
    page.route(badge,lambda route:route.fulfill(status=200,content_type='image/svg+xml',body=svg))
def goto(self,*args,**kwargs):
    install(self)
    return original_goto(self,*args,**kwargs)
def set_content(self,*args,**kwargs):
    install(self)
    return original_set(self,*args,**kwargs)
Page.goto,Page.set_content=goto,set_content
sys.argv=['pagerender','README.md','page-preview']
try:
    runpy.run_path(shutil.which('pagerender'),run_name='__main__')
finally:
    Page.goto,Page.set_content=original_goto,original_set
Path('evidence/page-preview.json').write_text(json.dumps({'source':'README.md','source_sha256':hashlib.sha256(Path('README.md').read_bytes()).hexdigest(),'renderer':'pagerender','viewports':[[390,844],[1366,900]],'remote_badge':{'url':badge,'preview_response':'Remote badge omitted visually in offline capture; no replacement button or status','reason':'Network is disabled in the execution container. README and its badge URL are unchanged.'},'scope':'GitHub-style body preview; no live site or live CI-status claim.'},indent=2)+'\n')
