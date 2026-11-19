"""Offline Pages first screens using the repository's captured site stylesheet."""
from pathlib import Path
import hashlib
import json
from markdown_it import MarkdownIt
from playwright.sync_api import sync_playwright
root=Path('.').resolve()
source=Path('README.md').read_text()
source=source.replace('![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)','Test workflow')
css=Path('evidence/pages-style.css').read_text()
body=MarkdownIt('gfm-like',{'html':True}).render(source)
html='<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><style>'+css+'</style><header class="top"><a class="brand">ARDHOLT</a><a class="book">Excel &amp; VBA casebook · all cases</a></header><main class="md">'+body+'</main>'
target=Path('.pages-preview.html');target.write_text(html)
out=Path('evidence/pages-preview');out.mkdir(exist_ok=True)
try:
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True,args=['--no-sandbox'])
        for width,height in [(390,844),(1366,900)]:
            page=browser.new_page(viewport={'width':width,'height':height},device_scale_factor=1)
            page.goto(target.resolve().as_uri(),wait_until='networkidle')
            assert page.locator('img').evaluate_all('(xs)=>xs.every(x=>x.complete && x.naturalWidth)')
            page.screenshot(path=str(out/f'{width}-first.png'))
        browser.close()
finally:target.unlink()
(out/'manifest.json').write_text(json.dumps({'readme_sha256':hashlib.sha256(Path('README.md').read_bytes()).hexdigest(),
    'css_sha256':hashlib.sha256(css.encode()).hexdigest(),
    'scope':'Offline Pages body with captured repository CSS and text header; excludes analytics and remote CI badge. Not a live-site visit.'},indent=2)+'\n')
