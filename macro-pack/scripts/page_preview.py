from pathlib import Path
import subprocess,json
source=Path('README.md').read_text()
badge='https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg'
temp=Path('.preview.md')
try:
    temp.write_text(source.replace('[![Tests]('+badge+')](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)', ''))
    subprocess.run(['pagerender',str(temp),'evidence/page'],check=True)
    Path('evidence/page/preview-note.json').write_text(json.dumps({'source':'README.md','preview_note':'README content with only the remote CI badge omitted from this offline render. The public README retains the real badge URL; live CI status was not fetched.'},indent=2)+'\n')
finally:
    temp.unlink(missing_ok=True)
