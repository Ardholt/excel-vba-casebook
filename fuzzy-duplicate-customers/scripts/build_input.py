"""Build the bundled customer register deterministically."""
import json
from pathlib import Path
from customers import workbook, write_csv, HEADERS

rows = []
words = ['Cedar', 'Maple', 'Willow', 'Alder', 'Birch', 'Aspen', 'Elm', 'Oak']
for n in range(1, 401):
    rows.append([f'C{n:04}', f'{words[(n-1)%8]} account {n:03}', f'AREA-{n:03}'])
for n in range(1, 81):
    original = rows[n-1]
    name = original[1]
    if n % 3 == 1:
        name = '  '+name.upper().replace(' ', '   ')+' '
    elif n % 3 == 2:
        name = name.replace('account', 'acount')
    else:
        name = name.replace('account', 'accounts')
    rows.append([f'C{400+n:04}', name, original[2].lower() if n % 2 else original[2]])
# A spelling error and shortened suffix in one explicitly approved pair.
rows[2][1] = 'Willow account 003 Ltd'
rows[402][1] = 'Willow acount 003 Limited'
# These are distinct account labels in the same area; similarity alone is insufficient.
rows[40][1] = 'Cedar account 041 North'
rows[440][1] = 'Cedar account 041 South'
# A blank area is deliberately not eligible for suggestions.
rows[79][2] = rows[479][2] = ''
workbook('input/input.xlsx', [('Customers', HEADERS, rows)])
write_csv(Path('input/confirmed.csv'), ['source_id', 'target_id'], [[f'C{400+n:04}', f'C{n:04}'] for n in range(1,41)])
Path('input/register.json').write_text(json.dumps({'source_records':len(rows), 'confirmed_links':40, 'must_remain_separate':['C0041','C0441'], 'blank_area_ids':['C0080','C0480'], 'approval_policy':'The bundled confirmation list is an explicit review decision; proposals are never copied into it automatically.'}, indent=2)+'\n')

import platform, openpyxl
Path('input/generation.json').write_text(json.dumps({'generator':'scripts/build_input.py','python':platform.python_version(),'openpyxl':openpyxl.__version__,'description':'Customer register with spelling variants and an explicit confirmation list.'},indent=2)+'\n')
