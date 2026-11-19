"""Write the case page from the current measured reports."""
import csv,json
from pathlib import Path
r=json.loads(Path('output/report.json').read_text())
a=json.loads(Path('output/accuracy.json').read_text())
v=json.loads(Path('evidence/runtime.json').read_text())
s=v['software'];n=a['denominator']
forms = r['input_forms']
size_line = (f'This case at its size ({forms} forms): £{10 * ((forms + 14) // 15)} fixed, £10 for each 15 pages.\n\n' if forms > 15 else '')
source=json.loads(Path('evidence/hero-source.json').read_text())
with Path('output/unresolved.csv').open() as f:
    queue=list(csv.DictReader(f))
assert len(queue)==a['reviewed']['unresolved']
examples=[queue[0],next(row for row in queue if row['image']==source['image'] and row['field']==source['field'])]
queue_text='\n'.join(f"| [{row['form_id']}](input/forms/{row['image']}) | {row['field'].lower()} |" for row in examples)
field_values={field['field']:field['value'] for field in source['fields']}
alt=f"{source['image']}: full saved row — form {field_values['Form ID']}, item code blank for checking, quantity {field_values['Quantity']}, amount {field_values['Amount']}."
title='Scans to a spreadsheet'
text=f'''# {title}

{r['input_forms']} forms to Excel, with doubtful entries left blank for checking.

[**{a['raw']['correct']}/{n} fields correct on first reading; {a['reviewed']['correct']}/{n} after {r['reviewed']['confirmations_applied']} supplied corrections.**](output/accuracy.json) {a['reviewed']['errors']} wrong values accepted in this batch against the answer key.

<img src="media/hero.png" width="560" alt="{alt}">

[Scan {source['form_id']}](input/forms/{source['image']}) → [complete saved row](output/reviewed.xlsx) · [image sources](evidence/hero-source.json)

**{len(queue)} fields still to check** · [Full queue](output/unresolved.csv). Two examples with uncertain readings:

| Form | Check against the scan |
|---|---|
{queue_text}

<sub>Case 05 of 8 in [Excel & VBA casebook](../README.md)</sub>

## What was wrong

- Clear figures and unreadable marks were mixed together in the scans.
- A guessed amount could look like a finished entry in the sheet.
- A blank cell alone would not tell the owner what needed checking.

## What I did

- I put form numbers, item codes, quantities and printed amounts into .xlsx workbooks.
- I left doubtful entries blank and listed where to find each one on the scan.
- I applied supplied corrections and kept a record of what changed.
- I made the process reusable for fresh batches in the same form layout, with a new workbook and check list each time. [Batch instructions](RUNNING.md).

## Proof

| Entries checked against the answer key | First reading | After corrections |
|---|---:|---:|
| Read correctly | [{a['raw']['correct']}](output/accuracy.json) | [{a['reviewed']['correct']}](output/accuracy.json) |
| Still need checking | [{a['raw']['unresolved']}](output/accuracy.json) | [{a['reviewed']['unresolved']}](output/accuracy.json) |
| Wrong values passed through as correct | [{a['raw']['errors']}](output/accuracy.json) | [{a['reviewed']['errors']}](output/accuracy.json) |

Each form has a form number, item code, quantity and amount: [{n} entries altogether](output/accuracy.json). The amount is copied as printed; the form has no unit price or calculated order total. See the [form layout](LAYOUT.md).

The [batch record](validation/batch.json) describes the clear, faint, blurred and obscured forms. [Every field comparison](output/accuracy-fields.csv) traces back to the [separate answer key](validation/answer-key.csv). These counts apply only to this batch. A readable-looking entry can still be flagged when the reader is unsure. The faint code in the picture remains blank for that reason.

**Download the scans and results:** [before: original scans](input/forms.zip) · [input inventory](output/input-inventory.xlsx) · [after: extracted workbook](output/extracted.xlsx) · [reviewed workbook](output/reviewed.xlsx) · [original check list](output/unclear-entries.csv) · [entries still to check](output/unresolved.csv) · [record of corrections](output/review-audit.csv).

[Source and output views](evidence/hero-source.json) · [readable scan beside its row](images/01-Readable.png) · [full run report](output/report.json).

The reusable process supports the [same printed form layout](LAYOUT.md). A different order layout needs its fields mapped before it can be processed.

## Try it yourself

**Optional technical check for a developer.** From this case folder, with the [setup](RUNNING.md) complete:

```bash
./run.sh
```

This rebuilds the workbooks and check lists, runs the tests, and prints PASS or FAIL. [Use a fresh batch](RUNNING.md).

## How it is tested

[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)

- [Business tests](tests/test_business.py) read fresh scans, check accepted values and require a reason for each blank.
- They check conflicting corrections, duplicate IDs, damaged files and repeat-run stability.
- [Recorded test results](evidence/final-tests.json) cover the business checks. The .xlsx workbooks were produced and checked with [Python {s['Python']} / openpyxl {s['openpyxl']} and {s['LibreOffice']}](evidence/runtime.json); workbook values matched after opening and saving in LibreOffice.

This batch has [{r['input_forms']} separate form images, one form per image](validation/batch.json). For this batch, each image counts as a page: [{r['input_forms']} pages](evidence/page-scope.json), above the advertised page limit below. A batch this size needs a written quote before work starts.

---

[£10 fixed, up to 15 pages, unclear entries listed for you, 2 working days.](https://ardholt.com/p/prices/) Order through the platform.

{size_line}More rows or pages: I quote the difference in writing before I start.
'''
Path('README.md').write_text(text)
