# Sheets and letters by hand? 1,200 rows put in order

<p align="center"><img src="media/hero.png" width="600" alt="Before: account names and amounts waiting to become letters. After: a complete letter with the account's own name and amount, ready to review."></p>

**[1,200 rows sorted into 6 sheets, 840 rows combined and 80 letters prepared](output/report.json). [2 tab names changed; 5 requests left unchanged with reasons](output/renames.csv).**

<sub>Case 07 of 8 in [Excel & VBA casebook](../README.md)</sub>

## What was wrong

- Sorting rows into separate sheets by hand made it easy to miss one.
- Combining tables put values under the wrong headings, and some requested tab names could not be used.
- Each letter needed the right account name and amount, without copying another account's details.

## What I did

- I made a separate sheet for each group and checked that every row was kept.
- I combined the tables by their headings, keeping repeated rows rather than silently removing them.
- I changed the usable tab names and left the others untouched, with a reason for each.
- I prepared a complete letter for each account, ready to review. The letters are saved as files; they are not sent.

## Proof

**Separate sheets, with every row accounted for.** The [group list](output/split.csv) shows all [6 groups and 1,200 rows](output/report.json). Each group has its own tab in the [finished workbook](output/after.xlsx).

![All six groups, their separate sheet names and their row counts](images/sheets/01-split-after.png)

**One table with the headings lined up.** [480 rows plus 360 rows become 840 rows](output/report.json), including [12 repeated rows kept intact](output/report.json). These are the opening rows; [download the complete table](output/merged.csv).

![Opening rows of the finished combined table](images/sheets/01-merge-after.png)

**Every requested tab change explained.** All [7 requests](output/renames.csv) appear below: [2 names changed and 5 left unchanged](output/report.json). No existing tab was overwritten.

![All seven requested tab changes, including the two accepted and five rejected requests](images/sheets/01-rename-after.png)

**A finished letter for each account.** The hero shows [the first complete letter](output/letters/C001.txt), with its own name and amount. [Download all 80 letters](output/letters.zip), or read [the last letter](output/letters/C080.txt). The [recipient list](input/input.xlsx) lets you check each one.

**Workbooks:** [before](input/input.xlsx) · [after](output/after.xlsx) · [macro workbook](macro-pack.xlsm). **Details:** [letter template](input/template.txt) · [result report](output/report.json) · [run log](output/run.txt).

## Try it yourself

Optional technical check for a developer, from this folder with Docker installed:

```bash
./run.sh
```

It rebuilds the reference output and runs the checks. [Setup, input format and macro instructions](USAGE.md).

## How it is tested

[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)

Run in: LibreOffice Calc 24.2.7.2 (headless), each macro checked against independent reference outputs.

The [tests](tests/test_business.py) check row and value conservation, header alignment, duplicate preservation, rejected renames, exact letter contents and repeat runs. [Test results](evidence/tests.txt) · [versions, timings and VBA checks](evidence/runtime.json).

The checked macro workbook has its output tabs prepared. Creating tabs for newly added groups is outside the validated VBA path. The command above runs the Python reference and its tests; the [separate VBA runner](scripts/run_vba.py) checks the actual macros.

The retained before/after animations use workbook snapshots from those macro runs: [split](images/split.gif), [merge](images/merge.gif), [rename](images/rename.gif), [letters](images/letters.gif). The full [split](output/vba/split-after.xlsx), [merge](output/vba/merge-after.xlsx), [rename](output/vba/rename-after.xlsx) and [letter](output/vba/letters-after.xlsx) snapshots are available to inspect.

Readable VBA: [split](vba/SplitRows.bas) · [merge](vba/MergeTables.bas) · [rename](vba/RenameSheets.bas) · [letters](vba/MakeLetters.bas).

---

[£39 a task automated, 5 working days.](https://ardholt.com/p/prices/) Order through the platform.
More rows or pages: I quote the difference in writing before I start.
