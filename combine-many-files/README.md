# Combine monthly shop files

Separate shop spreadsheets, with columns in different orders. Bring them into one table, then refresh without doubling the figures.

**[50 files: 46 filled, 2 empty, 2 rejected](output/manifest.csv).** [1 temporary file ignored, outside the 50](output/report.json).

<p>
  <img src="images/after-hero.png" width="780" alt="Finished combined table: 2025-01-01, amount 9.08 from 01_2025-01_B01.xlsx; 2025-01-01, amount 9.79 from 02_2025-01_B02.xlsx.">
</p>

Two entries from the [finished table, with their original source rows](evidence/table-preview.json). [Download the finished spreadsheet](output/consolidated.xlsx).

<sub>Case 8 of 8 in [Excel & VBA casebook](../README.md)</sub>

## What was wrong

- Each shop's monthly figures are in a separate spreadsheet.
- The columns moved around, so copying and pasting could put figures in the wrong place.
- Empty files and files with missing columns needed attention before they could be included.

## What I did

- I matched columns by their names, keeping the extra details some shops supplied.
- I kept the original filename and row number beside each entry so it could be checked.
- I listed files that needed fixing, instead of silently losing them.
- I set up a repeatable update and checked that adding a file brings in its entries once.

## Proof

The opening table shows one entry from each of the first two shop files, with their original values retained. Open the [full combined spreadsheet](output/consolidated.xlsx) to see all 5,520 rows.

| What was checked | What happened |
|---|---|
| All the shop files | [50 files: 46 filled, 2 empty, 2 needing fixes](output/manifest.csv) |
| Files needing fixes | [One is missing the amount column; one is missing the data sheet](output/rejections.csv) |
| Temporary editing file | [1 ignored; separate from the 50 shop files](output/report.json) |
| Entries brought together | [5,520 rows; amounts add up to 238,441.92](output/evidence.txt) |
| Update again | [Still 5,520 rows, with every value unchanged](evidence/vba-run.json) |
| Add a file containing another entry | [5,521 rows, matching the enlarged source folder](evidence/vba-run.json) |

**Download:** [original shop files](https://github.com/Ardholt/excel-vba-casebook/tree/main/combine-many-files/input/files) · [finished spreadsheet](output/consolidated.xlsx) · [files needing attention](output/rejections.csv) · [workbook with the update macro](combine.xlsm).

<details markdown="1">
<summary>Watch the recorded refresh replay</summary>

<img src="images/refresh.gif" width="510" alt="RefreshFolder runs: 0 to 5,520 rows; run again: 5,520 to 5,520; add one row and run: 5,520 to 5,521.">

Replay of the [recorded refresh log](evidence/refresh-trace.json), rather than an application screen recording. [Still image](media/hero.png).

</details>

<details markdown="1">
<summary>Developer details: reproduce the results and inspect the checks</summary>

## Try it yourself

Optional technical check for a developer, from this folder with Docker running:

```bash
./run.sh
```

It starts in a fresh results folder, keeps the sources read-only, checks the results and prints `PASS: combine checks` plus the results path. [Developer setup and file requirements](USAGE.md).

## How it is tested

[![Tests](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml/badge.svg)](https://github.com/Ardholt/excel-vba-casebook/actions/workflows/ci.yml)

Run in: [LibreOffice Calc 24.2.7.2 (headless)](evidence/runtime.json), the macro checked against an independent Python reference.

[Passing tests](evidence/tests.json) check row counts, amount totals, every source value and unchanged input files. They cover reordered columns, extra fields, empty and invalid files, temporary files, repeat updates and a newly added file. The [actual VBA run](evidence/vba-run.json) matched the reference cell by cell; its saved workbook used UNO export rather than VBA SaveAs.

The refresh animation is a [replay of events captured around those macro calls](evidence/proof-media.json), not an application screen recording. The run starts with an empty table. It refreshes twice, then refreshes a temporary copy of the source folder with one extra entry. Every result is compared with the Python reference. Launcher tests use a Docker command double to check read-only mounts and fresh destinations; the full demo pipeline is also run in a fresh folder inside the Linux container.

For the developer checking the implementation: [tests](tests/test_combine.py) · [VBA source](vba/CombineFiles.bas) · [versions and execution evidence](evidence/runtime.json).

</details>


---

[£39 fixed as one task automated, 5 working days.](https://ardholt.com/p/prices/) This covers one shared folder-to-table setup across your shops. I confirm the files and layout through the platform before starting.

Contact me through the platform where you found this case.

For work beyond that setup, I quote any difference in writing before I start.
