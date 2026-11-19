# Shared spreadsheet environment

Each case's `./run.sh` builds this directory as `excel-vba-casebook`, then runs
with the supplied input and output directories. Docker is the only host
prerequisite. The image needs internet access to build; case runs use no network.

- `xlrecalc IN.xlsx OUT.xlsx` recalculates with LibreOffice and writes XLSX.
- `xlrender IN.xlsx OUTDIR [--sheet NAME]` exports sheets to PDF and PNG.
- `pagerender README.md OUTDIR` captures Markdown at 390 and 1366 pixels.
- `/opt/vba/bin/python` provides pyOpenVBA for workbook packaging.

Calculation and rendering disable macros. A Calc result does not verify VBA or
desktop Excel behaviour. XLSX recalculation does not preserve macros.

The page renderer requires every image to resolve, including external badges.
It approximates the GitHub Markdown body. It does not build the Jekyll Pages site.

CI caches build layers in GitHub Actions and loads the image on each case runner.
It uses no package registry. [Dependency licences](LICENSES.md).
