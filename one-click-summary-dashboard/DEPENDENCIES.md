# Container and dependencies

Run `./run.sh` from this case in a clone or extracted download of the full repository. Prerequisite: a running Docker engine with permission to build and run containers. The first build may need internet access; the case container runs with networking disabled. No local Python or Office installation is required.

The wrapper builds [`../docker/Dockerfile`](../docker/Dockerfile) as `excel-vba-casebook`, mounts the case and input read-only and the selected output directory for writing, then invokes `scripts/inner.sh`. That script:

1. Builds and recalculates the Python reference using `xlrecalc` with macros disabled.
2. Uses `/opt/vba/bin/python scripts/package_macro.py` to build the macro workbook and its Forms button from the current source and selected inputs.
3. Runs `scripts/run_vba.py` with Python UNO. LibreOffice imports the VBA and the button binding. Accessibility activation presses the button for baseline, appended input and repeat checks, and again after saving and reopening the final `.xlsm`.
4. Preserves the VBA document module names in the exported OOXML package (`ThisWorkbook`, `Sheet1`–`Sheet4`), which Calc drops on save. Extracts the saved VBA to verify source identity, then runs pytest. Failure in any step fails the command.

`--input DIR` and `--output DIR` are passed to every data-producing step. Output contains fresh `sales-dashboard.xlsm`, `vba-before.xlsx`, `vba-after.xlsx`, `vba-repeat.xlsx`, reference workbooks, `report.json`, `runtime.json`, `macro-package.json`, `vba-run.json` and `tests.json`. The `.xlsx` files are value/formula snapshots for inspection; use the `.xlsm` to refresh. The optional before input must be a prefix of the final input; see [schema](input/SCHEMA.md).

Python uses openpyxl, XlsxWriter and pytest. The macro packaging library is pyOpenVBA 6.6.0 in `/opt/vba`, with [its template licence](vba/PACKAGING-LICENSE.txt). The container provides LibreOffice and Python UNO. [Runtime versions](output/runtime.json) and [button execution](output/vba-run.json) record the exercised software. The same inner command was run in the supplied container; the Docker wrapper builds and invokes that command.

Pillow typesets the proof graphic from saved workbook values. `xlrender` produces worksheet views, and `pagerender` previews the README. [Hero source](evidence/hero.json) and [page evidence](evidence/page-runtime.json) identify those artifacts. The [offline preview note](evidence/page-source.json) documents the remote badge substitution.
