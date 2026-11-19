# Reproducing the case

Run `./run.sh` from this folder in the full repository. Requires Bash and Docker. The shared `../docker/Dockerfile` builds `excel-vba-casebook`; it is the repository's container/dependency definition, based on Ubuntu 24.04. This case does not fetch or install packages during execution.

The [execution record](runtime.json) records the tested Python, openpyxl, pytest and LibreOffice versions. [requirements.lock](../requirements.lock) pins the Python package versions from that container. The shared image is owned by the repository; this case does not introduce a separate image or claim an image digest that was not supplied. Its container definition supplies the operating-system packages.

Inside the shared image, run:

```bash
bash scripts/inner.sh
python3 -m pytest -q -p no:cacheprovider tests
```

The pipeline writes `ledger.xlsx`, `ledger-checked.xlsx`, `input-overview.xlsx`, the CSV evidence, `report.json`, `reconciliation.md`, `runtime.json` and `checks.txt` under the output directory. It recalculates/saves through LibreOffice and compares workbook cell values against the original Python output. The second command independently reruns the business test suite. [Captured checks](../output/checks.txt).

`output/runtime.json` records the latest pipeline run. `evidence/runtime.json` captures the build used for this page. File hashes tie that capture to the statements, code and checked outputs. Rerunning changes output timestamps and runtime; business values stay the same. Actual elapsed times are recorded, not time-saving estimates.

`python3 scripts/check_paths.py` exercises the command with replacement statement values and input/output directories containing spaces, verifies source files are unchanged, then repeats the run and compares business results. [Results](input-options.json).

`python3 scripts/render_proof.py` captures the proof sheets after a successful default run. It isolates each requested worksheet before calling the sheet renderer so other sheets' print areas cannot enter its image. [Image record](images.json). Custom-input runs do not change the published images or README.

`python3 scripts/capture_page.py` invokes `pagerender` and captures the README body using Chromium and GitHub-style Markdown CSS at phone and desktop widths. The capture manifest is written to `evidence/page-preview/page-manifest.json`. The preview is an approximation of GitHub's README body, not a live GitHub or Pages browser visit. If the required remote CI badge cannot load, the preview alone omits that badge image; the exact workflow badge URL remains unchanged in README.md. [Capture disclosure](page-run.json).

The workbooks contain no formula-derived money totals. LibreOffice's round-trip verifies saved cell values. The clean-host Docker wrapper, main-branch CI, and live GitHub/Pages acceptance belong to the full repository integration.

A failed execution writes a FAIL runtime/check summary when the output directory exists and returns a nonzero exit code. It does not certify any workbook left from an earlier run. Fix the input issue and rerun before using the workbook as a newly checked result.

`python3 scripts/page_pass.py` rebuilds the compact hero from the source statements, ledger CSV and reconciliation CSV, records its provenance in `evidence/hero.json`, then captures the README. The hero is a typeset data summary using Pillow and the container’s DejaVu Sans fonts; [source rows and display conventions](../media/README.md) distinguish it from the worksheet renders. The page-preview tool uses Playwright (Apache-2.0), Chromium (BSD and bundled third-party licences) and github-markdown-css (MIT); these are rendering dependencies, not assets copied into the case.
