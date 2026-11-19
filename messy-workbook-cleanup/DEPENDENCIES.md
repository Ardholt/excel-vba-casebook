# Running and rebuilding

Use the repository's shared `../docker/Dockerfile`; `run.sh` builds it as `excel-vba-casebook`. The base is Ubuntu 24.04. Docker is the only host prerequisite. Package versions from the executed container validation are recorded in [runtime evidence](evidence/runtime.json); the shared Dockerfile controls the image's installation and pinning.

The case code needs Python, openpyxl, LibreOffice and pytest. Rendering also needs Pillow, Poppler and the repository build helpers `xlrender`. The rendered proof images are included. Rendering is not needed to clean a workbook or run its checks. No Python dependencies are fetched during a case run.

Inside the shared image, from this case directory:

```bash
bash scripts/inner.sh
python3 -m pytest -q tests
```

To rebuild the included inputs, run `python3 scripts/generate_input.py`. This deliberately replaces the included input workbook and rules, so it is separate from the normal cleanup command. To rebuild the proof after a default run:

```bash
python3 scripts/build_proof.py
xlrender output/proof-before.xlsx output/render-before --sheet Before
xlrender output/proof-after.xlsx output/render-after --sheet After
python3 scripts/finish_images.py
```

The images are actual Calc sheet renders. The before view comes from the original list; the after view comes from the cleaned Records tab. The row audit traces which entries were kept. The GIF alternates those renders, beginning with the cleaned result. No external images, fonts or seed assets are distributed.

The Python dependency pins are in [requirements.txt](requirements.txt): openpyxl 3.1.2, pytest 7.4.4 and Pillow 10.2.0. The executed system packages were Ubuntu 24.04 and LibreOffice 24.2.7.2; use the shared Dockerfile to install the system packages. No font files or third-party image assets are distributed.
