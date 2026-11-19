# Runtime and asset licences

The repository's [shared Dockerfile](../docker/Dockerfile) is the container definition for `run.sh`. It uses Ubuntu 24.04, Python 3, openpyxl, Pillow, pytest, LibreOffice Calc and Tesseract with English data. The checked Python packages are pinned in [requirements.lock](requirements.lock). Python 3.12.3, Tesseract 5.3.4 with English data, and LibreOffice 24.2.7.2 were used, as recorded in [runtime evidence](evidence/runtime.json); this case does not install packages at run time or silently switch OCR engines. Retain the built image by its Docker digest to reproduce the exact container. The shared Dockerfile is maintained at repository level.

The form images are original assets. The PNG forms are drawn by [make_forms.py](scripts/make_forms.py), using DejaVu Sans from the runtime. Proof is rendered from the output workbook with LibreOffice/Poppler. No font binaries are bundled.

| Component | Licence | Use |
|---|---|---|
| Tesseract OCR and English trained data | Apache-2.0 | Reads field image pixels |
| openpyxl | MIT | Workbook production and checks |
| Pillow | HPND | Printed forms and field crops |
| pytest | MIT | Business checks |
| LibreOffice | MPL-2.0 / LGPL-3.0+ | Workbook round-trip checks and proof rendering |
| Poppler | GPL-2.0-or-later | Rasterises the proof PDF |
| Inter SemiBold, Instrument Serif, JetBrains Mono | SIL Open Font Licence 1.1 | Lettering in the before/after figure; no font files distributed |
| DejaVu Sans | Bitstream Vera terms, with DejaVu changes in public domain | Printed and rendered lettering |

The bundled assets and case source follow the repository licence. Consult installed packages' copyright notices for their complete third-party notices.
