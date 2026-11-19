# Container dependencies

The repository's own code is MIT licensed. Installed software retains its own
licence. Ubuntu package notices and full terms are in
`/usr/share/doc/<package>/copyright` inside the image; Python distributions
include their licence files in their installed metadata. Those files govern
the particular build, including transitive dependencies.

| Component | Licence |
|---|---|
| Ubuntu base system | Multiple free software licences, including GPL, LGPL, BSD and MIT |
| LibreOffice Calc and Python UNO bindings | MPL-2.0; LGPL-3.0-or-later where applicable |
| Python 3 | PSF-2.0 and bundled component notices |
| pip | MIT |
| openpyxl | MIT |
| XlsxWriter | BSD-2-Clause |
| pandas | BSD-3-Clause |
| NumPy | BSD-3-Clause and bundled notices |
| pytest | MIT |
| Pillow | HPND |
| Tesseract and English language data | Apache-2.0 |
| Leptonica | BSD-2-Clause |
| ImageMagick | ImageMagick licence |
| FFmpeg | LGPL-2.1-or-later; Ubuntu's enabled GPL components add GPL terms |
| Poppler utilities | GPL-2.0-or-later |
| Git | GPL-2.0 |
| rsync | GPL-3.0-or-later |
| curl | curl licence |
| CA certificates | MPL-2.0 for Mozilla certificate material, with package notices |
| DejaVu fonts | Bitstream Vera licence; DejaVu changes in the public domain |
| Liberation, Carlito and Caladea fonts | SIL Open Font License 1.1 |
| pyOpenVBA 6.6.0 | MIT |
| Playwright 1.56.0 | Apache-2.0 |
| Chromium supplied by Playwright | BSD-3-Clause and bundled third-party licences |
| markdown-it-py 4.0.0 | MIT |
| mdit-py-plugins 0.5.0 | MIT |
| linkify-it-py 2.0.3 | MIT |
| github-markdown-css 5.8.1 | MIT; full text at `/opt/page/github-markdown-css-LICENSE` |

The base release is Ubuntu 24.04. Distribution packages follow that release's
security updates; their patch versions are not frozen. Python preview packages
and pyOpenVBA are version pinned, with a SHA-256 check for the pyOpenVBA wheel.
Playwright selects its matching Chromium revision. Case execution reports record
the versions actually used.
