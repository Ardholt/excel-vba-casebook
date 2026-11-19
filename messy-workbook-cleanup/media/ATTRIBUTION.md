# Hero sources and typography

The hero combines complete rows from this case’s LibreOffice worksheet renders with changes read from `output/changes.csv` and `output/duplicates.csv`. `scripts/build_hero.py` assembles it; `evidence/hero.json` records the crops and source hashes. No worksheet values are redrawn.

Captions use DejaVu Sans, under the [DejaVu fonts licence](https://dejavu-fonts.github.io/License.html). The worksheet renders use Carlito, under the SIL Open Font License 1.1. Both are supplied by the rendering container; the type is rendered into the image.
