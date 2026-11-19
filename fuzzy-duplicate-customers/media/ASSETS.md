# Visual assets

Sheet images are renders of the included review workbooks. The hero is a typeset extract of the two original customer entries, their saved confirmation row, the retained entry and its audit reason. It is not an interactive approval screen. [make_hero.py](../scripts/make_hero.py) reads the case files, verifies the join against the result workbook and audit, then produces [hero.svg](hero.svg) and `hero.png` with Chromium through Playwright. Regenerate it in the case container from the case folder with `python3 scripts/make_hero.py` after running the workflow. No external photographs or illustrations are used.

The hero uses the container's Liberation Sans and Liberation Mono fonts. Licence notices for the earlier casebook lettering are retained in [Inter](font-licenses/LICENSE-Inter.txt), [Instrument Serif](font-licenses/LICENSE-InstrumentSerif.txt) and [JetBrains Mono](font-licenses/LICENSE-JetBrainsMono.txt). Font binaries are not distributed here.
