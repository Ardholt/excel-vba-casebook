#!/usr/bin/env bash
set -euo pipefail
mkdir -p images media
python3 scripts/prepare_views.py
for name in before after invoice; do
  xlrender ".render-tmp/$name.xlsx" ".render-tmp/$name"
  image=$(find ".render-tmp/$name" -maxdepth 1 -name '*.png' -print -quit)
  convert "$image" -strip "images/$name.png"
done
bash ../tools/hero/make-hero images/before.png images/after.png media/hero.png \
  --title '120 invoices checked. What is still owed?' \
  --before-note 'Totals and payments unchecked' \
  --after-note 'Paid, still owed or credit — listed' \
  --before-where 'input.xlsx · Overview' \
  --after-where 'checked-ledger.xlsx · Payment list' \
  --tag 'EXCEL CASEBOOK · CASE 06'
