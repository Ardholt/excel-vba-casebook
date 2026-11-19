#!/usr/bin/env bash
set -euo pipefail
mkdir -p images media
python3 scripts/prepare_views.py
for name in before after invoice; do
  xlrender ".render-tmp/$name.xlsx" ".render-tmp/$name"
  image=$(find ".render-tmp/$name" -maxdepth 1 -name '*.png' -print -quit)
  convert "$image" -strip "images/$name.png"
done
python3 scripts/render_hero.py
