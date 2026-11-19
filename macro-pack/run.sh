#!/usr/bin/env bash
set -euo pipefail
root=$(cd -- "$(dirname -- "$0")" && pwd -P)
in="$root/input"; out="$root/output"
while (($#)); do
 case "$1" in
 --input) in=$2; shift 2;;
 --output) out=$2; shift 2;;
 *) echo 'Usage: ./run.sh [--input DIR --output DIR]' >&2; exit 2;;
 esac
done
in=$(cd -- "$in" && pwd -P)
mkdir -p -- "$out"
out=$(cd -- "$out" && pwd -P)
if [[ "$in" == "$out" || "$out/" == "$in/"* || "$in/" == "$out/"* ]]; then echo 'Input and output must be separate folders' >&2; exit 2; fi
trap 'echo "FAIL: macro pack run" >&2' ERR
docker build -t excel-vba-casebook "$root/../docker"
docker run --rm --network none --user "$(id -u):$(id -g)" --mount "type=bind,src=$root,dst=/case,readonly" --mount "type=bind,src=$in,dst=/input,readonly" --mount "type=bind,src=$out,dst=/output" -w /case excel-vba-casebook bash scripts/inner.sh --input /input --output /output
