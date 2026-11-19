#!/usr/bin/env bash
set -euo pipefail
root=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd -P)
input="$root/input"
output="$root/output"
while (($#)); do
  case "$1" in
    --input|--output)
      if (($# < 2)); then echo "Missing directory argument" >&2; exit 2; fi
      if [[ "$1" == --input ]]; then input=$2; else output=$2; fi
      shift 2 ;;
    *) echo "Usage: ./run.sh [--input DIR --output DIR]" >&2; exit 2 ;;
  esac
done
command -v docker >/dev/null || { echo 'FAIL: Docker is required' >&2; exit 1; }
input=$(CDPATH= cd -- "$input" && pwd -P)
[[ -f "$input/input.xlsx" ]] || { echo 'FAIL: input.xlsx is missing' >&2; exit 1; }
# Build the shared environment; no host Python is needed.
trap 'echo "FAIL: invoice checks stopped" >&2' ERR
docker build -t excel-vba-casebook "$root/../docker"
mkdir -p -- "$output"
output=$(CDPATH= cd -- "$output" && pwd -P)
case "$output/" in "$input/"*|"$root/" ) echo 'FAIL: output overlaps source' >&2; exit 2 ;; esac
case "$input/" in "$output/"*) echo 'FAIL: output contains input' >&2; exit 2 ;; esac
docker run --rm --network none --user "$(id -u):$(id -g)" \
  -e PYTHONDONTWRITEBYTECODE=1 -e HOME=/tmp \
  --mount "type=bind,src=$root,dst=/case,readonly" \
  --mount "type=bind,src=$input,dst=/source,readonly" \
  --mount "type=bind,src=$output,dst=/result" \
  -w /case excel-vba-casebook bash scripts/inner.sh --input /source --output /result
