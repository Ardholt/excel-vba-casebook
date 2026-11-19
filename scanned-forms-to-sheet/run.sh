#!/usr/bin/env bash
set -euo pipefail
trap 'echo "FAIL: scanned forms run" >&2' ERR
case_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
input_dir="$case_dir/input"
output_dir="$case_dir/output"
custom=no
while (($#)); do
  case "$1" in
    --input) [[ $# -ge 2 ]] || exit 2; input_dir="$2"; custom=yes; shift 2;;
    --output) [[ $# -ge 2 ]] || exit 2; output_dir="$2"; shift 2;;
    *) echo 'Usage: ./run.sh [--input DIR] [--output DIR]' >&2; exit 2;;
  esac
done
input_dir="$(cd -- "$input_dir" && pwd)"
mkdir -p -- "$output_dir"
output_dir="$(cd -- "$output_dir" && pwd)"
case "$output_dir/" in "$input_dir/"*) echo 'Output must be separate from input' >&2; exit 2;; esac
case "$input_dir/" in "$output_dir/"*) echo 'Output must be separate from input' >&2; exit 2;; esac
docker build -t excel-vba-casebook "$case_dir/../docker"
args=(--output /result)
if [[ "$custom" == yes ]]; then args+=(--input /source); fi
docker run --rm --network none --user "$(id -u):$(id -g)" --tmpfs /tmp:rw -e HOME=/tmp -e PYTHONDONTWRITEBYTECODE=1 \
  --mount "type=bind,source=$case_dir,target=/work,readonly" \
  --mount "type=bind,source=$input_dir,target=/source,readonly" \
  --mount "type=bind,source=$output_dir,target=/result" \
  -w /work excel-vba-casebook bash scripts/inner.sh "${args[@]}"
