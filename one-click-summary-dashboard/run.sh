#!/usr/bin/env bash
set -euo pipefail
trap 'echo "FAIL: dashboard run" >&2' ERR
case_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
input_dir="$case_dir/input"
output_dir="$case_dir/output"
while (($#)); do
  case "$1" in
    --input|--output)
      (($# >= 2)) || { echo "Missing directory for $1" >&2; exit 2; }
      if [[ "$1" == --input ]]; then input_dir="$2"; else output_dir="$2"; fi
      shift 2 ;;
    *) echo 'Usage: ./run.sh [--input DIR] [--output DIR]' >&2; exit 2 ;;
  esac
done
input_dir=$(cd -- "$input_dir" && pwd -P)
mkdir -p -- "$output_dir"
output_dir=$(cd -- "$output_dir" && pwd -P)
if [[ "$input_dir" == "$output_dir" || "$input_dir/" == "$output_dir/"* || "$output_dir/" == "$input_dir/"* ]]; then
  echo 'Input and output directories must be separate' >&2
  exit 2
fi
docker build -t excel-vba-casebook "$case_dir/../docker"
docker run --rm --network none --user "$(id -u):$(id -g)" \
  -v "$case_dir:/case:ro" -v "$input_dir:/input:ro" -v "$output_dir:/output:rw" \
  -w /case excel-vba-casebook bash scripts/inner.sh --input /input --output /output
echo 'PASS: fresh macro workbook, three LibreOffice button refreshes and business tests'
