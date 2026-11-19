#!/usr/bin/env bash
set -euo pipefail
trap 'printf "FAIL: case run failed.\n" >&2' ERR
case_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
input_dir="$case_dir/input"
output_dir="$case_dir/output"
while (( $# )); do
  case "$1" in
    --input|--output)
      if (( $# < 2 )); then printf 'FAIL: missing directory argument.\n' >&2; exit 2; fi
      if [[ "$1" == --input ]]; then input_dir="$2"; else output_dir="$2"; fi
      shift 2 ;;
    *) printf 'Usage: ./run.sh [--input DIR --output DIR]\n' >&2; exit 2 ;;
  esac
done
input_dir="$(cd -- "$input_dir" && pwd -P)"
mkdir -p -- "$output_dir"
output_dir="$(cd -- "$output_dir" && pwd -P)"
if [[ "$input_dir" == "$output_dir" || "$input_dir/" == "$output_dir/"* || "$output_dir/" == "$input_dir/"* ]]; then
  printf 'FAIL: input and output must be separate, non-nested directories.\n' >&2; exit 2
fi
docker build -t excel-vba-casebook "$case_dir/../docker"
docker run --rm --network none --user "$(id -u):$(id -g)" \
  -e PYTHONDONTWRITEBYTECODE=1 \
  -v "$case_dir:/case:ro" -v "$input_dir:/data/input:ro" -v "$output_dir:/data/output:rw" \
  -w /case excel-vba-casebook bash scripts/inner.sh --input /data/input --output /data/output
