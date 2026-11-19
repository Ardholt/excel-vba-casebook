#!/usr/bin/env bash
set -euo pipefail
case_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
input_dir="$case_dir/input"
output_dir="$case_dir/output"
while (($#)); do
  case "$1" in
    --input|--output)
      (($# >= 2)) || { echo "FAIL: missing directory" >&2; exit 2; }
      if [[ "$1" == --input ]]; then input_dir="$2"; else output_dir="$2"; fi
      shift 2 ;;
    *) echo "Usage: ./run.sh [--input DIR --output DIR]" >&2; exit 2 ;;
  esac
done
input_dir="$(cd -- "$input_dir" && pwd -P)"
mkdir -p -- "$output_dir"
output_dir="$(cd -- "$output_dir" && pwd -P)"
if [[ "$input_dir" == "$output_dir" || "$input_dir/" == "$output_dir/"* || "$output_dir/" == "$input_dir/"* ]]; then
  echo "FAIL: input and output must be separate, non-nested directories" >&2
  exit 2
fi
trap 'echo "FAIL — container build or case run failed" >&2' ERR
docker build -t excel-vba-casebook "$case_dir/../docker"
docker run --rm --network none --user "$(id -u):$(id -g)" \
  --env HOME=/tmp --env PYTHONDONTWRITEBYTECODE=1 \
  --mount "type=bind,source=$case_dir,target=/case,readonly" \
  --mount "type=bind,source=$input_dir,target=/input,readonly" \
  --mount "type=bind,source=$output_dir,target=/output" \
  --workdir /case excel-vba-casebook \
  bash scripts/inner.sh --input /input --output /output
