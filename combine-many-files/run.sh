#!/usr/bin/env bash
set -euo pipefail
case_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
input="$case_dir/input/files"
output="$case_dir/output"
custom=0
while (($#)); do
  case "$1" in
    --input|--output)
      (($# >= 2)) || { echo "Missing folder argument" >&2; exit 2; }
      if [[ $1 == --input ]]; then input=$2; else output=$2; fi
      custom=1; shift 2 ;;
    *) echo "Use: ./run.sh [--input DIR --output DIR]" >&2; exit 2 ;;
  esac
done
input=$(cd -- "$input" && pwd)
mkdir -p -- "$output"
output=$(cd -- "$output" && pwd)
if [[ $input == "$output" || $input == "$output/"* || $output == "$input/"* ]]; then
  echo "Input and output must be separate, non-nested folders" >&2; exit 2
fi
trap 'echo "FAIL: case checks did not complete" >&2' ERR
docker build -t excel-vba-casebook "$case_dir/../docker"
if ((custom)); then
  docker run --rm --network none --user "$(id -u):$(id -g)" -e HOME=/tmp -v "$case_dir:/work:ro" -v "$input:/source:ro" -v "$output:/result" -w /work excel-vba-casebook bash scripts/inner.sh --input /source --output /result
else
  docker run --rm --network none --user "$(id -u):$(id -g)" -e HOME=/tmp -v "$case_dir:/work" -w /work excel-vba-casebook bash scripts/inner.sh
fi
