#!/usr/bin/env bash
set -euo pipefail
case_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
input="$case_dir/input/files"
output=""
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
input=$(cd -- "$input" && pwd -P)
if [[ -z $output ]]; then
  mkdir -p -- "$case_dir/runs"
  output=$(mktemp -d "$case_dir/runs/run.XXXXXXXX")
fi
mkdir -p -- "$output"
output=$(cd -- "$output" && pwd -P)
if [[ $input == "$output" || $input == "$output/"* || $output == "$input/"* ]]; then
  echo "Input and output must be separate, non-nested folders" >&2; exit 2
fi
if [[ -n $(find "$output" -mindepth 1 -maxdepth 1 -print -quit) ]]; then
  echo "Output folder must be empty; choose a fresh folder" >&2; exit 2
fi
trap 'echo "FAIL: case checks did not complete" >&2' ERR
docker build -t excel-vba-casebook "$case_dir/../docker"
mounts=(-v "$case_dir:/case:ro" -v "$input:/source:ro" -v "$output:/result:rw")
if ((custom)); then
  docker run --rm --network none --user "$(id -u):$(id -g)" -e HOME=/tmp "${mounts[@]}" -w /case excel-vba-casebook bash scripts/inner.sh --input /source --output /result
else
  docker run --rm --network none --user "$(id -u):$(id -g)" -e HOME=/tmp "${mounts[@]}" -w /case excel-vba-casebook bash scripts/inner.sh --demo /source /result
fi
echo "Results: $output"
