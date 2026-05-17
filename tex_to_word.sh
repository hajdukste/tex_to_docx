#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CALL_DIR="$(pwd)"

usage() {
  cat <<'EOF'
Usage:
  ./tex_to_word.sh INPUT.tex [OUTPUT.docx] [-- PANDOC_ARGS...]

Converts a LaTeX manuscript to DOCX with Pandoc, then post-processes the DOCX
to make Word styles, page setup, line numbers, page breaks, and figure sizing
closer to the LaTeX-generated PDF.

Environment variables:
  DOCX_FONT            Word font to use (default: Latin Modern Roman)
  DOCX_BUILD_DIR       Build directory (default: INPUT_DIR/build)
  DOCX_REFERENCE_DOC   Reference DOCX template (default: build/latex-like-reference.docx)
  DOCX_BIBLIOGRAPHY    Comma-separated .bib files; auto-detected from \bibliography if unset
  DOCX_FIGURE_DIRS     Colon-separated figure dirs with PDFs/PNGs
                       (default: images-pdf-cropped:images-pdf:images beside INPUT.tex)
  DOCX_FIGURE_DPI      DPI for PDF figure previews (default: 300)

Examples:
  ./tex_to_word.sh paper.tex
  DOCX_FONT="Times New Roman" ./tex_to_word.sh paper.tex paper.docx
  DOCX_BIBLIOGRAPHY=refs.bib ./tex_to_word.sh paper.tex -- --csl nature.csl
EOF
}

resolve_from_call_dir() {
  case "$1" in
    /*) printf '%s\n' "$1" ;;
    *) printf '%s\n' "$CALL_DIR/$1" ;;
  esac
}

resolve_from_input_dir() {
  case "$1" in
    /*) printf '%s\n' "$1" ;;
    *) printf '%s\n' "$INPUT_DIR/$1" ;;
  esac
}

require_command() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing required command: $1" >&2
    exit 1
  fi
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

if [[ $# -lt 1 ]]; then
  usage >&2
  exit 2
fi

INPUT="$(resolve_from_call_dir "$1")"
shift

if [[ ! -f "$INPUT" ]]; then
  echo "Input TeX file not found: $INPUT" >&2
  exit 1
fi

INPUT_DIR="$(cd "$(dirname "$INPUT")" && pwd)"
INPUT_NAME="$(basename "$INPUT")"
INPUT="$INPUT_DIR/$INPUT_NAME"

if [[ $# -gt 0 && "${1:-}" != "--" ]]; then
  OUTPUT="$(resolve_from_call_dir "$1")"
  shift
else
  OUTPUT="$INPUT_DIR/${INPUT_NAME%.tex}.docx"
fi

if [[ $# -gt 0 && "${1:-}" == "--" ]]; then
  shift
fi

EXTRA_PANDOC_ARGS=("$@")

require_command pandoc
require_command python3
require_command pdftocairo
require_command pdfinfo

FONT="${DOCX_FONT:-Latin Modern Roman}"
BUILD_DIR="${DOCX_BUILD_DIR:-$INPUT_DIR/build}"
DOCX_FIGURE_DPI="${DOCX_FIGURE_DPI:-300}"
DOCX_FIGURE_DIR="$BUILD_DIR/docx-figures"
DEFAULT_REFERENCE="$SCRIPT_DIR/build/latex-like-reference.docx"
REFERENCE="${DOCX_REFERENCE_DOC:-$DEFAULT_REFERENCE}"
REFERENCE_RAW="$BUILD_DIR/pandoc-default-reference.docx"
GENERATED_REFERENCE="$BUILD_DIR/latex-like-reference.docx"
UNPOLISHED="$BUILD_DIR/$(basename "${OUTPUT%.docx}").raw.docx"

mkdir -p "$BUILD_DIR"
rm -rf "$DOCX_FIGURE_DIR"
mkdir -p "$DOCX_FIGURE_DIR"

if [[ ! -f "$REFERENCE" || "${DOCX_REBUILD_REFERENCE:-0}" == "1" ]]; then
  pandoc --print-default-data-file reference.docx > "$REFERENCE_RAW"
  python3 "$SCRIPT_DIR/tools/polish_pandoc_docx.py" "$REFERENCE_RAW" "$GENERATED_REFERENCE" --font "$FONT"
  REFERENCE="$GENERATED_REFERENCE"
fi

FIGURE_DIRS_RAW="${DOCX_FIGURE_DIRS:-images-pdf-cropped:images-pdf:images}"
IFS=':' read -r -a FIGURE_DIR_ITEMS <<< "$FIGURE_DIRS_RAW"
FIGURE_RESOURCE_PATHS=("$INPUT_DIR" "$DOCX_FIGURE_DIR")

for figure_dir_raw in "${FIGURE_DIR_ITEMS[@]}"; do
  [[ -n "$figure_dir_raw" ]] || continue
  figure_dir="$(resolve_from_input_dir "$figure_dir_raw")"
  [[ -d "$figure_dir" ]] || continue
  FIGURE_RESOURCE_PATHS+=("$figure_dir")
  while IFS= read -r -d '' pdf; do
    stem="$(basename "$pdf" .pdf)"
    pdftocairo -png -singlefile -r "$DOCX_FIGURE_DPI" "$pdf" "$DOCX_FIGURE_DIR/$stem"
  done < <(find "$figure_dir" -maxdepth 1 -type f -iname '*.pdf' -print0)
done

BIB_ARGS=()
if [[ -n "${DOCX_BIBLIOGRAPHY:-}" ]]; then
  IFS=',' read -r -a BIB_ITEMS <<< "$DOCX_BIBLIOGRAPHY"
  for bib_raw in "${BIB_ITEMS[@]}"; do
    [[ -n "$bib_raw" ]] || continue
    BIB_ARGS+=("--bibliography=$(resolve_from_input_dir "$bib_raw")")
  done
else
  while IFS= read -r bib; do
    [[ -n "$bib" ]] || continue
    BIB_ARGS+=("--bibliography=$bib")
  done < <(python3 - "$INPUT" <<'PY'
import re
import sys
from pathlib import Path

tex = Path(sys.argv[1])
text = tex.read_text(encoding="utf-8")
seen = set()

for body in re.findall(r"\\bibliography\{([^}]+)\}", text):
    for item in body.split(","):
        name = item.strip()
        if not name:
            continue
        path = (tex.parent / name).with_suffix(".bib")
        if path.exists() and path not in seen:
            seen.add(path)
            print(path)

for body in re.findall(r"\\addbibresource(?:\[[^\]]*\])?\{([^}]+)\}", text):
    path = tex.parent / body.strip()
    if path.exists() and path not in seen:
        seen.add(path)
        print(path)
PY
)
fi

RESOURCE_PATH="$(IFS=:; echo "${FIGURE_RESOURCE_PATHS[*]}")"

PANDOC_ARGS=(
  "$INPUT"
  --from=latex
  --to=docx
)

if [[ ${#BIB_ARGS[@]} -gt 0 ]]; then
  PANDOC_ARGS+=("--citeproc" "--metadata=reference-section-title:References")
  PANDOC_ARGS+=("${BIB_ARGS[@]}")
fi

PANDOC_ARGS+=(
  "--lua-filter=$SCRIPT_DIR/tools/docx_pdf_images_to_png.lua"
  "--lua-filter=$SCRIPT_DIR/tools/move_refs_before_backmatter.lua"
  "--reference-doc=$REFERENCE"
  "--resource-path=$RESOURCE_PATH"
)

if [[ ${#EXTRA_PANDOC_ARGS[@]} -gt 0 ]]; then
  PANDOC_ARGS+=("${EXTRA_PANDOC_ARGS[@]}")
fi

PANDOC_ARGS+=("--output=$UNPOLISHED")

pandoc "${PANDOC_ARGS[@]}"

python3 "$SCRIPT_DIR/tools/polish_pandoc_docx.py" "$UNPOLISHED" "$OUTPUT" --font "$FONT" --tex "$INPUT"

echo "Wrote $OUTPUT"
