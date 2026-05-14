#!/usr/bin/env bash
# Download OOXML / ECMA-376 specs and Microsoft Open Specifications to the
# local references/specs/downloads/ tree. Not auto-run; CI does not need this.
#
# Usage:
#   ./fetch.sh            # actually download (idempotent — skip already-present)
#   ./fetch.sh --dry-run  # print what would be downloaded
#
# We do NOT redistribute these documents — that's why they're gitignored.
# Re-fetch when ECMA or Microsoft publishes an update.
set -eu

dry_run=0
case "${1:-}" in
  --dry-run) dry_run=1 ;;
  '') ;;
  *) echo "unknown arg: $1" >&2; exit 2 ;;
esac

here="$(cd "$(dirname "$0")" && pwd)"
dl="$here/downloads"

ecma_base="https://ecma-international.org/wp-content/uploads"
specs=(
  # ECMA-376 5th edition (current as of 2026). Each ZIP contains PDFs +
  # XML schemas (XSD). Part 1 is the one with PresentationML; Part 2 is OPC.
  "ecma-376/Part1.zip|$ecma_base/ECMA-376-1_5th_edition_december_2016.zip"
  "ecma-376/Part2.zip|$ecma_base/ECMA-376-2_5th_edition_december_2021.zip"
  "ecma-376/Part3.zip|$ecma_base/ECMA-376-3_5th_edition_december_2015.zip"
  "ecma-376/Part4.zip|$ecma_base/ECMA-376-4_5th_edition_december_2016.zip"
)

# Microsoft Open Specifications — HTML hub pages (no canonical PDF on the open
# web). We mirror the top-level URL for offline reference; deeper crawl is
# manual. These docs supplement ECMA-376 with PowerPoint-specific behavior.
ms_specs=(
  "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/"
  "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/"
  "https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/"
)

run() {
  if [ "$dry_run" -eq 1 ]; then
    echo "DRY-RUN: $*"
  else
    echo "+ $*"
    "$@"
  fi
}

mkdir -p "$dl/ecma-376" "$dl/ms-openspecs"

for entry in "${specs[@]}"; do
  target="$dl/${entry%%|*}"
  url="${entry##*|}"
  if [ -f "$target" ]; then
    echo "-- $target :: already present, skipping"
    continue
  fi
  run mkdir -p "$(dirname "$target")"
  run curl -fsSL -o "$target" "$url"
done

for url in "${ms_specs[@]}"; do
  name="$(basename "${url%/}")"
  target="$dl/ms-openspecs/$name.html"
  if [ -f "$target" ]; then
    echo "-- $target :: already present, skipping"
    continue
  fi
  run curl -fsSL -o "$target" "$url"
done

echo
echo "Done. Downloads live under: $dl"
if [ "$dry_run" -eq 0 ]; then
  echo "Next step: unzip $dl/ecma-376/Part*.zip and copy XSDs to references/specs/xsd/"
fi
