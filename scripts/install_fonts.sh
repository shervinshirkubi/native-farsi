#!/usr/bin/env bash
# install_fonts.sh — make bundled Persian fonts visible to LibreOffice/Chrome/weasyprint.
# Run BEFORE any docx→PDF or pptx→PDF conversion, or Persian silently renders in DejaVu.
set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$SKILL_DIR/assets/fonts"

count=$(find "$SRC" -maxdepth 2 \( -iname '*.ttf' -o -iname '*.otf' \) 2>/dev/null | wc -l)
if [ "$count" -eq 0 ]; then
  echo "ERROR: no font files in $SRC" >&2
  echo "Bundled fonts are missing. Ask the user for the TTFs or run scripts/download_fonts.py (needs GitHub access)." >&2
  exit 1
fi

if [ "$(uname)" = "Darwin" ]; then
  # macOS has no fontconfig by default; Font Book/CoreText reads ~/Library/Fonts directly.
  DEST="$HOME/Library/Fonts"
  mkdir -p "$DEST"
  find "$SRC" -maxdepth 2 \( -iname '*.ttf' -o -iname '*.otf' \) -exec cp -f {} "$DEST/" \;
  echo "installed $count font file(s) to $DEST (macOS, no fc-cache needed):"
  ls "$DEST" | grep -iE 'vazirmatn|lalezar' | sed 's/^/  /'

  # LibreOffice.app on macOS is a self-contained bundle: it does NOT read
  # ~/Library/Fonts and ignores Homebrew's system fontconfig. It only sees
  # fonts placed in its own bundled font directory. Without this step,
  # every docx/pptx→PDF conversion silently falls back to DejaVu — found by
  # end-to-end testing, not documented upstream.
  LO_FONTS="/Applications/LibreOffice.app/Contents/Resources/fonts/truetype"
  if [ -d "$LO_FONTS" ] && [ -w "$LO_FONTS" ]; then
    find "$SRC" -maxdepth 2 \( -iname '*.ttf' -o -iname '*.otf' \) -exec cp -f {} "$LO_FONTS/" \;
    echo "also installed to LibreOffice's bundled fonts: $LO_FONTS"
  elif [ -d "/Applications/LibreOffice.app" ]; then
    echo "note: LibreOffice.app found but $LO_FONTS not writable — PDF conversion" >&2
    echo "      may still fall back to DejaVu. Fix permissions or copy fonts manually." >&2
  fi

  if command -v fc-list >/dev/null 2>&1; then
    fc-cache -f "$DEST" >/dev/null 2>&1 || true
  else
    echo "note: fontconfig not installed (fine for native macOS apps — Word, Preview)." >&2
    echo "      Install with 'brew install fontconfig' only if some other RTL tool" >&2
    echo "      needs it directly; LibreOffice conversion no longer depends on it." >&2
  fi
else
  # Linux: fontconfig is the standard path.
  DEST="$HOME/.fonts"
  mkdir -p "$DEST"
  find "$SRC" -maxdepth 2 \( -iname '*.ttf' -o -iname '*.otf' \) -exec cp -f {} "$DEST/" \;
  if ! command -v fc-cache >/dev/null 2>&1; then
    echo "ERROR: fc-cache not found. Install fontconfig (e.g. apt-get install fontconfig)." >&2
    exit 1
  fi
  fc-cache -f "$DEST" >/dev/null 2>&1 || fc-cache -f >/dev/null 2>&1
  echo "installed $count font file(s) to $DEST:"
  fc-list :lang=fa family 2>/dev/null | sort -u | sed 's/^/  /'
fi
