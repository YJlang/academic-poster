#!/usr/bin/env bash
# macOS only: strip the com.apple.quarantine attribute so PowerPoint opens the
# generated .pptx in EDIT mode instead of read-only "Protected View".
# Usage: ./unlock_macos.sh poster.pptx [more.pptx ...]
set -e
if [ "$#" -eq 0 ]; then echo "usage: $0 file.pptx [...]"; exit 1; fi
for f in "$@"; do
  xattr -c "$f" && echo "unlocked: $f"
done
