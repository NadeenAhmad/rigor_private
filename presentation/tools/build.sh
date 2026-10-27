#!/usr/bin/env bash
# Rebuild presentation/RIGOR_ISWC2026.pptx from the author's original deck.
#   Phase 2: restyle.py        (palette, fonts, backgrounds, tracker, text edits)
#   Phase 3: code_excerpts.py  (verified OWL Manchester Syntax excerpts)
set -euo pipefail
cd "$(dirname "$0")/../.."
tmp=$(mktemp -d)
python3 presentation/tools/restyle.py RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx "$tmp/p2.pptx"
python3 presentation/tools/code_excerpts.py apply "$tmp/p2.pptx" presentation/RIGOR_ISWC2026.pptx
rm -rf "$tmp"
