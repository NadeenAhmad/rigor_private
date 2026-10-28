#!/usr/bin/env bash
# Rebuild presentation/RIGOR_ISWC2026.pptx from the author's original deck.
#   Phase 2: restyle.py        (palette, fonts, backgrounds, tracker, text edits)
#   Phase 3: code_excerpts.py  (verified OWL Manchester Syntax excerpts)
#   Phase 4: results.py        (results R1 -> R2 -> R3, backup B1)
#   Phase 5: future_work.py    (slide 22: from schemas to instances)
#   Phase 6: speech.py         (speaker notes, closing takeaways; also writes SPEECH.md, REHEARSAL_CARD.md)
#            backups.py        (backup slides B2-B6)
#   Later:   cover.py          (cover slide: wordmark, authors, graph-paper panel with the figure)
#            references.py     (author-year references under slide claims)
set -euo pipefail
cd "$(dirname "$0")/../.."
tmp=$(mktemp -d)
python3 presentation/tools/restyle.py RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx "$tmp/p2.pptx"
python3 presentation/tools/code_excerpts.py apply "$tmp/p2.pptx" "$tmp/p3.pptx"
python3 presentation/tools/results.py "$tmp/p3.pptx" "$tmp/p4.pptx"
python3 presentation/tools/future_work.py "$tmp/p4.pptx" "$tmp/p5.pptx"
python3 presentation/tools/speech.py apply "$tmp/p5.pptx" "$tmp/p6a.pptx"
python3 presentation/tools/cover.py "$tmp/p6a.pptx" "$tmp/p6b.pptx"
python3 presentation/tools/references.py "$tmp/p6b.pptx" "$tmp/p6.pptx"
python3 presentation/tools/backups.py "$tmp/p6.pptx" presentation/RIGOR_ISWC2026.pptx
python3 presentation/tools/speech.py docs
rm -rf "$tmp"
