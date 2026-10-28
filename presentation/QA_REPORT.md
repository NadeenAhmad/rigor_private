# Phase 7: final QA report

Deck: `presentation/RIGOR_ISWC2026.pptx` (23 main slides + backups B1–B6).
Fallback: `presentation/RIGOR_ISWC2026.pdf` (29 pages, all fonts embedded).
Rebuild everything from the original deck with `presentation/tools/build.sh`.

## 1. Claims audit (every on-slide number against the paper)

| Slide | Claim on slide | Paper source | Result |
|---|---|---|---|
| 5 | First in 70% of expert rankings; highest BURR F1 among compared ontology-learning systems | Table 1; Table 13 | ok |
| 8 | Top-3 per source; ICD-10, Disease, Cell, Nutrition ontologies | §3.3, §4.3 | ok |
| 10 | Up to 2 retries, then fallback; examples: duplicate concept (Judge criterion 9) and a second domain on a reused property (P19, still in 3/9 runs) | §3.5; Table 7 | ok |
| 18 | 16/350, 31/559, 11/67 tables/columns; 30 ontologies; GPT-5.4; 10 experts | §4.1, §4.4, §4.7 | ok |
| 19 | Column coverage DM 0.78–0.99, Baseline 0.31–1.00, Non-It 0.06–0.78, RIGOR 0.77–1.00 | Table 9 (min/max of 3 DBs × 3 LLMs) | ok |
| 19 | Labels & comments: 0/3, 0/9, 9/9, 9/9 | Table 8, AnnAx > 0 | ok |
| 19 | Hierarchy + provenance: 0/3, 0/9, 0/9, 9/9 | Table 8, SubCls > 0 and Prov > 0 | ok |
| 19 | Fewest pitfalls in 9/9 like-for-like runs; zero on Chinook | Table 7 (checked against every other method, same DB and LLM) | ok (eICU/DeepSeek is 26 vs 27) |
| 20 | F1 .53/.11/.39, .67/.62/.44, .58/.35/.44, RIGOR .73/.67/.52 | Table 13 | ok |
| 20 | 9 tables · 46 columns · 11 FKs; matching rule; D2RQ export | §4.6; BURR Table 2 and §4 Eq. 3 | ok |
| 21 | 28/40 (70%); DM 5, Baseline 5, Non-It 2; W 0.39–0.72; 58/58 tables; 290 CQs | Tables 1, 12; §4.6 | ok |
| 22 | Directions: instance signals, tool feedback, symbolic reasoning | §5 | ok (rows are labelled synthetic) |
| B1 | Raw OOPS! totals | Table 7 | ok |
| B2 | 56 expert scores + averages | Table 11 | ok |
| B3 | Ablation P/R/F1 | Table 14 | ok |
| B4 | Cost rows; totals 552 / 4,937,000 / $33.28 / 616.1 min | Table 15 rows (printed totals differ; noted on slide) | ok |
| B5–B6 | 70%, 6%, 28/40, P19 3/9, 16/16, F1 0.059 vs 0.033, Gemini .65/.68/.56 | Tables 1, 7, 9, 14; BURR Table 3 | fixed: "16/16" now says it was the LLM judge |

All OWL on slides 3, 7, 9, 10 and 22 parses with the OWL API (`presentation/excerpts/`).

## 2. Visual QA

All 29 slides rendered at 150 dpi (Consolas widths emulated): no overflow, overlap or
off-slide content. Titles sit at the same position within each slide type, and slides 5–17
share one white background. Title and closing slides changed to white at the author's request
(option B); emerald stays the accent.

## 3. Accessibility

- Contrast: every text colour against the fill it sits on is at least 4.5:1 (34 pairs; check
  with `presentation/tools/qa_contrast.py`). Fixed in this phase: upcoming tracker steps and
  grey "0 / 9" counts (8A949B → 6B747B), the mint arrow label (→ 177A55), and the gold and pale
  text on the former green slides.
- Alt text: title motif described; QR codes now give their decoded targets
  (https://arxiv.org/abs/2506.01232 and https://github.com/NadeenAhmad/RIGOR, both verified).
- Small text: titles, key numbers and code glosses are 12 pt or larger. Secondary labels
  below 12 pt remain where enlarging would overflow: slide 9's compact "before" panel and tags,
  slide 20's example chips (10.5 pt), slide 22's code and table header (9.5–11 pt), and the
  backup tables.

## 4. Fonts and the venue PC

- The deck uses only Cambria (titles), Calibri (body) and Consolas (labels, code). All three
  ship with Windows and with Microsoft Office for Mac. IBM Plex and Aptos are gone, including
  the master's unused bullet-font settings.
- If the venue PC has no Office fonts (e.g. LibreOffice or a kiosk viewer), present the PDF.
  It was rendered with metric-compatible substitutes (Caladea for Cambria, Carlito for Calibri)
  and a width-corrected monospace for Consolas, so the layout matches. Glyph shapes differ
  slightly from the real fonts. A few special characters (ΔOᵣ, Oₜ) use a fallback glyph font.
- Optional: in PowerPoint for Windows, File → Options → Save → "Embed fonts in the file"
  makes the .pptx self-contained.

## 5. Open items for the author

1. Check the self-referencing-FK answer (B6, Chinook `Employee.ReportsTo`) against the code.
2. Read the script aloud once with a timer (planned end 12:24; target 13:30).
3. Do not repeat the paper's claim that BURR found LLMs trail rule-based systems (§2.1, §4.7);
   BURR's results say otherwise. Slide 20's notes have the safe wording.
4. Slide numbers were not added: the recap slides already show "1 / 6" … "6 / 6" bottom right
   and the closing slide shows the venue there. Ask if you want numbers on the other slides.

## 6. Changes after Phase 7 (author feedback)

- All em dashes removed from slides and notes (six titles and labels, rephrased).
- Slide 10 now shows errors a capable LLM still makes: a duplicate concept (`Patient` next to
  the core's `PatientData`) and a reused property with a second domain (P19). Both are verified
  with the OWL API, including the domain-axiom count before and after the correction.
- Slide 19 redrawn as an open table: hairline rules, a RIGOR band, and coverage as range bars
  on a 0 to 1 axis.
- Cover redesigned (`presentation/tools/cover.py`): RIGOR wordmark, authors on three balanced
  lines with true superscripts (presenter in bold), one affiliation per line in full, a links
  footer, and an emerald panel on the right third only, with the motif and the venue.
- Cover panel reworked after feedback (the motif lacked contrast on emerald): a near-white
  graph-paper panel with a vector drawing of the running example (admission table → RIGOR →
  Admission/PatientData with label, comment and provenance chips; BM_DEP and its label in gold).
- Cover figure restored at the author's request: the original table-to-ontology figure,
  recoloured for the light panel (charcoal table and nodes, emerald edges, amber mapped row and
  node), with a RIGOR label above the mapping and alt text describing it.
- Slide 3 now shows the full target example (13 lines) with a five-click reveal by knowledge source:
  schema, documentation, Disease Ontology (obo:DOID_4961, skos:relatedMatch), LLM knowledge, then
  "What this enables". Every reveal state is verified with the OWL API. Standalone copy:
  `presentation/RIGOR_slide3_target.pptx`; frame-by-frame preview: `presentation/slide3_reveal_steps.jpg`.
- Slide 4 now cites the approaches it calls limited (all from paper Section 2), author-year, under
  each card (`presentation/tools/references.py`). BURR is deliberately not cited there. Standalone
  copy: `presentation/RIGOR_slide4_gap.pptx`.
- Slide 9 AFTER panel now shows the same 13-line fragment as slide 3, one fixed text box per
  line (no overflow in PowerPoint), with an evidence tag beside each enriched line (SCHEMA, DOCS,
  ONTOLOGY, LLM; slide-3 colours). Header and labels follow the author's edits; the old caption is
  gone; provenance uses `admission:`. Verified with the OWL API; contrast 0 pairs below 4.5:1.
  Standalone copy: `presentation/RIGOR_slide9_generate.pptx`; preview: `presentation/slide9_preview.jpg`.
