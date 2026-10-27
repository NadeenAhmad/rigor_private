# RIGOR @ ISWC 2026: Phase 1 outline

Principle: **improve the existing 23-slide deck; don't replace it.** The running example,
the six-step build, the six cumulative architecture slides and the overall layout all
stay. Changes are targeted: the results section, the BURR explanation, future work,
the code excerpts, the colors and fonts, and constant backgrounds in the method section.

Time: 15 min speaking, Q&A separate. Spoken target **13:30** (leaves 1:30 for nerves,
clicks and pauses). Current notes: 1,517 words ending at 14:18. The plan keeps the
word count at about 1,420 (the current pace of ~105 words/min, pauses included) and pays for the new BURR explanation and the
future-work example by trimming the method slides by a few seconds each.

---

## 1. Thesis and takeaways

**Thesis (one sentence):** RIGOR anchors LLM ontology generation in the database itself
(order, deterministic seed, validation) and grounds it in retrieved knowledge, so the
ontology gains real domain meaning without losing the link back to the data.

**Three things the audience must remember:**
1. **Schema first, LLM second.** A deterministic seed plus FK order means the LLM
   enriches; it doesn't invent.
2. **The ontology grows and checks itself.** Retrieval from the growing core plus a
   judge before every merge.
3. **It works.** Experts ranked it first in 70% of rankings, it has the best F1 on the
   BURR benchmark, and it is the only method that produces hierarchy and provenance
   while keeping the schema.

---

## 2. Recommendation: how to structure the results

The OOPS! pitfall count is the weakest evidence for a best-paper audience. It mostly
counts missing annotations, which RIGOR is told to add, so it reads as "we passed our
own checklist". It shouldn't get its own slide.

RIGOR's real strengths line up with the question the deck already asks on slide 4:
*"Can we preserve schema fidelity **and** add domain semantics?"* So the results
should answer that question in three steps, ending on the strongest evidence, people:

| # | Slide (replaces) | Question it answers | Evidence | Why it convinces |
|---|---|---|---|---|
| R1 | **"Only RIGOR adds meaning without losing structure"** (replaces slide 20, pitfalls) | Does it add semantics *and* keep the schema? | One compact matrix, methods × what the ontology contains (Tables 8–9), plus one line on pitfalls | Shows things that *only* RIGOR produces, measured, not prompted |
| R2 | **"RIGOR keeps the link back to the data: best on BURR"** (slide 21, reworked) | Is the fidelity real, by an external yardstick? | BURR, explained first and then scored (Table 13) | An outside benchmark with an expert gold standard; reverses the benchmark authors' own finding that LLMs trail rule-based systems |
| R3 | **"Experts prefer RIGOR"** (slide 19, moved after BURR) | Is it actually better? | 28/40 first places (70%), Kendall's W; LLM judge 58/58 | Human judgment is the climax; the existing design stays |

**R1 matrix content (all numbers measured in the paper):**

| | Direct Mapping | LLM baseline | Non-Iterative | **RIGOR** |
|---|---|---|---|---|
| Columns kept, final coverage (Table 9) | 0.78–0.99 | 0.31–1.00 | **0.06**–0.78 | **0.77–1.00** |
| Labels & comments (Table 8, runs with any) | 0/3 | 0/9 | 9/9 | **9/9** |
| Class hierarchy, SubClassOf (Table 8) | 0/3 | 0/9 | 0/9 | **9/9** |
| Provenance to source column (Table 8) | 0/3 | 0/9 | 0/9 | **9/9** |

Pitfall line under the matrix (raw Table 7 counts, honest wording): *"Modeling pitfalls
(OOPS!): fewer than every other method in 9/9 like-for-like comparisons (same database,
same LLM); zero on Chinook with all three LLMs."* The raw table goes to a backup slide.

Notes on the matrix:
- The coverage row is the punchline: single-pass RAG (Non-Iterative) keeps as little
  as 6% of columns, while RIGOR keeps 77–100% *and* adds hierarchy and provenance.
- Disjointness (1/9 runs) is left off the main slide; it belongs to future work and Q&A.
- Use Table 8/9 measurements, **not** paper Table 3. Table 3 is a design checklist,
  and it wrongly marks Direct Mapping ✓ for provenance (Table 8 shows 0).
- Direct Mapping's 0.78 on Chinook reflects noise in the embedding-based coverage
  metric (it is complete by construction). Say "coverage proxy" in the caption.

Alternative order (if you prefer the headline first): R3 → R1 → R2. I recommend
R1 → R2 → R3, because it answers the slide-4 question in order and ends on people.

---

## 3. Slide list with time budget

Legend: **Keep** = content as-is (restyle only) · **Improve** = targeted edits ·
**Replace** = new content in the existing layout style.
Seconds include pauses (≈ 105 words/min effective, the current deck's own pace).

| # | Title (assertion) | Status | Message | Visual change | Paper source | Time | Words | Ends |
|---|---|---|---|---|---|---|---|---|
| 1 | Retrieval-Augmented Generation of Ontologies from Relational Databases | Keep | Schemas store data, not meaning | Restyle only | §1 | 0:25 | 45 | 0:25 |
| 2 | A database can tell us how data is stored. Can it tell us what the data means? | Keep | BM_DEP, ECOG, FK: meaning is implicit | Restyle | §1, §3.2 | 0:45 | 85 | 1:10 |
| 3 | Not just a cleaner schema. A usable semantic model. | Improve | The target: a linked, labelled, traceable ontology | Right panel becomes a real Manchester Syntax code panel with plain-English glosses (Phase 3) | §3 | 0:45 | 80 | 1:55 |
| 4 | Two partial solutions. Neither solves the whole problem. | Keep | Rules keep structure; LLMs suggest meaning; we need both | Restyle; this question is answered by R1–R3 | §1–2 | 0:40 | 70 | 2:35 |
| 5 | Build the ontology one table at a time. | Improve | Six moves, one table | Optional one-line teaser at the bottom: "Result: first in 70% of expert rankings · best F1 on BURR" | §3.1 | 0:25 | 45 | 3:00 |
| 6 | To understand this table, build its parent first. | Keep | FK order makes parents available | Restyle; same background as 5–17 | §3.1–3.2 | 0:35 | 60 | 3:35 |
| 7 | Guarantee schema coverage before LLM enrichment. | Improve | Deterministic seed covers every cleaned column | Output panel becomes real code (class, data properties, FK object property, existential restriction) + glosses | §3.2 | 0:40 | 75 | 4:15 |
| 8 | The draft is structurally faithful, but what does it mean? | Keep | Three retrieval sources | Restyle | §3.3 | 0:45 | 75 | 5:00 |
| 9 | The LLM enriches a schema-grounded draft. | Improve | The delta: labels, value meanings, provenance | Before/after as code; added lines highlighted; keep the evidence tags (SCHEMA/DOCS) | §3.4 | 0:45 | 80 | 5:45 |
| 10 | Why validate? Some errors are plausible, and still wrong. | Improve | Judge for semantic errors, rules for structural ones | Wrong lines in code with red marks; add a visible "illustrative" tag (currently only in the notes) | §3.5 | 0:50 | 85 | 6:35 |
| 11 | The ontology we built now helps build the next part. | Keep | Union merge; the core feeds the next table | Restyle | §3.6 | 0:35 | 60 | 7:10 |
| 12–17 | Reassembling the architecture (1/6 … 6/6) | Keep | Reconnect each step to paper Fig. 1 | Same background as 5–11 (no dark switch) | Fig. 1 | 1:20 total (12,12,14,12,12,18 s) | 130 | 8:30 |
| 18 | Does the method work beyond our running example? | Improve | 3 DBs × 3 LLMs × 4 methods; 30 ontologies | "independent LLM judge" → "independent evaluation judge (GPT-5.4)" | §4.1–4.6 | 0:35 | 65 | 9:05 |
| 19 → R1 | Only RIGOR adds meaning without losing structure | **Replace** (was pitfalls, slide 20) | Matrix from §2 + pitfall line | Matrix in the existing table style, RIGOR column highlighted | Tables 7–9 | 0:55 | 100 | 10:00 |
| 20 → R2 | RIGOR keeps the link back to the data: best on BURR | **Improve** (was slide 21) | What BURR measures, then the scores | Left: mini worked example (table → class, FK → relation, column → attribute, compared with the gold mapping). Right: existing F1 table | Table 13, §4.6 | 1:05 | 115 | 11:05 |
| 21 → R3 | Experts prefer RIGOR | Keep (moved, was slide 19) | 70% first place; LLM judge 58/58 | Restyle only | Tables 1, 11, 12 | 0:50 | 85 | 11:55 |
| 22 | Next: let the data speak, from schemas to instances | **Replace** (was limitations) | Instances + documentation → candidate knowledge (cisplatin / bone marrow depression), reviewed by experts | 3-stage flow + code panel (Phase 5); bottom line: tool feedback in the loop | §5 | 1:00 | 105 | 12:55 |
| 23 | Thank you. Questions? | Improve | Restate the thesis + 3 takeaways | Keep QR codes; the three takeaways replace the tagline | — | 0:35 | 60 | 13:30 |
| B1–B6 | Backup (untimed) | New | Raw OOPS! table; per-dimension scores; ablation; cost; hard questions | — | Tables 7, 10–12, 14, 15 | — | — | — |

**Where the time comes from:** the new BURR explanation (+10 s) and the future-work
example (+12 s) are covered by trimming slides 9, 10 and 11 (−5 to −15 s each),
merging the pitfall commentary into the R1 line, and a tighter close.

---

## 4. Claims audit of the current deck

| Slide | Claim | Verdict | Fix |
|---|---|---|---|
| 4 | "…without expert intervention during generation?" | ok | Precise wording; keep (better than the paper's "no human oversight") |
| 6 | FK-guided breadth-first order | ok | Q&A risk: self-referencing FKs (Chinook `Employee.ReportsTo`) → backup answer |
| 7 | "Every cleaned-schema column is represented in the deterministic seed" | ok | Keep the note's point that this applies to the seed, not the final ontology |
| 8 | External ontologies "suggest established concepts" | ok as the paper's method | Q&A risk: the Gen-LLM prompt in the appendix has no external-ontology slot, and no alignment counts are reported → backup answer |
| 9 | ECOG comment, provenance to the source column | ok | Already marked illustrative on the slide |
| 10 | Hypothetical errors | reword | The "hypothetical" label is only in the notes; add a small visible "illustrative" tag |
| 10 | "up to 2 retries", fallback | ok | — |
| 11 | Merge = axiom-set union, reuse of IRIs | ok | Q&A risk: union merge + shared properties → P19 (multiple domains) |
| 18 | "independent LLM judge" | reword | The in-loop judge is the *same* model as the generator; say "independent evaluation judge (GPT-5.4)" |
| 18 | 30 ontologies; 3 × 3 × 3 + 3 | ok | — |
| 19 | 28/40 (70%); DM 5, Baseline 5, Non-It 2; W 0.39–0.72; 58/58; 290 CQs | ok | All match Tables 1 and 12 (80 + 155 + 55 = 290) |
| 19 notes | Experts valued domain enrichment, clarity, completeness | ok | Matches the largest gaps in Table 11 |
| 20 | "RIGOR has the fewest modeling pitfalls in all nine comparisons" | reword → moved | True only like-for-like (same database + LLM). In raw counts Non-It/DeepSeek (18) < RIGOR/Claude (20) on Liver; eICU/DeepSeek is 26 vs 27. Replaced by the R1 line |
| 20 | P19 in 3/9 RIGOR runs | ok | Liver Cl 5, Liver Mi 7, eICU DS 13 |
| 21 | F1 .73/.67/.52 etc.; 9 tables / 46 columns / 11 FKs; D2RQ export | ok | Add *what* is evaluated (R2) |
| 21 | OntoGenix "LLM-based", RDB2Onto "rule-based" | ok | — |
| 22 | Coverage 0.77–1.00; P19 3/9; disjointness rare | ok | Coverage moves into R1; P19 and disjointness move into future work / backup |
| 23 | "ISWC 2026 · Bari" | check | Confirm the venue city from the conference website before the talk |

### Deck-wide issues found while auditing

1. **Fonts are a venue risk.** The deck uses IBM Plex Sans (1,695 runs) and IBM Plex
   Mono (554), which are not installed on typical venue PCs, plus Aptos (44 runs),
   which is missing from older Office versions. Without embedding they silently fall
   back to other fonts and text reflows. Phase 2 swaps them for fonts that ship with
   Office, or embeds them, and Phase 7 produces a PDF fallback.
2. **Backgrounds alternate inside the method section.** Slide 5 is dark, 6–8 light,
   9 dark, 10–11 light, 12–17 dark. Phase 2 makes slides 5–17 one continuous background.
3. **Colors are hard-coded**, not theme-based (about 25 distinct hex values, mainly the
   navy 111A28 / blue 84ADF7 / orange F9AB5E family). A new palette is therefore a
   controlled color-map replacement over the existing XML, with no layout changes.
   That is how Phase 2 restyles without "destroying" the deck.

---

## 5. Decisions recorded from the author

- Speaking slot: 15 min, Q&A separate → spoken target 13:30.
- Keep all six cumulative architecture slides (12–17).
- Code excerpts: real Manchester Syntax plus a plain-English gloss beside the key lines.
- Pitfalls: no dedicated slide; results restructured as R1 → R2 → R3 (section 2).
- Improve the existing deck; no rebuild from scratch.

## 6. Open questions for the author

1. Results order: R1 → R2 → R3 (recommended), or experts first?
2. Slide 5 teaser line: yes or no?
3. Venue city on slides 1 and 23: please confirm.
