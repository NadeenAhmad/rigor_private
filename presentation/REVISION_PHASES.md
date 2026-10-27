# RIGOR ISWC 2026 talk: revision plan in phases

Each phase below is a self-contained prompt. Run them **in order**, one per message
(paste the prompt as-is). Every phase reads the **Shared context** section first,
commits its output to branch `claude/bold-cori-h54gh3`, and ends by reporting what
changed plus anything that needs the author's decision.

Source files in the repo:

- Paper: `2506.01232v2 (1).pdf` (33 pages; the source of truth for every number)
- Current deck: `RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx` (23 slides, with speaker notes)
- From Phase 2 on: `presentation/build_deck.js` (generator) → `presentation/RIGOR_ISWC2026.pptx`

---

## Shared context (read before every phase)

**Goal.** A 15-minute ISWC 2026 research-track talk strong enough to compete for best
paper. Spoken part ≤ 14:00 (aim for 13:30, leaving room for slower delivery). Q&A is separate.

**What the deck does well (keep it).** The single running example (`admission` →
`PatientData`, `BM_DEP`, `ecog_stage`) carried through all six steps. The honesty about
what is illustrative versus measured. The six-step structure: ORDER, SEED, RETRIEVE,
GENERATE, VALIDATE, MERGE.

**What the author wants changed.**
1. Slide 20 (OOPS! pitfalls) is weak: a 12-number table nobody can read in 50 seconds.
2. Slide 21 (BURR) never tells the audience *what is being evaluated*.
3. The limitations slide should become a **future-work** slide built around an
   instance-level example (bone marrow depression + cisplatin; see Phase 5 for the
   semantic-web framing, which is mandatory).
4. Ontology excerpts look "childish" (pseudo-notation such as `patient_id → PatientData`).
   They must be **real OWL 2 Manchester Syntax** in a monospace code font, styled like a
   code editor.
5. The look (dark navy + orange, dark/light alternation) duplicates another talk's
   template. Choose a new, professional, distinctive identity. Dark mode is not required.
6. **The background must not change within the method walkthrough** (overview → ORDER →
   … → MERGE → architecture). One continuous visual space.

**Claims hygiene.** Every number on a slide must trace to a paper table or section,
cited in small text on the slide (e.g. "Paper Table 7"). Avoid claims the paper's own
tables contradict, in particular:
- "fewest pitfalls than all other methods": on Liver Cancer, Non-Iterative/DeepSeek (18)
  is below RIGOR/Claude (20) and RIGOR/Mistral (23). Only *matched* comparisons (same
  database + same LLM) are 9/9 wins. eICU/DeepSeek is 26 vs 27, essentially a tie in
  raw counts.
- "independent Judge-LLM": inside the pipeline the judge is the *same* model as the
  generator. Only the *evaluation* judge (GPT-5.4) is independent.
- "no human oversight": say "no human in the loop during generation".
- Coverage: the direct-mapping seed is complete by construction; final column coverage
  is 0.77–1.00.

**Facts used repeatedly (verified against the paper).**
- Databases: Liver Cancer 16 tables / 350 columns / 15 FKs (private, so no training
  leakage). eICU-CRD 31 / 559 / 30. Chinook 11 / 67 / 11.
- Generators: Claude Opus 4.6, Mistral-Small-24B, DeepSeek-V3. Evaluation judge:
  GPT-5.4. 30 ontologies = 3 × 3 × 3 + 3 direct mappings.
- Experts: 10 experts, 4 medical tables, 40 rankings. RIGOR ranked 1st 28/40 (70%),
  DM 5, Baseline 5, Non-Iterative 2. Kendall's W 0.39–0.72, all significant.
- LLM judge: RIGOR ranked 1st on 16/16, 31/31, 11/11 tables (58/58), with 5 CQs per
  table (290 CQs). Claude outputs only.
- OOPS! totals (Table 7), in the order Claude / Mistral / DeepSeek:
  - Liver: DM 731; Baseline 422/397/403; Non-It 29/61/18; RIGOR 20/23/3
  - eICU: DM 665; Baseline 434/440/421; Non-It 39/220/27; RIGOR 2/2/26
  - Chinook: DM 102; Baseline 95/74/81; Non-It 24/35/14; RIGOR 0/0/0
  - Critical P19 remains in 3/9 RIGOR runs (Liver Cl 5, Liver Mi 7, eICU DS 13).
- BURR ISWC scenario (Table 13), class / relation / attribute F1:
  W3C-Mapper/Ontop .53/.11/.39 · RDB2Onto .67/.62/.44 · OntoGenix .58/.35/.44 ·
  RIGOR .73/.67/.52. 9 tables, 46 columns, 11 FKs. RIGOR used Claude; its
  `prov:wasDerivedFrom` annotations were exported as a D2RQ mapping.

**Tooling rules.** Use the `anthropic-skills:pptx` skill. Build the deck from a
`pptxgenjs` generator script as a structured deck (theme, layouts, placeholders,
sections), so later phases edit the script, not XML. After every build, run
`validate.py`, render to images, and inspect every changed slide for overflow and
overlap. Keep speaker notes in `slide.addNotes()`.

---

## Phase 1: Storyline, slide list, 15-minute budget

```
Read presentation/REVISION_PHASES.md (Shared context) and the current deck's text and
notes (pptx skill: scripts/slide_text.py). Do not build slides yet.

Produce presentation/OUTLINE.md:
1. One-sentence thesis of the talk and the 3 things the audience must remember.
2. The final slide list with, per slide: title (an assertion, not a topic), the one
   message, the visual, the source in the paper, and a time budget. The total spoken
   time must be ≤ 13:30.
3. Default structure (change only with a stated reason):
   Title → Motivation → Target ontology → Gap → "One table at a time" overview →
   Steps 1–6 → ONE full-architecture slide (replaces the six cumulative recap slides
   12–17, saving about 1:15 for results and future work) → Setup → Experts + LLM judge →
   Pitfalls → BURR → Future work: from schemas to instances → Take-home + Q&A
   (merge conclusion and thank-you into one closing slide). Then backup slides (Phase 6).
4. Put a one-line result teaser early (e.g. on the overview slide: "ranked 1st by
   experts in 70% of rankings · best F1 on BURR") so the audience knows the payoff.
5. Audit every claim in the current deck against the paper using the "Claims hygiene"
   list. Write a table: slide, claim, verdict (ok / reword / remove), and fix.
Commit OUTLINE.md and report the slide list with timings.
```

## Phase 2: New visual identity and deck rebuild

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/OUTLINE.md.
Rebuild the deck from scratch as presentation/build_deck.js (pptxgenjs, LAYOUT_WIDE,
structured deck), following OUTLINE.md's slide list. Port the existing content and notes
as they are; later phases rewrite specific slides.

Visual identity, "Lab notebook / editorial light" (deliberately unlike the dark
navy + orange template):
- Background: white FFFFFF on every content slide. Panels and cards: F3F5F4. Hairline
  borders: D9DEDB. No dark/light alternation. The ENTIRE method section (overview,
  steps 1–6, architecture) shares one background and the same frame.
- Ink: graphite 1F2328 for text, slate 5B6770 for secondary text.
- One dominant accent: deep emerald 0B6E4F ("ontology green") for RIGOR, highlights and
  the active step.
- One sharp signal color, used rarely: vermilion D64C2B, only for errors, warnings and
  the "what goes wrong" examples.
- Title and closing slides may use a full emerald (0B6E4F) background with white text.
  These are the only colored backgrounds.
- Typography: headings in Georgia (editorial serif; leave about 10% width slack), body
  in Calibri, code in Consolas (Courier New fallback). Title 36–40 pt, body ≥ 16 pt,
  code ≥ 14 pt, captions 11–12 pt.
- Motif, used everywhere: a growing graph. The step tracker on steps 1–6 is a row of six
  nodes joined by edges; completed steps are filled emerald, the current step has an
  emerald ring, future steps are hollow grey. It sits in the same position on every
  method slide and replaces the current text-box tracker. Reuse the node glyph as the
  section marker. No accent bars, no underline rules, no edge stripes.
- Code panels (styling only here; content comes in Phase 3): background F6F8FA, 1 px
  border D0D7DE, 8 px rounded corners, a small top-left tab label (e.g.
  "admission.omn · Manchester Syntax"), Consolas 14–16 pt, line height 1.15. Syntax
  colors: keywords (Class:, ObjectProperty:, SubClassOf:, Domain:, Range:,
  Annotations:, some) in emerald bold; IRIs / prefixed names in graphite; string
  literals and language tags in A3401E; datatypes (xsd:*) in 6F42C1; prov:* annotations
  in slate.
- Diagrams: hand-built shapes in the palette (tables as cards with header rows,
  ontology classes as rounded nodes, properties as labelled arrows).
- Reuse the two QR code images from the original deck (Picture 1051/1053 on slide 23).

Then build, validate (validate.py), render every slide, and fix overflow, overlap and
misalignment. Check that titles sit at the same position on every content slide and
that the method section's background never changes. Commit build_deck.js and the
.pptx, and send me a contact sheet of all slides.
```

## Phase 3: Ontology excerpts as real OWL code

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/build_deck.js.
Replace every pseudo-notation ontology excerpt (target slide, SEED, GENERATE, VALIDATE,
MERGE, closing slide) with real OWL 2 Manchester Syntax in the Phase 2 code-panel style.

Rules:
- One consistent namespace for the whole talk: Prefix: : <http://example.org/rigor#>,
  plus rdfs:, xsd:, prov:, owl:. Same IRIs on every slide (:Admission, :PatientData,
  :hasPatientData, :bmDep, :ecogStage).
- Correct keywords with colons: Class:, ObjectProperty:, DataProperty:, Domain:,
  Range:, SubClassOf:, Annotations:. Language-tagged labels ("…"@en). Existential
  restriction for a NOT NULL FK: SubClassOf: :hasPatientData some :PatientData.
- SEED slide shows DMᵣ exactly as the deterministic mapping produces it (class, data
  properties with xsd ranges, FK as object property, existential restriction), with no
  labels or comments. GENERATE slide shows the same fragment after enrichment;
  highlight the added lines (emerald gutter mark or a faint emerald line background),
  so the delta is visible at a glance. Keep each panel ≤ 12 lines, using
  "…" (# other columns omitted) instead of cramming.
- ECOG comment uses the standard scale: 0 fully active · 1 restricted in strenuous
  activity · 2 ambulatory, no work · 3 limited self-care · 4 completely disabled · 5 dead.
- VALIDATE slide: show the wrong fragment (DataProperty: :patientId with Range:
  xsd:integer, plus the reversed ECOG comment) with vermilion marks, and the corrected
  lines beside it.
- Verify syntax mechanically. Assemble all excerpts into
  presentation/excerpts/rigor_example.omn (with prefixes and declarations) and parse it
  with the OWL API (ROBOT: `robot convert -i rigor_example.omn -o out.ttl`; fetch the
  ROBOT jar if the network allows). If parsing is impossible here, say so explicitly
  and hand-check against the W3C Manchester Syntax grammar instead. The slides must
  show exactly the verified text (excerpted lines only).
- Do NOT reproduce the paper's Appendix A.1 example format, which omits the colons
  after Domain/Range.
Rebuild, validate, render, fix, commit. Show me before/after images of each changed slide.
```

## Phase 4: Results that persuade (experts, pitfalls, BURR)

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/build_deck.js.
Rework the three results slides. Every number cites its paper table on the slide.

A) Experts + LLM judge: keep the 70% (28/40) hero number. Show the 4-method first-place
   counts as a native bar chart (DM 5, Baseline 5, Non-Iterative 2, RIGOR 28; RIGOR
   emerald, others grey), with "Kendall's W 0.39–0.72, p < .01" as a caption. Put the
   independent LLM judge as a second hero number, 58/58 tables. Say "Claude outputs;
   5 CQs per table".

B) Pitfalls (replaces slide 20). Message: "RIGOR removes almost all modeling pitfalls,
   in every matched comparison." Recommended design: a native clustered bar chart,
   one cluster per database, bars = method. Use pitfalls per 100 ontology entities
   (Table 7 ÷ (classes + object + data properties) from Table 8), averaged over the
   3 LLMs (DM has one run). Size-normalized counts pre-empt the "your ontologies are
   just different sizes" objection, and RIGOR then wins all 9 matched pairs clearly:
     per-100-entity values (Claude / Mistral / DeepSeek):
     Liver:   DM 198.6 · Baseline 105.2/104.2/103.3 · Non-It 14.3/122.0/17.5 · RIGOR 6.2/6.1/1.0
     eICU:    DM 200.3 · Baseline 112.4/115.5/108.2 · Non-It 40.6/123.6/17.9 · RIGOR 0.5/0.4/11.3
     Chinook: DM 217.0 · Baseline 146.2/145.1/147.3 · Non-It 37.5/145.8/28.6 · RIGOR 0/0/0
   Recompute these from the tables before use. Label the metric as derived ("computed
   from paper Tables 7–8"). Next to the chart, two callouts: "9/9 matched wins" and
   "Chinook: 0 pitfalls with all three LLMs". Footer (honest): "Most removed pitfalls
   are missing annotations (P08); critical P19 remains in 3/9 runs." Raw counts move to
   a backup slide. If the author prefers raw counts on the main slide, use a log-scale
   axis, never a 12-cell number grid.

C) BURR (replaces slide 21). The audience must understand what is measured before
   seeing numbers. Left half, "What BURR asks": a small worked diagram. An ISWC table
   (e.g. `person`), one FK, and one column map to gold ontology elements (class,
   object property, data property). The system's own mapping is compared with BURR's
   expert gold mapping; F1 per element type (Classes, Relations, Attributes). One line
   on why it matters: "Enrichment must not break the link back to the data; that link
   is what OBDA and KG population run on." One line on how RIGOR is scored: its
   prov:wasDerivedFrom provenance exported as a D2RQ mapping. Right half: the 4×3 F1
   results as a native grouped bar chart (or a clean table), RIGOR emerald. Caption:
   "BURR's authors found LLM methods trail rule-based ones on this task; here RIGOR
   leads both. One scenario, Claude." Before writing the left half, check how BURR
   actually matches elements (Laskowski et al., PACMMOD 2025, "BURR: A Benchmark for
   Ontology Learning from Relational Databases"; fetch it if the network allows).
   Describe only what you can confirm; otherwise keep the description at the level
   above.
Rebuild, validate, render, fix, commit, and show me the three slides.
```

## Phase 5: Future work, from schemas to instances (cisplatin example)

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/build_deck.js.
Replace the limitations slide with a future-work slide: "Next: let the data speak,
from schemas to instances".

Semantic-web framing (mandatory; already checked):
- Medically sound: myelosuppression is a documented, dose-related adverse reaction of
  cisplatin (drug label). Cisplatin is used in liver/biliary oncology (e.g. gemcitabine
  + cisplatin for biliary tract cancer; cisplatin in TACE/HAIC).
- NOT sound as stated "we can attribute that cisplatin caused it". Co-occurrence across
  patients is correlation. OWL reasoning is deductive and open-world: it cannot infer
  causation from instance statistics. A class axiom such as
  `Cisplatin SubClassOf causes some BoneMarrowDepression` would also be wrong, because
  it asserts something universal about every instance. Real confounders exist in this
  population: co-administered gemcitabine (also myelosuppressive), radiotherapy /
  radioembolization, and cirrhosis with hypersplenism (cytopenias without any drug).
- Correct framing: instance-level patterns + documentation → a CANDIDATE, provenance-
  annotated, reviewable association (an adverse-event association in the sense of
  OAE, the Ontology of Adverse Events, which defines adverse events as occurring
  after an intervention without asserting causality). It is proposed for expert
  validation, not asserted as a causal axiom. Temporality (BMD recorded after the
  first cisplatin cycle) is part of the evidence.

Slide content (left to right, as a 3-stage flow in the house style):
1. "Instances": 4–5 de-identified illustrative patient rows (patient, regimen,
   BM_DEP, date order), clearly labelled "synthetic illustration". Pattern: BM_DEP = 1
   recorded after cisplatin cycles. Note that instances also resolve encodings
   (is BM_DEP 0/1 or a CTCAE grade?), the safer and immediate win.
2. "+ Documentation": a drug-label snippet, "myelosuppression is a known, dose-related
   adverse reaction".
3. "→ Candidate knowledge": a Manchester Syntax code panel (verified as in Phase 3)
   asserting an individual of :CandidateAdverseDrugAssociation with :hasDrug
   :Cisplatin, :hasCondition :BoneMarrowDepression, an evidence status "candidate —
   expert review", and prov:wasDerivedFrom pointing to both the data pattern and the
   label. No `causes` axiom. If you add external IRIs (ChEBI for cisplatin, OAE), verify
   each ID before using it.
Bottom strip: the other two next steps in one line each: "OOPS!/reasoner feedback
inside the loop (agentic tool calls) → target the residual P19" and "aggregate-only
instance statistics, so clinical data never leaves the hospital".
Speaker note must say "may have contributed / candidate association", never "caused".
Rebuild, validate, render, fix, commit, and show me the slide.
```

## Phase 6: The speech, timing, and Q&A backups

```
Read presentation/REVISION_PHASES.md (Shared context), OUTLINE.md and build_deck.js.
1. Write the full spoken script, slide by slide, into presentation/SPEECH.md and into
   each slide's notes: first-person, conversational, short sentences, one idea per
   sentence, no reading of slide text. At 130 words/min the whole script is ≤ 1,750
   words (≈ 13:30). Per slide: [start–end] timestamp, word count, the script, one
   pointer cue, one-line takeaway. Mark 2–3 deliberate pauses (after the 70%, after
   "0 pitfalls on Chinook").
2. Opening hook in the first 20 s (the inherited clinical database with BM_DEP).
   Closing in 3 sentences that restate the thesis and the three takeaways.
3. Use transitions that hand off between slides ("…which raises the question: which
   table first?").
4. Add backup slides after the closing slide (not timed): raw OOPS! table; per-dimension
   expert and LLM-judge scores (Tables 10–11); cost/runtime (Table 15, quoting row
   values, not its printed totals, which do not add up); the retrieval ablation
   (Table 14); and a "Hard questions" slide with crisp, honest answers to:
   same-model judge in the loop · is the Non-Iterative baseline a strawman (its prompt
   asks for less) · LLM-judge fragment order and position bias · P19 and union merge ·
   why schema-only (privacy, portability) · coverage 0.77 · data leakage (the private
   liver DB) · self-referencing FKs in the traversal (Chinook Employee.ReportsTo).
5. Print a one-page rehearsal card (presentation/REHEARSAL_CARD.md): slide → cumulative
   time checkpoints at 3:00, 7:00, 10:30, 13:30, with what to cut if behind.
Commit and report the total word count and estimated duration.
```

## Phase 7: Final QA and delivery package

```
Read presentation/REVISION_PHASES.md (Shared context). Do a final pass on
presentation/RIGOR_ISWC2026.pptx:
1. Claims audit: every on-slide number versus the paper; list each with its source
   table and fix any mismatch.
2. Visual QA on full-resolution renders of every slide: overflow, overlap, < 0.5"
   margins, inconsistent title position, chart label styles that differ between
   charts, code panels with mismatched fonts or colors, and the method-section
   background constant.
3. Accessibility: contrast ≥ 4.5:1 for text, ≥ 3:1 for chart marks; alt text on images
   and diagrams; readable from the back of a room (body ≥ 16 pt, code ≥ 14 pt).
4. Fonts: confirm Georgia/Calibri/Consolas availability on Windows and macOS Office;
   produce presentation/RIGOR_ISWC2026.pdf as a fallback for the venue PC.
5. validate.py clean; slide numbers on; the final file name and the PDF committed.
Report the findings and fixes, then send me the final .pptx and .pdf.
```
