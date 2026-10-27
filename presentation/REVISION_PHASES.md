# RIGOR ISWC 2026 talk: revision plan in phases

Each phase below is a self-contained prompt. Run them **in order**, one per message
(paste the prompt as-is). Every phase reads the **Shared context** section first,
commits its output to branch `claude/bold-cori-h54gh3`, and ends by reporting what
changed plus anything that needs the author's decision.

Source files in the repo:

- Paper: `2506.01232v2 (1).pdf` (33 pages; the source of truth for every number)
- Current deck: `RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx` (23 slides, with speaker notes)
- Phase 1 output: `presentation/OUTLINE.md` (slide list, timings, results structure, claims audit). **Done.**
- From Phase 2 on, the working deck is `presentation/RIGOR_ISWC2026.pptx`, produced by
  editing a copy of the current deck (never rebuilt from scratch), with repeatable
  scripts in `presentation/tools/`.

---

## Shared context (read before every phase)

**Goal.** A 15-minute ISWC 2026 research-track talk strong enough to compete for best
paper. Spoken part ≤ 14:00 (aim for 13:30, leaving room for slower delivery). Q&A is separate.

**Author's rule: improve, don't destroy.** The author likes the current deck. Keep its
layouts, composition, running example, six-step build and the six cumulative
architecture slides (12–17). Every phase makes targeted edits to the existing slides;
none rebuilds the deck.

**Decisions already made (Phase 1):** 15 min speaking, Q&A separate, spoken target
13:30 · keep slides 12–17 · code excerpts use real Manchester Syntax **plus a
plain-English gloss** beside the key lines · no dedicated pitfalls slide; results
are restructured as R1 "meaning without losing structure" → R2 BURR → R3 experts
(see OUTLINE.md §2).

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

**Tooling rules.** Use the `anthropic-skills:pptx` skill's editing workflow: unzip →
edit `ppt/slides/slideN.xml` → zip. Use `add_slide.py` to duplicate slides and
`defusedxml`/lxml-safe edits, never ElementTree round-trips. Keep changes repeatable
as scripts in `presentation/tools/`. After every change run
`validate.py --original RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx`, render
to images, and inspect every changed slide for overflow and overlap. Speaker notes
live in the slides' notes pages.

---

## Phase 1: Storyline, slide list, 15-minute budget (DONE → OUTLINE.md)

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

## Phase 2: Restyle the existing deck (new identity, same layouts) (DONE)

Done via `presentation/tools/restyle.py` (rerunnable from the original deck) and
`presentation/tools/qa_render.py` (QA render that emulates Consolas widths).
Deviations from the prompt below: titles use **Cambria**, not Georgia (QA renders it with
a metric-compatible substitute, and its numerals line up). The step tracker uses
done = emerald tint, current = solid emerald, upcoming = hollow grey, so the current
step is the strongest. Later phases edit `presentation/RIGOR_ISWC2026.pptx` directly.

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/OUTLINE.md.
Copy RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx to
presentation/RIGOR_ISWC2026.pptx and RESTYLE it in place. Do not move, resize or delete
shapes, and do not change text, except where this prompt says so. The author likes
the current composition.

1. Palette, "lab notebook / editorial light", deliberately unlike the dark navy + orange
   template. Write presentation/tools/restyle.py with an explicit old→new hex map
   covering every srgbClr in the slides (inventory them first; there are about 25,
   mainly 111A28/182740/21314A navy, 84ADF7/4D71DE/7799F0 blue, F9AB5E/F4A261 orange).
   - Backgrounds: white FFFFFF. Panels/cards: F3F5F4. Hairlines: D9DEDB.
   - Ink: graphite 1F2328 for text, slate 5B6770 for secondary text.
   - Dominant accent: deep emerald 0B6E4F for RIGOR, highlights and the active step
     (replaces the orange).
   - Secondary accent: muted teal 3A7D8C for diagram strokes and links (replaces the
     blues).
   - Signal: vermilion D64C2B, only for errors and warnings (slide 10's wrong examples).
   - Title (1) and closing (23) slides: full emerald background with white text. The
     only colored backgrounds.
   Map dark-slide colors to their light equivalents (dark text on light), keeping the
   same visual hierarchy. Check contrast ≥ 4.5:1 for text.
2. Constant background through the method: slides 5–17 all get the same white
   background and identical frame. No dark/light switching anywhere between the
   overview and architecture 6/6. Other content slides are white too.
3. Fonts. Replace IBM Plex Sans and Aptos (not installed on venue PCs) with Calibri for
   body text and Georgia for slide titles (allow ~10% width slack; check wraps). Replace
   IBM Plex Mono with Consolas. Keep the sizes.
4. Step tracker (the "1 ORDER … 6 MERGE" row on slides 5–11): restyle in place so
   completed steps are emerald-filled, the current step has an emerald outline, and
   future steps are grey outlines. Same position on every slide.
5. Slide 5: add the confirmed teaser line (OUTLINE.md §6) under the step row, in the
   existing caption style.
6. Apply the claims-audit rewordings from OUTLINE.md §4 that are pure text edits:
   slide 18 "independent LLM judge" → "independent evaluation judge (GPT-5.4)";
   slide 10 gets a small visible "illustrative" tag in the existing caption style.
7. Do NOT yet touch code excerpts (Phase 3), results slides 19–21 (Phase 4), or the
   limitations slide 22 (Phase 5).
Validate (with --original), render all slides, compare side by side with the original
renders, and fix overflow, wrapping changes from the font swap, contrast and leftover
old colors. Commit restyle.py and the .pptx. Send me an old/new contact sheet.
```

## Phase 3: Ontology excerpts as real OWL code (DONE)

Done via `presentation/tools/code_excerpts.py` (run by `presentation/tools/build.sh`).
Every excerpt is parsed with the OWL API (`OwlCheck.java`) and checked for OWL 2 DL; the
verified files are in `presentation/excerpts/`. Slides 3, 7, 9 and 10 now show real
Manchester Syntax with plain-English glosses (slide 9 keeps its evidence tags instead).
Slides 11 and 23 show class diagrams, not pseudo-syntax, so they were left as they are.
Provenance uses `src:` = <http://example.org/provenance/admission/>, which expands to the
paper's IRI pattern.

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/OUTLINE.md, and
work on presentation/RIGOR_ISWC2026.pptx. Replace the pseudo-notation ontology excerpts
(slides 3, 7, 9, 10, 11 and the closing slide 23) with real OWL 2 Manchester Syntax,
INSIDE the existing panels (same position and size; the composition stays). The
audience must still understand them at a glance.

Code-panel style: background F6F8FA, 1 px border D0D7DE, rounded corners, a small tab
label (e.g. "admission · OWL Manchester Syntax"), Consolas ≥ 14 pt. Colors: keywords
(Class:, ObjectProperty:, DataProperty:, SubClassOf:, Domain:, Range:, Annotations:,
some) emerald 0B6E4F bold; names graphite; string literals A3401E; xsd:* 6F42C1; prov:*
slate.

Plain-English glosses (author's requirement): beside each key line, a short gloss in
the body font, slate, ≤ 6 words, aligned to that line with a thin leader, e.g.
  SubClassOf: :hasPatientData some :PatientData   ← every admission has a patient
  Range: xsd:integer                               ← stored as a number
  rdfs:label "Bone Marrow Depression"@en           ← human-readable name
At most 3–4 glosses per panel. Code lines without a gloss must be self-explanatory or
dropped.

Rules:
- One consistent namespace for the whole talk: Prefix: : <http://example.org/rigor#>,
  plus rdfs:, xsd:, prov:, owl:. Same IRIs on every slide (:Admission, :PatientData,
  :hasPatientData, :bmDep, :ecogStage).
- Correct keywords with colons: Class:, ObjectProperty:, DataProperty:, Domain:,
  Range:, SubClassOf:, Annotations:. Language-tagged labels ("…"@en). Existential
  restriction for a NOT NULL FK: SubClassOf: :hasPatientData some :PatientData.
- SEED slide (7) shows DMᵣ exactly as the deterministic mapping produces it (class, data
  properties with xsd ranges, FK as object property, existential restriction), with no
  labels or comments. GENERATE slide (9) shows the same fragment after enrichment;
  highlight the added lines (emerald gutter mark or a faint emerald line background),
  so the delta is visible at a glance, and keep the existing SCHEMA/DOCS evidence tags.
  Keep each panel ≤ 10 lines, using "…" instead of cramming.
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
Validate (--original), render, fix, commit. Show me before/after images of each changed slide.
```

## Phase 4: Results that show RIGOR's real strength (DONE)

Done via `presentation/tools/results.py` (run by `build.sh`). Final order: 18 Setup ·
19 R1 · 20 R2 · 21 R3 · 22 Limitations (Phase 5) · 23 Close · 24 B1 (raw OOPS! table).
Notes on R1/R2 are interim; Phase 6 writes the final script. BURR's metric description is
limited to what could be confirmed: systems output a database-to-ontology mapping, which
is compared with a gold-standard mapping, with F1 per class/relation/attribute. The full
BURR paper was not reachable from this environment.

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/OUTLINE.md §2–3,
and work on presentation/RIGOR_ISWC2026.pptx. Restructure the results into
R1 → R2 → R3, reusing the existing slides' styling (table style, hero-number style,
card style). Every number cites its paper table on the slide.

Reorder: old slide 20 (pitfalls) → R1 (replaced content), old 21 (BURR) → R2, old 19
(experts) → R3. Final order: 18 Setup, 19 R1, 20 R2, 21 R3, 22 Future work, 23 Close.

R1 "Only RIGOR adds meaning without losing structure" (replaces the pitfalls content).
  A 4×4 matrix (rows = what the ontology contains; columns = Direct Mapping, LLM
  baseline, Non-Iterative, RIGOR; RIGOR column highlighted emerald), values exactly as
  in OUTLINE.md §2:
    Columns kept, final coverage (Table 9): 0.78–0.99 | 0.31–1.00 | 0.06–0.78 | 0.77–1.00
    Labels & comments (Table 8):           0/3 | 0/9 | 9/9 | 9/9
    Class hierarchy, SubClassOf (Table 8):  0/3 | 0/9 | 0/9 | 9/9
    Provenance to source column (Table 8): 0/3 | 0/9 | 0/9 | 9/9
  Recompute each cell from the paper tables before writing it. Emphasize Non-Iterative's
  0.06 (single-pass RAG drops columns). One line below: "Modeling pitfalls (OOPS!): fewer
  than every other method in 9/9 like-for-like comparisons (same database, same LLM);
  zero on Chinook with all three LLMs." Caption: "coverage = embedding-based proxy;
  runs = 3 databases × 3 LLMs".

R2 "RIGOR keeps the link back to the data: best on BURR" (rework old slide 21; keep
  its F1 table on the right). Add on the left "What BURR checks": a mini worked example
  in the deck's table/class style. An ISWC table (e.g. person) maps to a class, a foreign
  key to a relation, and a column to an attribute; the system's mapping is compared with
  BURR's expert gold mapping, giving an F1 per element type. Add one line on why it
  matters: "Enrichment must not break the link to the data; OBDA and KG population depend
  on it." Add one line on how RIGOR is scored: its prov:wasDerivedFrom links are exported
  as a D2RQ mapping. Caption: "BURR's authors found LLM methods trail rule-based ones;
  here RIGOR leads both · one scenario, Claude." Before writing the left half, check
  how BURR matches elements (Laskowski et al., PACMMOD 2025; fetch it if the network
  allows). State only what you can confirm.

R3 "Experts prefer RIGOR" (old slide 19 moved; content unchanged apart from the
  restyle).

Backup: move the raw OOPS! table (the old slide 20 grid, Table 7) after the closing
slide as B1.
Update the speaker notes of the three slides to match (final wording comes in Phase 6).
Validate (--original), render, fix, commit, and show me the three slides plus B1.
```

## Phase 5: Future work, from schemas to instances (cisplatin example)

```
Read presentation/REVISION_PHASES.md (Shared context) and presentation/OUTLINE.md, and
work on presentation/RIGOR_ISWC2026.pptx. Replace the content of the limitations slide (22),
keeping the deck's card/flow styling, with a future-work slide: "Next: let the data speak,
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
Validate (--original), render, fix, commit, and show me the slide.
```

## Phase 6: The speech, timing, and Q&A backups

```
Read presentation/REVISION_PHASES.md (Shared context) and OUTLINE.md, and work on
presentation/RIGOR_ISWC2026.pptx. Use OUTLINE.md §3's per-slide time and word budgets.
1. Write the full spoken script, slide by slide, into presentation/SPEECH.md and into
   each slide's notes: first-person, conversational, short sentences, one idea per
   sentence, no reading of slide text. At the deck's own pace (~105 words/min, pauses
   included) the whole script is ≈ 1,420 words, ending by 13:30. Per slide: [start–end] timestamp, word count, the script, one
   pointer cue, one-line takeaway. Mark 2–3 deliberate pauses (after the 70%, after
   "0 pitfalls on Chinook").
2. Opening hook in the first 20 s (the inherited clinical database with BM_DEP).
   Closing in 3 sentences that restate the thesis and the three takeaways.
3. Use transitions that hand off between slides ("…which raises the question: which
   table first?").
4. Add backup slides after the closing slide (not timed), after B1 (raw OOPS! table): per-dimension
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
