# Open items for the next revision

The author is editing the deck by hand and will send the edited file for revision. That file
becomes the new source: revise it directly, and do not regenerate the deck with
`presentation/tools/build.sh`, which rebuilds from the original deck and would discard the
author's edits.

## To fix

1. **Admission table has no primary key.** The author removed the "excerpt" labels, so the table
   now reads as complete. Slide 2 marks `patient_id` as PK in PatientData but shows no PK for
   admission. A table without a PK maps rows to blank nodes under the W3C Direct Mapping.
   - Add `admission_id  PK` as the first row of the admission table on every slide that shows it
     (2, 3, 5, 6, 7, and the cover figure if relevant).
   - Seed code (slide 7, and the compact copy on slide 9): add
     `admission_id some xsd:integer,   ← primary key: every admission has an id`
     (paper §3.2: primary-key columns get an existential restriction). Verify with the OWL API.
   - Check that slide 9's compact "before" panel still fits.

## Suggested, waiting for the author's decision

2. **Slide 3, "What this enables":** use the standard terms "Ontology-based data access (OBDA)",
   "Knowledge graph population", "Semantic data integration".
3. **Slide 5, "+ justified axioms / links":** replace with
   "+ logical axioms (hierarchy, disjointness)  + links to standard vocabularies".
4. **Slide 8, retrieval strip:** add the running example, e.g.
   "(e.g., for admission table → query: admission + patient_id, BM_DEP, ecog_stage + PatientData)".
5. **Slide 8, middle card:** "Disease Ontology: bone marrow disease" as the external-ontology hit.
6. **Slide 9 speech:** "We prompt the Gen-LLM with the seed, the retrieved context and our
   enrichment instructions." (already in the generated notes).

## Resolved

- Slide 9 rebuilt on the author's edited slide (`tools/slide9_diff.py`, applied directly, not via
  `build.sh`): BEFORE (direct mapping, with `admission_id` PK) and AFTER (delta ontology) share one
  row grid; insertions fill the empty rows (mint band, emerald mark, evidence tag), renamed names
  are highlighted, unchanged lines are plain. Both sides verified with the OWL API. Header
  "BEFORE / DIRECT MAPPING" replaces the caption "Structure only. Meaning remains implicit.";
  speaker notes updated. Output: `RIGOR_slide9_generate.pptx`.

- Slide 3 reveal (author's version: six separate slides for the PDF) rebuilt with
  `tools/slide3_steps.py`: the target shows the final slide-9 fragment (19 rows, fixed grid,
  later-step rows left empty), yellow highlight on what each step adds, glosses in their own
  boxes (Google Slides drops tab stops), `admission_id PK` added to the input table. All four
  reveal states verified with the OWL API. Output: `RIGOR_slide3_steps.pptx`. Still to do on the
  other slides that show the table (2, 5, 6, 7 and the slide 7 seed): the PK row.

## Author's decisions to carry into the next revision

- **`DisjointWith: PatientData` removed** from the running example (the author: it does not make
  sense for this domain). Remove it everywhere the example appears: slide 9 AFTER panel and
  `code_excerpts.DIFF_ROWS` (slide 3 steps and slide 9 are generated from it), then re-verify.
  Without it, slide 3's LLM step adds only the ECOG comment line.
- **No "LLM knowledge" gloss on slide 3**: LLMs have not been introduced at that point in the talk.
  Slide 9 (after the Gen-LLM is introduced) can keep its LLM tag on the ECOG comment.
  Open point: what slide 3 says, if anything, beside the ECOG comment line instead.

## Decided: keep as is

- Slide 7 seed restrictions (`some` / `only`), and their glosses.
