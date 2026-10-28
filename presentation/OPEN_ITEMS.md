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
4. **Slide 9 provenance prefix:** slide 9 uses `src:ecog_stage`, while slide 3 now uses
   `admission:`. Use `admission:` on both.

## Decided: keep as is

- Slide 7 seed restrictions (`some` / `only`), and their glosses.
