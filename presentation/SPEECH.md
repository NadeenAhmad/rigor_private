# RIGOR talk: spoken script (final deck)

Timing model: 105 words per minute (clicks included) plus 2 s per [pause]. Total: 1394 words, ending at 13:27. Each physical slide's notes hold only what is said while it is on screen.

## 1. Title  (0:00–0:25)

Good morning. I'm Nadeen Fathallah, from the University of Stuttgart. Relational databases run hospitals, banks and governments. But their schemas tell us how data is stored, not what it means. This talk is about RIGOR: turning those schemas into rich OWL ontologies, automatically.

> Calm, slow start. Do not read the author list.

## 2. Motivation  (0:25–0:55)

Imagine you inherit this clinical database. An admission points to its patient through a foreign key. But how should an ontology represent that link? What does BM_DEP mean? What does an ECOG stage of three mean? The schema cannot tell us. Yet these meanings are exactly what ontology-based access, knowledge graphs and integration need.

> Point at the FK arrow, then the three questions on the right, top to bottom.

## 3. What we want to build  (0:55–1:48)

**Reveal 1/6.** Here is what we want to build from this table.

> Empty target. One breath, then click.

**Reveal 2/6.** First, from the schema alone: admission becomes a class with a real relation to PatientData, and every property points back to its source column.

> The class box and arrow, then the two 'source column' lines.

**Reveal 3/6.** The documentation adds human-readable labels: Bone Marrow Depression, ECOG Performance Status.

> The two highlighted labels.

**Reveal 4/6.** A standard vocabulary links BM_DEP to bone marrow disease in the Disease Ontology.

> The skos:relatedMatch line.

**Reveal 5/6.** And a comment says what the ECOG codes mean, from zero, fully active, to five, dead.

> The rdfs:comment line. Do not mention LLMs yet.

**Reveal 6/6.** That is a usable semantic model. The challenge: add all this without losing a single column or link.

> The 'What this enables' box, then pause.

## 4. The research gap  (1:48–2:21)

Existing work solves half of this. Schema-first rules, like direct mapping, preserve the structure: every table, column and foreign key. But the domain meaning stays missing. LLMs can suggest the meaning: labels, concepts, relations. But they omit elements or model them incorrectly. So our question: can we keep schema fidelity and add domain semantics, across a whole database?

> Left card, right card, then land on the question.

## 5. Our answer: RIGOR  (2:21–2:37)

Our answer is RIGOR. It builds the ontology one table at a time, in six steps: order, seed, retrieve, generate, validate, merge. Let's build our admission example together.

> From 'what we have' to 'what we want', then sweep along the six steps.

## 6. Step 1: order  (2:37–3:05)

Step one: order. Admission points to PatientData, so we build PatientData first. In general, RIGOR follows the foreign keys breadth-first: a table is processed once the tables it references are done. So when we reach admission, PatientData already exists, and we reuse it instead of inventing a duplicate.

> Follow the FK arrow, then the processing order.

## 7. Step 2: seed  (3:05–3:41)

Step two: seed. Before any LLM is involved, a deterministic direct mapping turns the table into OWL. Admission becomes a class. The foreign key becomes an object property, and every admission must have a patient. Each column becomes a data property with its XSD type. So every column is in the seed, by construction. The LLM's job is to enrich, not to invent.

> Read the glosses: 'each has a patient', 'the FK becomes a link'.

## 8. Step 3: retrieve  (3:41–4:24)

**Reveal 1/4.** The seed is faithful, but it does not say what BM_DEP means. So step three retrieves context from three sources. Documentation defines cryptic names, like BM_DEP.

> Documentation card.

**Reveal 2/4.** External ontologies, like the Disease Ontology, suggest established concepts.

> External ontologies card.

**Reveal 3/4.** And the third source is the distinctive one: the ontology we have built so far. Earlier tables become evidence for the next.

> Growing core card; slow down here.

**Reveal 4/4.** One query per table, built from its name, columns and foreign-key targets, retrieves the top matches from each source.

> The query strip.

## 9. Step 4: generate  (4:24–5:15)

**Reveal 1/2.** Step four: generate. We prompt the Gen-LLM with the direct mapping from step two, the retrieved context, and our enrichment instructions.

> The subtitle, then the BEFORE panel.

**Reveal 2/2.** What comes back is a delta ontology: the same lines, with new ones filled into the gaps. The tags show where each addition came from. The foreign key becomes the named relation hasPatientData, and provenance links each property to its column: both from the schema. The labels come from the documentation, the link to bone marrow disease from the Disease Ontology, and the ECOG scale from the LLM itself.

> Left to right on one row, then down the highlighted lines and their tags.

## 10. Step 5: validate  (5:15–6:18)

**Reveal 1/3.** Step five: validate. Plausible fragments can still break the ontology. A later table, lab_result, reuses hasPatientData with a second domain. OWL would then infer that every lab result linked to patient data is an admission. Rule-based graph checks catch this.

> Structural card: the red line, then the explanation.

**Reveal 2/3.** Other errors are semantic. Here the model links admission to a new class, Patient, although the core already has PatientData. Valid OWL, but now two classes mean the same thing. The Judge-LLM sees the core and asks to reuse it.

> Semantic card.

**Reveal 3/3.** Failed fragments go back to the Gen-LLM, with up to two retries. If one still fails, the Judge-LLM picks the better of its last version and the plain direct mapping.

> The correction strip.

## 11. Step 6: merge  (6:18–6:43)

Step six: merge. The validated delta joins the core ontology. Because RIGOR reuses identifiers, this is a simple union of axioms. And here is the loop: the expanded ontology becomes a retrieval source for the next table. Each table builds on the ones before.

> Core plus delta, then the loop strip.

## 12. Reassembling the architecture  (6:43–7:22)

**Reveal 1/6.** Let me put the pieces together, as in the paper's Figure 1. First, the schema decides the order.

> Quick clicks from here. Reconnect, do not re-explain.

**Reveal 2/6.** Each table gets its deterministic seed.

> Direct mapping box.

**Reveal 3/6.** Three channels retrieve meaning: external ontologies, documentation, and the core built so far.

> The three retrieval boxes.

**Reveal 4/6.** The Gen-LLM combines seed and context into a delta ontology.

> Gen-LLM.

**Reveal 5/6.** A Judge-LLM and rules validate it.

> Validate + refine.

**Reveal 6/6.** And the merged core feeds retrieval for the next table. That loop is RIGOR.

> Trace the return arrow from (i) back to (e).

## 13. Evaluation setup  (7:22–8:20)

**Reveal 1/7.** Does this work beyond one example? We tested two medical databases: a private liver-cancer registry from Leipzig, and the large public eICU intensive-care database.

> Liver Cancer, then eICU-CRD. Do not read the numbers.

**Reveal 2/7.** And Chinook, a music store, as a second domain.

> Chinook card.

**Reveal 3/7.** We compare four methods: direct mapping; an LLM baseline that sees only the table; Non-Iterative RAG, which adds retrieval in a single pass; and RIGOR.

> The four method cards, left to right.

**Reveal 4/7.** The three LLM methods each run with three models, Claude, Mistral and DeepSeek, on the three databases: twenty-seven ontologies.

> The 27 line.

**Reveal 5/7.** Plus three direct mappings,

> The 3 line.

**Reveal 6/7.** thirty ontologies in total.

> The 30 line.

**Reveal 7/7.** We check them with automatic quality checks, an LLM judge, ten human experts, and the BURR benchmark.

> The evaluation strip.

## 14. Result 1: valid, fewer errors  (8:20–8:54)

We go from the easiest check to the hardest. First, the basics: all thirty ontologies parse without errors and are logically consistent. Modeling errors are another story. Direct mapping averages about five hundred OOPS! pitfalls per ontology, RIGOR eight. Same LLM, same database: Mistral on eICU makes two hundred twenty with Non-Iterative RAG, two with RIGOR. [pause]

> One breath for validity; then the averages; then the Mistral eICU column.

## 15. Result 2: keeps the schema  (8:54–9:29)

**Reveal 1/2.** Fewer errors, but is anything lost? On average, RIGOR covers ninety-three percent of the columns, as much as direct mapping. Non-Iterative RAG, with the same retrieved context but no seed, covers only a third.

> RIGOR's average next to Direct Mapping's, then the Non-Iterative RAG average.

**Reveal 2/2.** And it swings with the LLM, from six to seventy-eight percent. With RIGOR, every LLM keeps at least seventy-seven percent, in every database. [pause]

> The outlined Non-Iterative RAG row, then the strip.

## 16. Result 3: adds meaning  (9:29–9:55)

Nothing is lost. Is meaning added? Direct mapping and the LLM baseline add no labels, provenance or hierarchy, in any run. Non-Iterative RAG adds labels, but no provenance and no hierarchy. Only RIGOR adds all three, with every LLM, in every database. [pause]

> The red band first, then the three RIGOR rows.

## 17. Result 4: BURR  (9:55–10:37)

Meaning is added. But is the link back to the data still correct? BURR, an external benchmark, checks exactly this: every class, relation and attribute must map back to the right table, foreign key or column, scored with F1 against gold mappings. On BURR's ISWC scenario, RIGOR has the highest F1 on all three, ahead of rule-based and LLM-based systems. Against the LLM-based OntoGenix, relations nearly double, because foreign keys become named relations, like hasPatientData.

> Left card first, then land on the RIGOR row.

## 18. Result 5: LLM judge and experts  (10:37–11:29)

Automatic checks agree. Do evaluators agree? For each table, an LLM judge and ten ontology experts saw the schema, five competency questions, the questions the ontology must answer, and four anonymized ontologies. They scored each on seven dimensions and ranked them first to fourth. The LLM judge put RIGOR first in every ranking. [pause] The experts put RIGOR first in seventy percent of their rankings, with the highest average score. [pause] Direct mapping gives structure without meaning. LLMs alone give meaning without structure. RIGOR gives both.

> The 'How they judged' box in one breath; then the two headlines.

## 19. Future work: validators in the loop  (11:29–12:08)

**Reveal 1/2.** Where next? First, the validators. RDFLib, the reasoner and OOPS! only checked our final ontologies; NeOn-GPT first proposed this stack of checks for evaluation. Next, they flag each error while RIGOR builds, and their feedback prompts the LLM to repair the ontology.

> The sentence on the slide, then click.

**Reveal 2/2.** We have realized this loop in the RIGOR App, presented in yesterday's demo session. The code and a demo video are behind these QR codes.

> Screenshot, then the two QR codes.

## 20. Future work: let the data speak  (12:08–12:54)

**Reveal 1/2.** Second, let the data speak. Today RIGOR reads schemas and documentation, but instance data can propose new knowledge. Here, the treatment table says which admissions received cisplatin, and the admission table records bone marrow depression.

> The two tables; A1 and A3 in both.

**Reveal 2/2.** Joined on the admission id, every cisplatin admission has bone marrow depression. That link is not in the documentation, so RIGOR can propose it as candidate knowledge, with provenance to both columns. A pattern, not a cause: an expert reviews it before it joins the ontology.

> The pattern bars, then the candidate code and its badge.

## 21. Thank you  (12:54–13:27)

To conclude. RIGOR anchors LLM generation in the database itself, grounds it in retrieved knowledge, and checks every step before the ontology grows. The result: meaning without losing structure, the link to the data intact, and ontologies that experts prefer. The paper and code are behind these QR codes. Thank you. I'm happy to take your questions.

> Finish speaking by 13:30. Leave this slide up during questions.
