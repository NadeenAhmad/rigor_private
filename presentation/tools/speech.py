#!/usr/bin/env python3
"""Phase 6: the spoken script, as the single source for speaker notes, SPEECH.md and the
rehearsal card.

Timing model (the original deck's own pace): 105 spoken words per minute, which already
includes normal clicks and visual pauses, plus 2 s for every explicit [pause].

Usage:
  python3 speech.py apply <in.pptx> <out.pptx>   # write notes for slides 1-23, closing takeaways
  python3 speech.py docs                         # write presentation/SPEECH.md and REHEARSAL_CARD.md
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from results import NS, R, set_notes, set_runs  # noqa: E402

WPM, PAUSE_S = 105, 2
OUT_DIR = Path(__file__).resolve().parent.parent

# One entry per slide, in presentation order.
#   say: the spoken script   cue: pointer / delivery   takeaway   source
#   cut: an optional sentence (must appear verbatim in `say`) to drop if running late
SLIDES = [
    dict(title="Title",
         say="Good morning. I'm Nadeen Fathallah, from the University of Stuttgart. Relational databases run "
             "hospitals, banks and governments. But their schemas tell us how data is stored, not what it "
             "means. This talk is about RIGOR: turning those schemas into rich OWL ontologies, automatically.",
         cue="Calm, slow start. Do not read the author list.",
         takeaway="Schemas store data, not meaning.",
         source="Paper Section 1."),
    dict(title="Motivation",
         say="Imagine you inherit this clinical database. Admission points to PatientData through a foreign "
             "key. That link is clear. But look at the columns. What is BM_DEP? What does an ECOG stage of three "
             "mean? The schema cannot tell you. Yet these meanings are exactly what ontology-based access, "
             "knowledge graphs and data integration need. So our problem is not reading the SQL. It is "
             "recovering the domain meaning, while keeping the structure we can trust.",
         cue="Point at the FK arrow, then BM_DEP, then ecog_stage. Let the questions hang for a second.",
         takeaway="The schema knows storage and links, not clinical meaning.",
         source="Illustrative table assembled from examples in paper Sections 3.2 and 3.4."),
    dict(title="What we want to build",
         say="Here is what we want. [click] First, what the schema alone gives us: Admission becomes a class "
             "with a real relation to PatientData, and every property points back to its source column. [click] "
             "The documentation adds human-readable labels: Bone Marrow Depression, ECOG Performance Status. "
             "[click] A standard vocabulary adds a link to the Disease Ontology. [click] And the LLM's own "
             "knowledge adds what the ECOG codes mean, and that an admission is not a patient. [click] That is a "
             "usable semantic model. The challenge is to add all this without losing a single column or link.",
         cue="Five clicks, one per knowledge source: schema, documentation, Disease Ontology, LLM knowledge, "
             "then 'What this enables'. The gloss colours match the sources.",
         takeaway="A usable model combines schema, documentation, standard vocabularies and LLM knowledge.",
         source="Illustrative fragment, verified with the OWL API at every reveal step "
                "(presentation/excerpts/slide03_*.omn). obo:DOID_4961 = Disease Ontology 'bone marrow disease'; "
                "skos:relatedMatch, not exactMatch, because BM_DEP is a yes/no column, not the disease. "
                "Provenance is column-level, as defined in paper Section 3.2."),
    dict(title="Research gap",
         say="Existing work solves half of this. Schema-first rules, like direct mapping, preserve the "
             "structure: every table, column and foreign key. But they add no domain meaning. LLMs can suggest "
             "meaning: labels, concepts, relations. But they omit columns, or model keys wrongly. So our question "
             "is: can we keep schema fidelity and add domain semantics, across a whole database, with no expert "
             "in the loop during generation?",
         cue="Left card, right card, then land on the question.",
         takeaway="Rules keep structure; LLMs suggest meaning; we need both.",
         source="Paper Sections 1-2. References on the slide (all cited in paper Section 2): Direct Mapping "
                "[Arenas 2012; Sequeda 2012], BootOX [ISWC 2015], Ben Mahria et al. 2021, RODI [SWJ 2017]; LLMs4OL "
                "[ISWC 2023], Mateiu & Groza 2023, OntoGenix 2025, LLM4VKG [IJCAI 2025]. BURR deliberately not cited "
                "here: its results do not show LLMs trailing rule-based systems."),
    dict(title="Our answer: RIGOR",
         say="Our answer is RIGOR. It builds the ontology one table at a time, in six steps: order, seed, "
             "retrieve, generate, validate, merge. And as a preview: experts ranked it first in seventy percent "
             "of their rankings. Let's build it on our example.",
         cue="Sweep along the six steps. Point at the result preview once.",
         takeaway="Six steps, one table at a time.",
         source="Paper Section 3.1, Fig. 1; preview from Table 1."),
    dict(title="Step 1: order",
         say="Step one: order. Admission points to PatientData, so we build PatientData first. In general, "
             "RIGOR follows the foreign keys breadth-first: a table is processed once the tables it references "
             "are done. Why? By the time we reach Admission, PatientData already exists in our ontology. So we "
             "reuse that class, instead of inventing a second, incompatible one.",
         cue="Follow the FK arrow, then the processing order.",
         takeaway="Parents first, so later tables reuse known classes.",
         source="Paper Sections 3.1-3.2 (FK-guided breadth-first order)."),
    dict(title="Step 2: seed",
         say="Step two: seed. Before any LLM is involved, a deterministic direct mapping turns the table into "
             "OWL. Admission becomes a class. The foreign key becomes an object property, and every admission "
             "must have a patient. Each column becomes a data property with its XSD type. This seed covers every "
             "column of the cleaned table, by construction. So the LLM's job is to enrich, not to invent.",
         cue="Read the glosses: 'each has a patient', 'the FK becomes a link'.",
         takeaway="A deterministic seed guarantees coverage before the LLM starts.",
         source="Paper Section 3.2. The guarantee holds for the seed; final coverage is reported later.",
         cut="So the LLM's job is to enrich, not to invent."),
    dict(title="Step 3: retrieve",
         say="The seed is faithful, but it still doesn't say what BM_DEP means. So step three retrieves context "
             "from three sources. Documentation explains cryptic names and codes. External domain ontologies "
             "suggest established concepts. And the third source is the distinctive one: the ontology we have "
             "built so far. Earlier tables become evidence for the next. One query per table, built from its "
             "name, columns and foreign keys, retrieves the top three matches per source.",
         cue="Walk the three cards; slow down on the growing core.",
         takeaway="Three sources supply meaning; the growing core is unique to RIGOR.",
         source="Paper Section 3.3 (all-MiniLM-L6-v2, Faiss, k = 3 per source).",
         cut="One query per table, built from its name, columns and foreign keys, retrieves the top three "
             "matches per source."),
    dict(title="Step 4: generate",
         say="Step four: generate. We prompt the Gen-LLM with the direct mapping from step two, the retrieved "
             "context and our enrichment instructions. What comes back is a delta ontology: the same lines as on "
             "the left, with new ones inserted into the gaps. The tags show where each change came from. The "
             "foreign key becomes the named relation hasPatientData, and provenance links each property to its "
             "column: both from the schema. The labels come from the documentation, the link to bone marrow "
             "disease from the Disease Ontology, and the disjointness and the ECOG scale from the model's own "
             "knowledge.",
         cue="Left to right on one row, then down the highlighted lines and their tags.",
         takeaway="The LLM enriches a grounded draft, and every addition is traceable.",
         source="Paper Sections 3.2-3.4. Row-aligned BEFORE/AFTER, both verified with the OWL API."),
    dict(title="Step 5: validate",
         say="A capable LLM rarely gets syntax wrong. The real risk is a fragment that looks right but clashes "
             "with the ontology built so far. So step five validates every delta before it is merged. On the "
             "left, the model names the target class Patient, although the core already has PatientData. That is "
             "valid OWL, but now two classes mean the same thing. The Judge-LLM sees the core and asks to reuse "
             "it. On the right, a later table reuses hasPatientData with a second domain. OWL then infers that "
             "every lab result is an admission. Rule-based graph checks catch this. Failed fragments go back for "
             "correction, with up to two retries.",
         cue="Left, right, then the correction strip.",
         takeaway="The Judge catches clashes in meaning; rules catch clashes in structure.",
         source="Paper Section 3.5 (Judge criterion 9: core ontology consistency; graph check: domain/range "
                "cardinality). Two domains = OOPS! pitfall P19, still reported in 3/9 RIGOR runs (Table 7). "
                "Examples verified with the OWL API (presentation/excerpts/slide10_*.omn).",
         extra=["IF ASKED: P19 is exactly the error that sometimes still slips through (3 of 9 runs); feeding "
                "validator output back into generation is our next step (slide 22)."]),
    dict(title="Step 6: merge",
         say="Step six: merge. The validated delta joins the core ontology. Because RIGOR reuses identifiers, "
             "this is a simple union of axioms. And here is the loop. The expanded ontology becomes a retrieval "
             "source for the next table. RIGOR never starts from scratch: each table builds on the ones before.",
         cue="Core plus delta, then the arrow back to retrieval.",
         takeaway="The ontology is both the output and the next input.",
         source="Paper Section 3.6."),
    dict(title="Architecture 1/6",
         say="Let me put the pieces together, as in the paper's Figure 1. First, the schema decides the order.",
         cue="Quick click. Reconnect, do not re-explain.", takeaway="Order.", source="Paper Fig. 1(a)."),
    dict(title="Architecture 2/6",
         say="Each table then gets its deterministic seed: the structural anchor for everything that follows.",
         cue="Quick click.", takeaway="Seed.", source="Paper Fig. 1(b)."),
    dict(title="Architecture 3/6",
         say="Three retrieval channels add meaning: external ontologies, documentation, and the core ontology "
             "built so far.",
         cue="Quick click.", takeaway="Retrieve.", source="Paper Fig. 1(c-e)."),
    dict(title="Architecture 4/6",
         say="The Gen-LLM combines seed and context into a delta ontology.",
         cue="Quick click.", takeaway="Generate.", source="Paper Fig. 1(f-g)."),
    dict(title="Architecture 5/6",
         say="A Judge-LLM and rule-based checks validate the delta before it can join.",
         cue="Quick click.", takeaway="Validate.", source="Paper Fig. 1(h)."),
    dict(title="Architecture 6/6",
         say="And the validated delta is merged into the core, which feeds retrieval for the next table. That "
             "loop is RIGOR.",
         cue="Trace the green return arrow from (i) back to (e).", takeaway="Merge, and loop.",
         source="Paper Fig. 1(i)."),
    dict(title="Experimental setup",
         say="Does this work beyond one example? We tested three databases: a private liver-cancer registry, the "
             "large eICU intensive-care schema, and Chinook, a music store, as a second domain. Three generator "
             "LLMs. And four methods: direct mapping, a zero-shot LLM, single-pass retrieval, and RIGOR. That "
             "gives thirty ontologies, judged by quality checks, an independent evaluation judge, ten human "
             "experts, and the BURR benchmark.",
         cue="Datasets, then the four methods. Do not read every number.",
         takeaway="Three databases, three LLMs, four methods, four kinds of evidence.",
         source="Paper Sections 4.1-4.6. 30 ontologies = 3 LLMs x 3 methods x 3 databases + 3 direct mappings."),
    dict(title="Result 1: valid, fewer errors",
         say="We climb from the easiest check to the hardest. First, the basics: all thirty ontologies are valid "
             "OWL and logically consistent, for every method. Modeling errors are another story: "
             "direct mapping averages about five hundred per ontology, RIGOR eight. Look at this column: same "
             "model, same database. The small Mistral model on eICU makes two hundred twenty errors without RIGOR, "
             "two with it. [pause]",
         cue="One breath for validity; then the averages; then point at the Mistral eICU column.",
         takeaway="Every ontology is valid; RIGOR's have far fewer modeling errors, with every LLM.",
         source="Paper Sec. 4.7 (RDFLib, HermiT) and Table 7 (OOPS! totals per run). 72% of the other methods' "
                "pitfalls are P08, missing annotations.",
         cut="for every method",
         extra=["IF ASKED ABOUT CRITICAL ERRORS: P19 (multiple domains or ranges) is the only pitfall OOPS! rates "
                "critical, and it is the second-domain error from the validation slide. Per ontology: Direct "
                "Mapping 15, LLM baseline 17, Non-Iterative 5.4, RIGOR 2.8; zero in 6 of RIGOR's 9 ontologies "
                "(Table 7). RIGOR's prompt allows one domain and one range, and the graph check rejects a second."]),
    dict(title="Result 2: keeps the schema",
         say="Fewer errors. But is anything lost? On average RIGOR covers ninety-three percent of the database "
             "columns, as much as direct mapping. Non-Iterative, with the same retrieved context but no seed, "
             "covers a third, and swings with the LLM, from six to seventy-eight percent. With RIGOR, every LLM "
             "keeps at least seventy-seven percent, in every database. [pause]",
         cue="RIGOR's average next to Direct Mapping's; then the red Non-Iterative row.",
         takeaway="RIGOR keeps the schema as well as Direct Mapping, with every LLM.",
         source="Paper Table 9 (column coverage per run; embedding similarity >= 0.55).",
         cut="and swings with the LLM, from six to seventy-eight percent",
         extra=["IF ASKED WHY NON-ITERATIVE DROPS BUT THE BASELINE DOES NOT: the baseline only translates the table, "
                "so it keeps column names but adds no meaning. Non-Iterative does everything in one pass with a lot "
                "of retrieved context and silently drops properties (e.g. 6 data properties for 350 Liver Cancer "
                "columns with Mistral, Table 8); nothing checks coverage. RIGOR starts from a seed with every "
                "column, works one table at a time, and its judge rejects fragments that drop a column.",
                "IF ASKED ABOUT MODEL SIZE AND COST: with RIGOR, Mistral Small 24B reaches 98% coverage (Claude 95%) "
                "for about $0.12 across the three databases, against about $33 with Claude (Tables 9, 15)."]),
    dict(title="Result 3: adds meaning",
         say="Nothing is lost. Is meaning added? Direct mapping and the LLM baseline add none, in any run. "
             "Non-Iterative RAG adds labels, but no provenance and no class hierarchy. Only RIGOR adds all three, "
             "with every LLM, in every database. [pause]",
         cue="Point at the red block first, then the three RIGOR rows.",
         takeaway="Only RIGOR adds labels, provenance and a class hierarchy, with every LLM.",
         source="Paper Table 8. Hierarchy as inheritance richness (subclass axioms per class; OntoQA, Tartir et al. "
                "2005; OntoMetrics, Lantow 2016), computed from Table 8 counts. Shading is within each database.",
         cut="Non-Iterative RAG adds labels, but no provenance and no class hierarchy.",
         extra=["IF ASKED 'DOES CHINOOK NEED SO MANY SUBCLASSES?': Count is not quality, so we report whether a "
                "hierarchy exists, and only RIGOR builds one. RIGOR's Chinook ontologies have zero OOPS! pitfalls and "
                "the judge ranks RIGOR first on all 11 Chinook tables; it does score RIGOR's conciseness lower there "
                "(2.59 vs 3.32 for the baseline, Table 10), so tighter hierarchies are future work.",
                "DOMAIN PATTERN (notes only): inheritance richness is 0.10-0.70 on the clinical schemas and 2.1-4.1 "
                "on the music store; the LLM proposes more specializations in a familiar domain."]),
    dict(title="Result 4: BURR",
         say="Meaning is added. But is the link back to the data still correct? BURR, an external benchmark, "
             "checks exactly this: every class, relation and attribute must map back to the right table, foreign "
             "key or column, scored with F1 against a gold mapping. On BURR's ISWC scenario, RIGOR has the highest F1 on all three, ahead of rule-based and "
             "LLM-based systems. Relations stand out: zero point six seven against zero point one one, because "
             "foreign keys become named relations, like hasPatientData.",
         cue="Left card first, then land on the RIGOR row.",
         takeaway="Semantic enrichment did not break the correspondence that OBDA and KG population need.",
         source="Paper Table 13. BURR = Laskowski et al., PACMMOD 2025; matching rule from BURR Section 4, Eq. 3. "
                "Left-card element names are illustrative.",
         cut="scored with F1 against a gold mapping",
         extra=["DO NOT SAY: 'BURR found LLMs trail rule-based systems'. BURR's results show Gemini 2.5 Pro ahead "
                "of the rule-based systems overall (rule-based systems win only on NPD).",
                "IF ASKED ABOUT GEMINI 2.5 PRO (BURR Table 3, ISWC .65 / .68 / .56): in BURR it is a "
                "general-purpose LLM prompted with few-shot examples and a checklist to write the D2RQ mapping "
                "directly. It is a mapping generator, not an ontology-learning system that builds an enriched "
                "ontology, so we compare with ontology-learning systems. Against it, RIGOR's class F1 is higher, "
                "relations are on par, attributes lower, and RIGOR also produces labels, hierarchy and "
                "provenance. It is prompted, not fine-tuned."]),
    dict(title="Result 5: LLM judge and experts",
         say="Automatic checks agree. Do evaluators agree? Both judged each table against its competency "
             "questions: what the ontology must be able to answer, five per table, verified by a domain expert. "
             "An LLM judge ranked RIGOR first on all fifty-eight tables, with its largest lead on the liver "
             "registry that no LLM has seen. [pause] Ten ontology experts ranked RIGOR first in twenty-eight of "
             "forty rankings. [pause] Direct mapping gives structure without meaning. LLMs alone give meaning "
             "without structure. RIGOR gives both, and checks them.",
         cue="Read the competency question aloud; then 58 of 58; then 28 of 40; then the closing strip.",
         takeaway="Judged against competency questions, an LLM judge and human experts both rank RIGOR first.",
         source="Paper Sec. 4.2 (5 CQs per table, LLM-generated, expert-verified; 290 in total), Table 6 (example), "
                "Table 1 (scores, rankings), Table 12 (Kendall's W 0.39-0.72). Judge: Claude ontologies.",
         cut="with its largest lead on the liver registry that no LLM has seen"),
    dict(title="Future work",
         say="Where next? Let the data speak. Remember BM_DEP? Without documentation, the name alone is a guess. "
             "The data gives two clues: it is only ever zero or one, so it is a flag; and it is one exactly when "
             "the patient received cisplatin, a chemotherapy known to suppress the bone marrow. So RIGOR can "
             "propose: bone marrow depression, with provenance to both columns, for an expert to confirm.",
         cue="The '?' on BM_DEP, then the two clues, then the reveal.",
         takeaway="When documentation is missing, instance data can explain cryptic columns.",
         source="Paper Sec. 5 (instance-level signals for cryptic columns when documentation is unavailable). Rows "
                "are synthetic. OWL verified (presentation/excerpts/slide22_future_candidate.omn).",
         cut="Without documentation, the name alone is a guess.",
         extra=["IF ASKED 'WHY NOT JUST ASK THE LLM WHAT BM_DEP MEANS?': from the name alone it is a guess. The data "
                "narrows it down: a 0/1 flag that switches on with cisplatin. That is evidence, and we keep "
                "provenance to both columns.",
                "IF ASKED ABOUT PRIVACY: only aggregate statistics (value ranges, co-occurrence) need to leave the "
                "database, not patient rows."]),
    dict(title="Thank you",
         say="To conclude. RIGOR anchors LLM generation in the database itself, grounds it in retrieved knowledge, "
             "and checks every step before the ontology grows. The result: meaning without losing structure, the "
             "link to the data intact, and ontologies that experts prefer. The paper and code are behind these QR "
             "codes. Thank you. I'm happy to take your questions.",
         cue="Finish speaking by 13:30. Leave this slide up during questions; backups follow it.",
         takeaway="Grounded, retrieved, validated.",
         source="Recap of the paper; contact address on the slide."),
]

CLOSING_TAKEAWAYS = "Meaning without losing structure · link to the data intact · preferred by experts"


def words(text):
    return len(re.sub(r"\[(pause|click)\]", "", text).split())


def seconds(text):
    return words(text) / WPM * 60 + text.count("[pause]") * PAUSE_S


def timeline():
    t, rows = 0.0, []
    for n, s in enumerate(SLIDES, 1):
        d = seconds(s["say"])
        rows.append((n, s, t, t + d, d))
        t += d
    return rows


def mmss(sec):
    sec = int(round(sec))
    return f"{sec // 60}:{sec % 60:02d}"


def notes_for(n, s, start, end):
    paras = [f"[{mmss(start)}–{mmss(end)}]  SPOKEN SCRIPT ({words(s['say'])} words)", s["say"], "",
             f"PRESENTER / POINTER: {s['cue']}", f"AUDIENCE TAKEAWAY: {s['takeaway']}",
             f"SOURCE AND SCOPE: {s['source']}"]
    if s.get("cut"):
        paras.append(f"IF RUNNING LATE, DROP: \"{s['cut']}\"")
    for extra in s.get("extra", []):
        paras += ["", extra]
    return paras


def apply(src, dst):
    for s in SLIDES:
        assert not s.get("cut") or s["cut"] in s["say"], s["title"]
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    pres = etree.parse(str(work / "ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in etree.parse(str(work / "ppt/_rels/presentation.xml.rels")).getroot()}
    order = [rels[s.get(f"{{{R}}}id")] for s in pres.find("p:sldIdLst", NS)]
    for (n, s, start, end, _), target in zip(timeline(), order):
        slide = work / "ppt" / target
        srels = etree.parse(str(slide.parent / "_rels" / f"{slide.name}.rels")).getroot()
        notes = next(r.get("Target") for r in srels if r.get("Type").endswith("/notesSlide"))
        set_notes(slide.parent / notes, notes_for(n, s, start, end))
        if n == 23:
            tree = etree.parse(str(slide))
            set_runs(tree, 1049, [(CLOSING_TAKEAWAYS, None, None)])
            for pic in tree.iter(f"{{{NS['p']}}}pic"):  # QR alt text (targets decoded from the images)
                cnv = pic.find(".//p:cNvPr", NS)
                cnv.set("descr", {"Picture 1051": "QR code: https://arxiv.org/abs/2506.01232 (paper)",
                                  "Picture 1053": "QR code: https://github.com/NadeenAhmad/RIGOR (code)"}
                        .get(cnv.get("name"), cnv.get("descr")))
            tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    total = timeline()[-1][3]
    print(f"wrote {dst}  ({sum(words(s['say']) for s in SLIDES)} words, ends at {mmss(total)})")


# ---- Documents ------------------------------------------------------------------------------

HARD_QUESTIONS = [
    ("The Judge-LLM is the same model as the generator. Isn't that self-evaluation?",
     "Inside the loop, yes: the same model generates and judges, so RIGOR needs no second model. Two things "
     "do not depend on it. The rule-based graph checks are deterministic. And the evaluation used a different "
     "model family, GPT-5.4, plus ten human experts, who ranked RIGOR first in 70% of cases."),
    ("Is the Non-Iterative baseline a strawman?",
     "It receives the same retrieved context, but its prompt asks for less (labels and comments). Its biggest "
     "failure is dropping columns, down to 6%, which RIGOR prevents by construction with the seed. A "
     "prompt-matched ablation is a fair next experiment."),
    ("Could the LLM judge have position bias?",
     "Fragments were anonymized, but the order was fixed (paper Fig. 7), so we do not rely on the LLM judge "
     "alone. The human study points the same way: 28 of 40 first places, with significant agreement. Future "
     "runs should randomize the order."),
    ("Why do critical P19 pitfalls remain (multiple domains or ranges)?",
     "They appear in 3 of 9 runs. A likely source is the union merge: when a property such as a patient link "
     "is reused across tables, each table adds its own domain. Merge-time checks and validator feedback inside "
     "the loop are the planned fix."),
    ("Why only schemas and documentation, not the data?",
     "Portability and privacy. Clinical data often cannot leave the hospital, while schemas and documentation "
     "can. Instance-level signals, computed as aggregates, are our next step."),
    ("Column coverage drops to 0.77. Why?",
     "That is DeepSeek on eICU. The seed is complete by construction, but enrichment can drop columns, and "
     "coverage is an embedding-based proxy. Stronger per-fragment coverage checks are planned."),
    ("Could the LLMs have seen these databases in training?",
     "eICU and Chinook are public. The liver-cancer registry is private, so no model saw it, and the LLM judge "
     "still ranked RIGOR first on all 16 of its tables."),
    ("Are the external ontologies actually used?",
     "Retrieved external concepts are passed as context, and alignment axioms are added only on a good match. "
     "In the ablation, full RIGOR has the highest external-ontology alignment (F1 0.059 vs 0.033 without "
     "retrieval), though absolute values are small."),
    ("How does the traversal handle self-referencing foreign keys (Chinook Employee.ReportsTo)?",
     "VERIFY IN THE CODE BEFORE THE TALK. Expected answer: a self-reference does not wait on another table, so "
     "it does not affect the order; it becomes an object property from the class to itself, which the prompt "
     "allows for meaningful recursive relations."),
    ("Why not compare with Gemini 2.5 Pro on BURR?",
     "In BURR it is a general-purpose LLM prompted to write the mapping directly: a mapping generator, not an "
     "ontology-learning system. Against it, RIGOR's class F1 is higher (.73 vs .65), relations are on par (.67 "
     "vs .68), attributes lower (.52 vs .56), and RIGOR also produces labels, hierarchy and provenance."),
]


def docs():
    rows = timeline()
    total_words = sum(words(s["say"]) for s in SLIDES)
    end = rows[-1][3]
    out = [
        "# RIGOR @ ISWC 2026: speech",
        "",
        f"Spoken script for the 23 main slides: **{total_words} words**, ending at **{mmss(end)}** at "
        f"{WPM} words per minute (the original deck's pace, clicks included) plus {PAUSE_S} s per marked "
        "[pause]. The slot is 15 minutes with Q&A separate, so this leaves a margin for a slower delivery.",
        "",
        "The same text is in each slide's speaker notes. This file is generated by "
        "`presentation/tools/speech.py`; edit the script there, not here.",
        "",
    ]
    for n, s, start, stop, dur in rows:
        out += [f"## {n}. {s['title']}  ·  {mmss(start)}–{mmss(stop)}  ·  {words(s['say'])} words", "",
                s["say"].replace("[pause]", "**[pause]**").replace("[click]", "*[click]*"), "",
                f"*Pointer:* {s['cue']}  ", f"*Takeaway:* {s['takeaway']}  ", f"*Source:* {s['source']}"]
        if s.get("cut"):
            out.append(f"  \n*If late, drop:* \"{s['cut']}\"")
        for extra in s.get("extra", []):
            out += ["", f"> {extra}"]
        out.append("")
    out += ["## Hard questions (backup slides B5–B6)", ""]
    for q, a in HARD_QUESTIONS:
        out += [f"**Q: {q}**  ", f"A: {a}", ""]
    (OUT_DIR / "SPEECH.md").write_text("\n".join(out), encoding="utf-8")

    checkpoints = [(6, "the method starts"), (12, "the architecture recap starts"), (19, "the results start"),
                   (22, "future work starts"), (23, "the close starts")]
    card = ["# Rehearsal card", "",
            f"Target: finish by **13:30**. Planned end: **{mmss(end)}** ({total_words} words).", "",
            "| Slide | Starts at | Ends at | Words |", "|---|---|---|---|"]
    for n, s, start, stop, _ in rows:
        card.append(f"| {n}. {s['title']} | {mmss(start)} | {mmss(stop)} | {words(s['say'])} |")
    card += ["", "## Checkpoints", "",
             "Glance at the clock when these slides come up. More than 30 s behind: apply the cuts below for "
             "the slides that follow.", ""]
    for n, label in checkpoints:
        card.append(f"- **Slide {n}** should come up at **{mmss(rows[n - 1][2])}** ({label}).")
    saved = 0
    card += ["", "## Cuts if running late", ""]
    for n, s, *_ in rows:
        if s.get("cut"):
            sec = seconds(s["cut"])
            saved += sec
            card.append(f"- Slide {n} ({s['title']}): drop \"{s['cut']}\" (saves ~{round(sec)} s)")
    card += ["", f"All cuts together save about {round(saved)} s.", "",
             "Slides 12–17 are quick clicks (about 5–10 s each). If very late, say only the first and last "
             "sentences and click through the middle four."]
    (OUT_DIR / "REHEARSAL_CARD.md").write_text("\n".join(card) + "\n", encoding="utf-8")
    print(f"wrote SPEECH.md and REHEARSAL_CARD.md ({total_words} words, ends at {mmss(end)})")


if __name__ == "__main__":
    {"apply": lambda: apply(sys.argv[2], sys.argv[3]), "docs": docs}[sys.argv[1]]()
