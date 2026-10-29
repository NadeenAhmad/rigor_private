#!/usr/bin/env python3
"""Speaker notes for the author's final deck (RIGOR_ISWC2026_4.pptx), one entry per reveal.

Each talk slide is revealed over one or more physical slides. Every physical slide gets only the
words spoken while it is on screen, so the notes never run ahead of the slide. Cue, takeaway,
source and late-cut notes sit on the reveal where they apply.

Timing model as before: 105 spoken words per minute (clicks included) plus 2 s per [pause].
Target: finish speaking by about 13:30 of the 15-minute slot.

Usage:
  python3 speech_final.py apply <in.pptx> <out.pptx>
  python3 speech_final.py docs            # write presentation/SPEECH.md and REHEARSAL_CARD.md
  python3 speech_final.py timing          # print the timeline
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from results import NS, R, set_notes  # noqa: E402

WPM, PAUSE_S = 105, 2
OUT_DIR = Path(__file__).resolve().parent.parent

# One entry per talk slide. `builds` holds one (say, cue) pair per physical slide of its reveal.
TALK = [
    dict(title="Title",
         builds=[("Good morning. I'm Nadeen Fathallah, from the University of Stuttgart. Relational databases run "
                  "hospitals, banks and governments. But their schemas tell us how data is stored, not what it "
                  "means. This talk is about RIGOR: turning those schemas into rich OWL ontologies, automatically.",
                  "Calm, slow start. Do not read the author list.")],
         takeaway="Schemas store data, not meaning.",
         source="Paper Section 1."),
    dict(title="Motivation",
         builds=[("Imagine you inherit this clinical database. An admission points to its patient through a "
                  "foreign key. But how should an ontology represent that link? What does BM_DEP mean? What "
                  "does an ECOG stage of three mean? The schema cannot tell us. Yet these meanings are exactly "
                  "what ontology-based access, knowledge graphs and integration need.",
                  "Point at the FK arrow, then the three questions on the right, top to bottom.")],
         takeaway="The schema knows storage and links, not domain meaning.",
         source="Illustrative table assembled from examples in paper Sections 3.2 and 3.4."),
    dict(title="What we want to build",
         builds=[("Here is what we want to build from this table.",
                  "Empty target. One breath, then click."),
                 ("First, from the schema alone: admission becomes a class with a real relation to PatientData, "
                  "and every property points back to its source column.",
                  "The class box and arrow, then the two 'source column' lines."),
                 ("The documentation adds human-readable labels: Bone Marrow Depression, ECOG Performance Status.",
                  "The two highlighted labels."),
                 ("A standard vocabulary links BM_DEP to bone marrow disease in the Disease Ontology.",
                  "The skos:relatedMatch line."),
                 ("And a comment says what the ECOG codes mean, from zero, fully active, to five, dead.",
                  "The rdfs:comment line. Do not mention LLMs yet."),
                 ("That is a usable semantic model. The challenge: add all this without losing a single column "
                  "or link.",
                  "The 'What this enables' box, then pause.")],
         takeaway="A usable model combines schema, documentation and standard vocabularies, with provenance.",
         source="Illustrative fragment, verified with the OWL API (presentation/excerpts/slide03_*.omn). "
                "obo:DOID_4961 = Disease Ontology 'bone marrow disease'; skos:relatedMatch, not exactMatch, "
                "because BM_DEP is a yes/no column, not the disease. Provenance is column-level (paper Section 3.2)."),
    dict(title="The research gap",
         builds=[("Existing work solves half of this. Schema-first rules, like direct mapping, preserve the "
                  "structure: every table, column and foreign key. But the domain meaning stays missing. LLMs "
                  "can suggest the meaning: labels, concepts, relations. But they omit elements or model them "
                  "incorrectly. So our question: can we keep schema fidelity and add domain semantics, across a "
                  "whole database?",
                  "Left card, right card, then land on the question.")],
         takeaway="Rules keep structure; LLMs suggest meaning; we need both.",
         source="Paper Sections 1-2."),
    dict(title="Our answer: RIGOR",
         builds=[("Our answer is RIGOR. It builds the ontology one table at a time, in six steps: order, seed, "
                  "retrieve, generate, validate, merge. Let's build our admission example together.",
                  "From 'what we have' to 'what we want', then sweep along the six steps.")],
         takeaway="Six steps, one table at a time.",
         source="Paper Section 3.1, Fig. 1."),
    dict(title="Step 1: order",
         builds=[("Step one: order. Admission points to PatientData, so we build PatientData first. In general, "
                  "RIGOR follows the foreign keys breadth-first: a table is processed once the tables it "
                  "references are done. So when we reach admission, PatientData already exists, and we reuse "
                  "it instead of inventing a duplicate.",
                  "Follow the FK arrow, then the processing order.")],
         takeaway="Parents first, so later tables reuse known classes.",
         source="Paper Sections 3.1-3.2 (FK-guided breadth-first order)."),
    dict(title="Step 2: seed",
         builds=[("Step two: seed. Before any LLM is involved, a deterministic direct mapping turns the table into "
                  "OWL. Admission becomes a class. The foreign key becomes an object property, and every "
                  "admission must have a patient. Each column becomes a data property with its XSD type. So "
                  "every column is in the seed, by construction. The LLM's job is to enrich, not to invent.",
                  "Read the glosses: 'each has a patient', 'the FK becomes a link'.")],
         takeaway="A deterministic seed guarantees coverage before the LLM starts.",
         source="Paper Section 3.2 (W3C Direct Mapping as specified by Sequeda et al., WWW 2012). The guarantee "
                "holds for the seed; final coverage is Result 2.",
         cut="The LLM's job is to enrich, not to invent."),
    dict(title="Step 3: retrieve",
         builds=[("The seed is faithful, but it does not say what BM_DEP means. So step three retrieves context "
                  "from three sources. Documentation defines cryptic names, like BM_DEP.",
                  "Documentation card."),
                 ("External ontologies, like the Disease Ontology, suggest established concepts.",
                  "External ontologies card."),
                 ("And the third source is the distinctive one: the ontology we have built so far. Earlier "
                  "tables become evidence for the next.",
                  "Growing core card; slow down here."),
                 ("One query per table, built from its name, columns and foreign-key targets, retrieves the top "
                  "matches from each source.",
                  "The query strip.")],
         takeaway="Three sources supply meaning; the growing core is unique to RIGOR.",
         source="Paper Section 3.3 (all-MiniLM-L6-v2, Faiss, k = 3 per source).",
         cut="One query per table, built from its name, columns and foreign-key targets, retrieves the top "
             "matches from each source."),
    dict(title="Step 4: generate",
         builds=[("Step four: generate. We prompt the Gen-LLM with the direct mapping from step two, the "
                  "retrieved context, and our enrichment instructions.",
                  "The subtitle, then the BEFORE panel."),
                 ("What comes back is a delta ontology: the same lines, with new ones filled into the gaps. The "
                  "tags show where each addition came from. The foreign key becomes the named relation "
                  "hasPatientData, and provenance links each property to its column: both from the schema. The "
                  "labels come from the documentation, the link to bone marrow disease from the Disease "
                  "Ontology, and the ECOG scale from the LLM itself.",
                  "Left to right on one row, then down the highlighted lines and their tags.")],
         takeaway="The LLM enriches a grounded draft, and every addition is traceable.",
         source="Paper Sections 3.2-3.4. Row-aligned BEFORE/AFTER, both verified with the OWL API."),
    dict(title="Step 5: validate",
         builds=[("Step five: validate. Plausible fragments can still break the ontology. A later table, "
                  "lab_result, reuses hasPatientData with a second domain. OWL would then infer that every lab "
                  "result linked to patient data is an admission. Rule-based graph checks catch this.",
                  "Structural card: the red line, then the explanation."),
                 ("Other errors are semantic. Here the model links admission to a new class, Patient, although "
                  "the core already has PatientData. Valid OWL, but now two classes mean the same thing. The "
                  "Judge-LLM sees the core and asks to reuse it.",
                  "Semantic card."),
                 ("Failed fragments go back to the Gen-LLM, with up to two retries. If one still fails, the "
                  "Judge-LLM picks the better of its last version and the plain direct mapping.",
                  "The correction strip.")],
         takeaway="Rules catch clashes in structure; the Judge catches clashes in meaning.",
         source="Paper Section 3.5 (Judge criterion 9: core ontology consistency; graph check: domain/range "
                "cardinality; retry budget k = 2). Two domains = OOPS! pitfall P19, still reported in 3 of 9 RIGOR "
                "runs (Table 7). Examples verified with the OWL API (presentation/excerpts/slide10_*.omn).",
         extra=["IF ASKED: P19 is exactly the error that sometimes still slips through (3 of 9 runs); putting "
                "the validators inside the generation loop is our next step (slide 19)."]),
    dict(title="Step 6: merge",
         builds=[("Step six: merge. The validated delta joins the core ontology. Because RIGOR reuses "
                  "identifiers, this is a simple union of axioms. And here is the loop: the expanded ontology "
                  "becomes a retrieval source for the next table. Each table builds on the ones before.",
                  "Core plus delta, then the loop strip.")],
         takeaway="The ontology is both the output and the next input.",
         source="Paper Section 3.6."),
    dict(title="Reassembling the architecture",
         builds=[("Let me put the pieces together, as in the paper's Figure 1. First, the schema decides the "
                  "order.", "Quick clicks from here. Reconnect, do not re-explain."),
                 ("Each table gets its deterministic seed.", "Direct mapping box."),
                 ("Three channels retrieve meaning: external ontologies, documentation, and the core built so "
                  "far.", "The three retrieval boxes."),
                 ("The Gen-LLM combines seed and context into a delta ontology.", "Gen-LLM."),
                 ("A Judge-LLM and rules validate it.", "Validate + refine."),
                 ("And the merged core feeds retrieval for the next table. That loop is RIGOR.",
                  "Trace the return arrow from (i) back to (e).")],
         takeaway="Order, seed, retrieve, generate, validate, merge, loop.",
         source="Paper Fig. 1 (a)-(i)."),
    dict(title="Evaluation setup",
         builds=[("Does this work beyond one example? We tested two medical databases: a private liver-cancer "
                  "registry from Leipzig, and the large public eICU intensive-care database.",
                  "Liver Cancer, then eICU-CRD. Do not read the numbers."),
                 ("And Chinook, a music store, as a second domain.", "Chinook card."),
                 ("We compare four methods: direct mapping; an LLM baseline that sees only the table; "
                  "Non-Iterative RAG, which adds retrieval in a single pass; and RIGOR.",
                  "The four method cards, left to right."),
                 ("The three LLM methods each run with three models, Claude, Mistral and DeepSeek, on the three "
                  "databases: twenty-seven ontologies.",
                  "The 27 line."),
                 ("Plus three direct mappings,", "The 3 line."),
                 ("thirty ontologies in total.", "The 30 line."),
                 ("We check them with automatic quality checks, an LLM judge, ten human experts, and the BURR "
                  "benchmark.", "The evaluation strip.")],
         takeaway="Three databases, three LLMs, four methods, four kinds of evidence.",
         source="Paper Sections 4.1-4.6. 30 ontologies = 3 LLM methods x 3 databases x 3 LLMs + 3 direct "
                "mappings. Models: Claude Opus 4.6, Mistral Small 24B, DeepSeek-Chat (paper footnotes 9-11)."),
    dict(title="Result 1: valid, fewer errors",
         builds=[("We go from the easiest check to the hardest. First, the basics: all thirty ontologies are "
                  "valid and logically consistent. Modeling errors are another story. Direct mapping averages "
                  "about five hundred OOPS! pitfalls per ontology, RIGOR eight. Same LLM, same database: Mistral "
                  "on eICU makes two hundred twenty with Non-Iterative RAG, two with RIGOR. [pause]",
                  "One breath for validity; then the averages; then the Mistral eICU column.")],
         takeaway="Every ontology is valid; RIGOR's have far fewer modeling errors, with every LLM.",
         source="Paper Sec. 4.6-4.7 (RDFLib, HermiT) and Table 7 (OOPS! totals per run). 72% of the other "
                "methods' pitfalls are P08, missing annotations.",
         extra=["IF ASKED ABOUT CRITICAL ERRORS: P19 (multiple domains or ranges) is the only pitfall OOPS! rates "
                "critical, and it is the second-domain error from the validation slide. Per ontology: Direct "
                "Mapping 15, LLM baseline 17, Non-Iterative RAG 5.4, RIGOR 2.8; zero in 6 of RIGOR's 9 "
                "ontologies (Table 7). RIGOR's prompt allows one domain and one range, and the graph check "
                "rejects a second."]),
    dict(title="Result 2: keeps the schema",
         builds=[("Fewer errors, but is anything lost? On average, RIGOR covers ninety-three percent of the "
                  "columns, as much as direct mapping. Non-Iterative RAG, with the same retrieved context but "
                  "no seed, covers only a third.",
                  "RIGOR's average next to Direct Mapping's, then the Non-Iterative RAG average."),
                 ("And it swings with the LLM, from six to seventy-eight percent. With RIGOR, every LLM keeps at "
                  "least seventy-seven percent, in every database. [pause]",
                  "The outlined Non-Iterative RAG row, then the strip.")],
         takeaway="RIGOR keeps the schema as well as Direct Mapping, with every LLM.",
         source="Paper Table 9 (column coverage per run; embedding similarity >= 0.55, all-MiniLM-L6-v2).",
         extra=["IF ASKED WHY Non-Iterative RAG DROPS BUT THE BASELINE DOES NOT: the baseline only translates "
                "the table, so it keeps column names. Non-Iterative RAG does everything in one pass with a lot "
                "of retrieved context and silently drops properties (e.g. 6 data properties for 350 Liver Cancer "
                "columns with Mistral, Table 8); nothing checks coverage. RIGOR starts from a seed with every "
                "column, works one table at a time, and its validator rejects fragments that drop a column.",
                "IF ASKED ABOUT MODEL SIZE AND COST: with RIGOR, Mistral Small 24B reaches 98% coverage (Claude "
                "95%) for about $0.12 across the three databases, against about $33 with Claude (Tables 9, 15)."]),
    dict(title="Result 3: adds meaning",
         builds=[("Nothing is lost. Is meaning added? Direct mapping and the LLM baseline add no labels, "
                  "provenance or hierarchy, in any run. Non-Iterative RAG adds labels, but no provenance and no "
                  "hierarchy. Only RIGOR adds all three, with every LLM, in every database. [pause]",
                  "The red band first, then the three RIGOR rows.")],
         takeaway="Only RIGOR adds labels, provenance and a class hierarchy, with every LLM.",
         source="Paper Table 8. Hierarchy as inheritance richness (subclass axioms per class; OntoQA, Tartir et "
                "al. 2005; OntoMetrics, Lantow 2016), computed from Table 8 counts. Shading is within each database.",
         cut="Non-Iterative RAG adds labels, but no provenance and no hierarchy.",
         extra=["IF ASKED 'DOES CHINOOK NEED SO MANY SUBCLASSES?': count is not quality, so we report whether a "
                "hierarchy exists, and only RIGOR builds one. RIGOR's Chinook ontologies have zero OOPS! pitfalls "
                "and the judge ranks RIGOR first on all 11 Chinook tables; it does score RIGOR's conciseness "
                "lower there (2.59 vs 3.32 for the baseline, Table 10), so tighter hierarchies are future work.",
                "DOMAIN PATTERN (notes only): inheritance richness is 0.10-0.70 on the clinical schemas and "
                "2.1-4.1 on the music store; the LLM proposes more specializations in a familiar domain."]),
    dict(title="Result 4: BURR",
         builds=[("Meaning is added. But is the link back to the data still correct? BURR, an external benchmark, "
                  "checks exactly this: every class, relation and attribute must map back to the right table, "
                  "foreign key or column, scored with F1 against gold mappings. On BURR's ISWC scenario, RIGOR has "
                  "the highest F1 on all three, ahead of rule-based and LLM-based systems. Against the LLM-based "
                  "OntoGenix, relations nearly double, because foreign keys become named relations, like "
                  "hasPatientData.",
                  "Left card first, then land on the RIGOR row.")],
         takeaway="Semantic enrichment did not break the correspondence that OBDA and KG population need.",
         source="Paper Table 13. BURR = Laskowski et al., PACMMOD 2025; matching rule from BURR Section 4, Eq. 3. "
                "Left-card element names are illustrative. Relations: RIGOR 0.67, OntoGenix 0.35, best "
                "rule-based (RDB2Onto) 0.62.",
         cut="scored with F1 against gold mappings",
         extra=["DO NOT SAY: 'BURR found LLMs trail rule-based systems'. BURR's results show Gemini 2.5 Pro ahead "
                "of the rule-based systems overall (rule-based systems win only on NPD).",
                "IF ASKED ABOUT GEMINI 2.5 PRO (BURR Table 3, ISWC .65 / .68 / .56): in BURR it is a "
                "general-purpose LLM prompted with few-shot examples and a checklist to write the D2RQ mapping "
                "directly. It is a mapping generator, not an ontology-learning system, so we compare with "
                "ontology-learning systems. Against it, RIGOR's class F1 is higher, relations are on par, "
                "attributes lower, and RIGOR also produces labels, hierarchy and provenance."]),
    dict(title="Result 5: LLM judge and experts",
         builds=[("Automatic checks agree. Do evaluators agree? For each table, an LLM judge and ten ontology "
                  "experts saw the schema, five competency questions, the questions the ontology must answer, "
                  "and four anonymized ontologies. They scored each on seven dimensions and ranked them first to "
                  "fourth. The LLM judge put RIGOR first in every ranking. [pause] The experts put RIGOR first in "
                  "seventy percent of their rankings, with the highest average score. [pause] Direct mapping "
                  "gives structure without meaning. LLMs alone give meaning without structure. RIGOR gives both.",
                  "The 'How they judged' box in one breath; then the two headlines.")],
         takeaway="Judged blind against competency questions, an LLM judge and human experts both rank RIGOR first.",
         source="Paper Sec. 4.2 and 4.6 (5 expert-verified CQs per table, 7 dimensions, ranking), Table 1 (judge "
                "scores and first places: 58 of 58 tables), Table 11 (expert scores), Table 12 (Kendall's W "
                "0.39-0.72). Judge evaluated the Claude ontologies; experts 4 medical tables, 28 of 40 first places.",
         cut="with the highest average score",
         extra=["IF ASKED HOW MANY: the judge ranked RIGOR first on 58 of 58 tables (16 Liver Cancer, 31 eICU-CRD, "
                "11 Chinook); the experts in 28 of 40 rankings (10 experts x 4 tables)."]),
    dict(title="Future work: validators in the loop",
         builds=[("Where next? First, the validators. RDFLib, the reasoner and OOPS! only checked our final "
                  "ontologies; NeOn-GPT first proposed this stack of checks for evaluation. Next, they flag each "
                  "error while RIGOR builds, and their feedback prompts the LLM to repair the ontology.",
                  "The sentence on the slide, then click."),
                 ("We have realized this loop in the RIGOR App, presented in yesterday's demo session. The code "
                  "and a demo video are behind these QR codes.",
                  "Screenshot, then the two QR codes.")],
         takeaway="The evaluation checks move inside generation; the RIGOR App already does this.",
         source="Evaluation stack from NeOn-GPT (Fathallah et al., ESWC 2024; paper ref. 10), extended in "
                "Extended NeOn-GPT (Semantic Web 17(4), 2026; ref. 9). RIGOR App: Fathallah, Nayyeri & Staab, "
                "ISWC 2026 Demos."),
    dict(title="Future work: let the data speak",
         builds=[("Second, let the data speak. Today RIGOR reads schemas and documentation, but instance data can "
                  "propose new knowledge. Here, the treatment table says which admissions received cisplatin, "
                  "and the admission table records bone marrow depression.",
                  "The two tables; A1 and A3 in both."),
                 ("Joined on the admission id, every cisplatin admission has bone marrow depression. That link is "
                  "not in the documentation, so RIGOR can propose it as candidate knowledge, with provenance to "
                  "both columns. A pattern, not a cause: an expert reviews it before it joins the ontology.",
                  "The pattern bars, then the candidate code and its badge.")],
         takeaway="Instance data can propose new, reviewable knowledge that the documentation does not contain.",
         source="Rows are synthetic. OWL verified (presentation/excerpts/slide22_future_candidate.omn): a typed "
                "individual with provenance, no causal axiom.",
         extra=["IF ASKED WHY NOT AN OWL CAUSAL AXIOM: OWL reasoning is deductive and open-world; it cannot derive "
                "causation from statistics, and 'Cisplatin SubClassOf causes some BoneMarrowDepression' would "
                "wrongly claim it for every case. Other treatments or conditions could explain the pattern.",
                "IF ASKED ABOUT PRIVACY: only aggregate statistics (co-occurrence counts) need to leave the "
                "database, not patient rows."]),
    dict(title="Thank you",
         builds=[("To conclude. RIGOR anchors LLM generation in the database itself, grounds it in retrieved "
                  "knowledge, and checks every step before the ontology grows. The result: meaning without "
                  "losing structure, the link to the data intact, and ontologies that experts prefer. The paper "
                  "and code are behind these QR codes. Thank you. I'm happy to take your questions.",
                  "Finish speaking by 13:30. Leave this slide up during questions.")],
         takeaway="Grounded, retrieved, validated.",
         source="Recap of the paper; contact address on the slide."),
]


def words(text):
    return len(re.sub(r"\[(pause|click)\]", "", text).split())


def seconds(text):
    return words(text) / WPM * 60 + text.count("[pause]") * PAUSE_S


def mmss(sec):
    sec = int(round(sec))
    return f"{sec // 60}:{sec % 60:02d}"


def physical():
    """One row per physical slide: (talk number, build index, build count, entry, say, cue, start, end)."""
    t, rows = 0.0, []
    for n, s in enumerate(TALK, 1):
        for k, (say, cue) in enumerate(s["builds"], 1):
            d = seconds(say)
            rows.append((n, k, len(s["builds"]), s, say, cue, t, t + d))
            t += d
    return rows


def notes(n, k, m, s, say, cue, start, end):
    head = f"[{mmss(start)}–{mmss(end)}]  SLIDE {n}" + (f" · REVEAL {k} OF {m}" if m > 1 else "")
    paras = [f"{head}  ({words(say)} words)", say, "", f"PRESENTER / POINTER: {cue}"]
    if s.get("cut") and s["cut"] in say:
        paras.append(f"IF RUNNING LATE, DROP: \"{s['cut']}\"")
    if k == m:
        paras += [f"AUDIENCE TAKEAWAY: {s['takeaway']}", f"SOURCE AND SCOPE: {s['source']}"]
        for extra in s.get("extra", []):
            paras += ["", extra]
    return paras


def slide_order(work):
    pres = etree.parse(str(work / "ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in etree.parse(str(work / "ppt/_rels/presentation.xml.rels")).getroot()}
    return [work / "ppt" / rels[s.get(f"{{{R}}}id")] for s in pres.find("p:sldIdLst", NS)]


def apply_notes(work):
    rows = physical()
    order = slide_order(work)
    assert len(rows) == len(order), (len(rows), len(order))
    for row, slide in zip(rows, order):
        srels = etree.parse(str(slide.parent / "_rels" / f"{slide.name}.rels")).getroot()
        target = next(r.get("Target") for r in srels if r.get("Type").endswith("/notesSlide"))
        set_notes(slide.parent / target, notes(*row))


def talk_numbers():
    """Talk-slide number for every physical slide: a reveal keeps one number."""
    return [n for n, *_ in physical()]


def repack(work, dst):
    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())


def apply(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    apply_notes(work)
    repack(work, dst)
    shutil.rmtree(work)
    summary()


def summary():
    rows = physical()
    total = sum(words(r[4]) for r in rows)
    print(f"{total} words, ends at {mmss(rows[-1][7])}")


def timing():
    for n, s in enumerate(TALK, 1):
        rows = [r for r in physical() if r[0] == n]
        w = sum(words(r[4]) for r in rows)
        print(f"{n:2}  {mmss(rows[0][6]):>5}–{mmss(rows[-1][7]):<5} {w:4}w  {len(rows)} reveal(s)  {s['title']}")
    summary()


def docs():
    rows = physical()
    lines = ["# RIGOR talk: spoken script (final deck)", "",
             f"Timing model: {WPM} words per minute (clicks included) plus {PAUSE_S} s per [pause]. "
             f"Total: {sum(words(r[4]) for r in rows)} words, ending at {mmss(rows[-1][7])}. "
             "Each physical slide's notes hold only what is said while it is on screen.", ""]
    for n, s in enumerate(TALK, 1):
        mine = [r for r in rows if r[0] == n]
        lines += [f"## {n}. {s['title']}  ({mmss(mine[0][6])}–{mmss(mine[-1][7])})", ""]
        for _, k, m, _, say, cue, *_ in mine:
            prefix = f"**Reveal {k}/{m}.** " if m > 1 else ""
            lines += [prefix + say, "", f"> {cue}", ""]
    (OUT_DIR / "SPEECH.md").write_text("\n".join(lines))
    card = ["# Rehearsal card", "", "| # | Slide | Starts | Takeaway |", "|---|---|---|---|"]
    for n, s in enumerate(TALK, 1):
        start = next(r[6] for r in rows if r[0] == n)
        card.append(f"| {n} | {s['title']} | {mmss(start)} | {s['takeaway']} |")
    card += ["", f"Finish by {mmss(rows[-1][7])}; the slot is 15:00."]
    (OUT_DIR / "REHEARSAL_CARD.md").write_text("\n".join(card) + "\n")
    print("wrote SPEECH.md and REHEARSAL_CARD.md")


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "apply":
        apply(sys.argv[2], sys.argv[3])
    elif cmd == "docs":
        docs()
    else:
        timing()
