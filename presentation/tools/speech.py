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
    dict(title="Result 1: meaning and modeling errors",
         say="First result: what do the ontologies contain, and how many modeling errors do they have? On the "
             "left, what each method adds. Direct mapping and the LLM baseline add no labels, no provenance and "
             "no hierarchy. Single-pass retrieval adds labels, nothing more. Only RIGOR adds all three. [pause] On "
             "the right, the modeling errors that OOPS! finds, for the ontologies generated with Claude. On eICU, "
             "direct mapping has six hundred sixty-five, RIGOR has two. On Chinook, RIGOR has none. Mistral and "
             "DeepSeek show the same pattern. [pause]",
         cue="Left table top to bottom, land on the RIGOR row; then the eICU bars, then Chinook's zero.",
         takeaway="Only RIGOR adds labels, provenance and hierarchy, with far fewer modeling errors.",
         source="Paper Tables 7 (OOPS! totals, Claude) and 8 (annotations, provenance, SubClassOf; all 9 runs per "
                "method).",
         cut="Mistral and DeepSeek show the same pattern."),
    dict(title="Result 2: BURR",
         say="Second: is the link back to the data real? We checked with BURR, an external benchmark for ontology "
             "learning from relational databases. BURR asks every system for a mapping: which table becomes a "
             "class, which foreign key a relation, which column an attribute. Two elements match when they map to "
             "the same database elements, whatever they are called. BURR scores this against a gold mapping, with "
             "F1 per element type. RIGOR's provenance gives us this mapping directly; we export it as D2RQ. On "
             "BURR's ISWC scenario, RIGOR has the highest F1 for classes, relations and attributes among the "
             "compared ontology-learning systems, including both rule-based ones. So enrichment did not cost us "
             "fidelity.",
         cue="Left card first, then land on the RIGOR row.",
         takeaway="Semantic enrichment did not break the correspondence that OBDA and KG population need.",
         source="Paper Table 13. BURR = Laskowski et al., PACMMOD 2025; matching rule from BURR Section 4, Eq. 3. "
                "Left-card element names are illustrative.",
         cut="Two elements match when they map to the same database elements, whatever they are called.",
         extra=["DO NOT SAY: 'BURR found LLMs trail rule-based systems'. BURR's results show Gemini 2.5 Pro ahead "
                "of the rule-based systems overall (rule-based systems win only on NPD).",
                "IF ASKED ABOUT GEMINI 2.5 PRO (BURR Table 3, ISWC .65 / .68 / .56): in BURR it is a "
                "general-purpose LLM prompted with few-shot examples and a checklist to write the D2RQ mapping "
                "directly. It is a mapping generator, not an ontology-learning system that builds an enriched "
                "ontology, so we compare with ontology-learning systems. Against it, RIGOR's class F1 is higher, "
                "relations are on par, attributes lower, and RIGOR also produces labels, hierarchy and "
                "provenance. It is prompted, not fine-tuned."]),
    dict(title="Result 3: LLM judge and experts",
         say="Finally: do evaluators prefer RIGOR? Both evaluations were blind: four anonymized ontologies per "
             "table, ranked first to fourth. An LLM judge, GPT-5.4, ranked RIGOR first on all fifty-eight "
             "tables, across the three databases. [pause] Ten ontology experts ranked RIGOR first in twenty-eight "
             "of forty rankings, seventy percent, and they agreed significantly with each other. Machine and human "
             "evaluators reach the same verdict.",
         cue="Left number first, then the right number and its bars.",
         takeaway="An independent LLM judge and human experts both rank RIGOR first.",
         source="Paper Table 1 (rankings) and Table 12 (Kendall's W 0.39-0.72). LLM judge: Claude outputs, 290 "
                "competency questions. Experts: 4 tables from the two medical databases.",
         cut="and they agreed significantly with each other"),
    dict(title="Future work",
         say="Where next? Today RIGOR reads schemas and documentation. The next step is to let the data speak. "
             "Imagine several patients whose records show bone marrow depression after cisplatin cycles. The rows "
             "alone show a pattern. The drug label adds that bone-marrow suppression is a known adverse reaction "
             "of cisplatin. Together, they let RIGOR propose new knowledge: a candidate association between "
             "cisplatin and bone marrow depression, with provenance to both sources. Candidate, not cause. "
             "Gemcitabine, radiotherapy or cirrhosis could also explain it. So an expert reviews it. We also want "
             "validator and reasoner feedback inside the loop.",
         cue="Rows, label, code, left to right. Say 'candidate' or 'may have contributed', never 'caused'.",
         takeaway="Instances plus documentation can propose knowledge; experts confirm it.",
         source="Paper Section 5. Rows are synthetic. OWL verified "
                "(presentation/excerpts/slide22_future_candidate.omn): a typed individual with provenance, no "
                "causal axiom.",
         cut="We also want validator and reasoner feedback inside the loop.",
         extra=["IF ASKED WHY NOT AN OWL CAUSAL AXIOM: OWL reasoning is deductive and open-world; it cannot derive "
                "causation from statistics, and 'Cisplatin SubClassOf causes some BoneMarrowDepression' would "
                "wrongly claim it for every case."]),
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
