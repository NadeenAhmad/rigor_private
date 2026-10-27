#!/usr/bin/env python3
"""Phase 5: turn the limitations slide (22) into future work, "from schemas to instances".

Semantic-web framing: co-occurrence in instance data plus documentation yields a CANDIDATE,
provenance-annotated association for expert review, never a causal axiom (OWL reasoning is
deductive and open-world; it cannot infer causation from statistics). The OWL excerpt is
FUTURE in code_excerpts.py and is verified there with the OWL API.

Usage: python3 future_work.py <in.pptx> <out.pptx>
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import FUTURE, code_block  # noqa: E402
from results import (BODY, EMERALD, INK, MUTED, NS, TEAL, Ids, add, delete, run, set_notes,  # noqa: E402
                     set_runs, shape)

RED = "B23A1E"


def mono_rows(tree, ids, x, y, w, rows, highlight):
    """A small data table: header + rows, Consolas, tab-aligned; highlighted rows tinted emerald."""
    tabs = '<a:tabLst><a:tab pos="777240" algn="l"/><a:tab pos="1737360" algn="l"/></a:tabLst>'
    for i, cells in enumerate(rows):
        header = i == 0
        size = 950 if header else 1150
        color = MUTED if header else INK
        text = "\t".join(cells)
        el = shape(ids(), f"Instance row {i}", x, y + i * 0.30, w, 0.27, run(text, color, size, header, "Consolas"),
                   geom="roundRect", fill="E6F2EC" if i in highlight else None, anchor="ctr", inset=0.10)
        el.find(".//a:pPr", NS).append(etree.fromstring(f'<x xmlns:a="{NS["a"]}">{tabs}</x>')[0])
        add(tree, el)


def edit_slide22(tree):
    ids = Ids(tree)
    set_runs(tree, 1013, [("FUTURE WORK  /  FROM SCHEMAS TO INSTANCES", None, None)])
    set_runs(tree, 1014, [("Next: let the data speak.", None, None)])
    set_runs(tree, 1015, [("Today RIGOR reads schemas and documentation. Instance data could propose new "
                           "knowledge, for experts to review.", None, None)])
    delete(tree, *range(1016, 1031))

    top, h = 2.40, 3.24
    cards = [(0.90, 3.30), (4.62, 3.05), (8.09, 4.31)]
    for i, (x, w) in enumerate(cards):
        add(tree, shape(ids(), f"Future card {i + 1}", x, top, w, h, geom="roundRect", fill="FFFFFF",
                        line="D9DEDB"))
    for x in (4.24, 7.71):
        add(tree, shape(ids(), "Flow arrow", x, 3.78, 0.34, 0.42, geom="rightArrow", fill=EMERALD))

    # 1. Instances (synthetic).
    x, w = cards[0]
    add(tree, shape(ids(), "Card 1 label", x + 0.22, top + 0.18, w - 0.4, 0.30,
                    run("1  INSTANCES", TEAL, 1300, True, "Consolas")))
    add(tree, shape(ids(), "Card 1 sub", x + 0.22, top + 0.50, w - 0.4, 0.28, run("What patient rows show", MUTED, 1300)))
    mono_rows(tree, ids, x + 0.15, top + 0.90, w - 0.30,
              [("PATIENT", "CISPLATIN", "BM_DEP AFTER"), ("P01", "yes", "1"), ("P02", "yes", "1"),
               ("P03", "no", "0"), ("P04", "yes", "1"), ("P05", "no", "0")], highlight={1, 2, 4})
    add(tree, shape(ids(), "Card 1 note", x + 0.22, top + 2.78, w - 0.4, 0.42,
                    run("Synthetic illustration. Rows also reveal codes: BM_DEP is only 0 or 1.", MUTED, 1100)))

    # 2. Documentation.
    x, w = cards[1]
    add(tree, shape(ids(), "Card 2 label", x + 0.22, top + 0.18, w - 0.4, 0.30,
                    run("2  + DOCUMENTATION", TEAL, 1300, True, "Consolas")))
    add(tree, shape(ids(), "Card 2 sub", x + 0.22, top + 0.50, w - 0.4, 0.28, run("What the drug label says", MUTED, 1300)))
    quote = shape(ids(), "Drug label quote", x + 0.18, top + 0.90, w - 0.36, 1.22,
                  run("Myelosuppression (bone-marrow suppression) is a documented adverse reaction of cisplatin.",
                      INK, 1400), geom="roundRect", fill="F3F5F4", anchor="ctr", inset=0.14)
    quote.find(".//a:rPr", NS).set("i", "1")
    add(tree, quote)
    add(tree, shape(ids(), "Quote source", x + 0.22, top + 2.16, w - 0.4, 0.22,
                    run("Cisplatin drug label, paraphrased", MUTED, 1000)))
    add(tree, shape(ids(), "Card 2 note", x + 0.22, top + 2.50, w - 0.4, 0.60,
                    run("Timing is evidence too: BM_DEP is recorded after the first cycle.", BODY, 1200)))

    # 3. Candidate knowledge (verified OWL).
    x, w = cards[2]
    add(tree, shape(ids(), "Card 3 label", x + 0.22, top + 0.18, w - 0.4, 0.30,
                    run("3  CANDIDATE KNOWLEDGE", EMERALD, 1300, True, "Consolas")))
    add(tree, shape(ids(), "Card 3 sub", x + 0.22, top + 0.50, w - 0.4, 0.28,
                    run("A reviewable hypothesis, not a fact", MUTED, 1300)))
    lines = len(FUTURE.splitlines())
    add(tree, code_block(ids(), x + 0.15, top + 0.84, w - 0.30, lines * 1250 / 100 / 72 + 0.16, FUTURE, 1100, 1250,
                         inset=0.10))
    add(tree, shape(ids(), "Card 3 note", x + 0.22, top + 2.66, w - 0.4, 0.52,
                    run("No causes axiom. ", RED, 1100, True)
                    + run("Co-administered gemcitabine, radiotherapy or cirrhosis can also suppress the marrow.",
                          BODY, 1100)))

    # Other next steps (paper Section 5).
    add(tree, shape(ids(), "Also next strip", 0.90, 5.84, 11.50, 0.46,
                    run("ALSO NEXT   ", EMERALD, 1250, True)
                    + run("Validator and reasoner feedback inside the loop, to catch residual P19  ·  "
                          "aggregate statistics only, so patient data stays in the hospital", INK, 1250),
                    geom="roundRect", fill="E6F2EC", anchor="ctr", inset=0.20))
    add(tree, shape(ids(), "Future caption", 0.93, 6.42, 11.47, 0.26,
                    run("Instance rows are synthetic. Future directions follow paper Section 5 (instance-level "
                        "signals, tool feedback, symbolic reasoning).", MUTED, 1050)))


NOTES = [
    "[11:55–12:55]  SPOKEN SCRIPT (interim; final wording in Phase 6)",
    "Where next? Today RIGOR reads schemas and documentation. The next step is to let the data speak. "
    "Imagine several patients whose records show bone marrow depression after cisplatin cycles. The rows "
    "alone show a pattern; the drug label tells us bone-marrow suppression is a known side effect of "
    "cisplatin. Together, they let RIGOR propose new knowledge: a candidate association between cisplatin "
    "and bone marrow depression, with provenance to both sources. Note the word candidate. Co-occurrence is "
    "not causation: gemcitabine, radiotherapy or cirrhosis can also suppress the marrow. So the ontology "
    "records a reviewable hypothesis, not a causal axiom, for an expert to confirm. Instance data also helps "
    "right away: it shows how BM_DEP is coded. And we want validator and reasoner feedback inside the loop.",
    "",
    "PRESENTER / POINTER: Left to right: rows, label, code. Say 'candidate' or 'may have contributed', never "
    "'caused'.",
    "AUDIENCE TAKEAWAY: Instances plus documentation can propose knowledge; experts confirm it.",
    "SOURCE AND SCOPE: Paper Section 5 (instance-level signals, OOPS!/tool feedback, symbolic reasoning). "
    "Rows are synthetic. The OWL excerpt is verified (presentation/excerpts/slide22_future_candidate.omn): "
    "a typed individual with facts and provenance, no causal axiom. If asked why not OWL causation: OWL "
    "reasoning is deductive and open-world; it cannot derive causation from statistics, and "
    "'Cisplatin SubClassOf causes some BoneMarrowDepression' would wrongly claim it for every case.",
]


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    slide = work / "ppt/slides/slide22.xml"
    tree = etree.parse(str(slide))
    edit_slide22(tree)
    tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
    rel = etree.parse(str(work / "ppt/slides/_rels/slide22.xml.rels")).getroot()
    target = next(r.get("Target") for r in rel if r.get("Type").endswith("/notesSlide"))
    set_notes(work / "ppt/slides" / target, NOTES)
    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print("wrote", dst)


if __name__ == "__main__":
    main(*sys.argv[1:3])
