#!/usr/bin/env python3
"""Results as a ladder of six questions, easiest check first, on the author's edited deck.

    1 VALID  2 KEEPS THE SCHEMA  3 ADDS MEANING  4 FEWER ERRORS  5 GOLD MAPPINGS  6 JUDGES

  A  RESULT 1  rungs 1-2  validity (Sec. 4.7) and column coverage (Table 9), averaged per ontology
  B  RESULT 2  rungs 3-4  added meaning (Table 8) and OOPS! pitfalls (Table 7), averaged per ontology
  C  RESULT 3  rung 5     the author's BURR slide, shifted down for the tracker, plus the relations line
  D  RESULT 4  rung 6     LLM judge and human experts (Tables 1, 10, 12), plus the closing sentence

Every number is an average over all 30 ontologies (each method's own runs), so no LLM breakdown is
needed. A is a copy of the author's Result 1 slide; the output deck holds only these four slides,
with speaker notes from speech.py.

Usage: python3 results_v4.py <in.pptx> <out.pptx>
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from arch_steps import PPTX_SCRIPTS  # noqa: E402
from code_excerpts import run_xml  # noqa: E402
from results import BODY, EMERALD, INK, MUTED, NS, P, RED, TEAL, Ids, add, set_notes, shape  # noqa: E402
from slide9_diff import geometry, text as sp_text  # noqa: E402
from speech import notes_for, timeline  # noqa: E402

LINE, PANEL, MINT, STRIP = "D9DEDB", "F3F5F4", "E8F6EF", "E6F2EC"
EMU = 914400
HEADER = {"eyebrow": (0.88, 0.56, 11.76, 0.34), "title": (0.88, 1.00, 11.77, 0.73),
          "subtitle": (0.90, 1.80, 11.68, 0.36)}
LADDER = ("VALID", "KEEPS THE SCHEMA", "ADDS MEANING", "FEWER ERRORS", "GOLD MAPPINGS", "JUDGES")
TOP = 2.85  # content starts below the ladder

METHODS = ("Direct Mapping", "LLM baseline", "Non-Iterative", "RIGOR")
COVERAGE = (91, 74, 34, 93)            # Table 9: mean column coverage per ontology, %
ADDED = {"Labels & comments": (0, 0, 389, 777),   # Table 8: mean per ontology
         "Provenance links": (0, 0, 0, 196),
         "Subclass axioms": (0, 0, 0, 59)}
PITFALLS = (499, 307, 52, 8)           # Table 7: mean OOPS! pitfalls per ontology
JUDGE = (("Liver Cancer", 16), ("eICU-CRD", 31), ("Chinook", 11))
EXPERT_FIRST = (5, 5, 2, 28)           # Table 1: first places of 40


def r(t, color=INK, size=1400, bold=False, font="Calibri", italic=False):
    return run_xml(t, color, bold, size, font=font, italic=italic)


def text(ids, x, y, w, h, runs, align="l", anchor="t", name="Text"):
    return shape(ids(), name, x, y, w, h, runs, align=align, anchor=anchor)


def place(sp, x=None, y=None, w=None, h=None):
    off, ext = sp.find(".//a:xfrm/a:off", NS), sp.find(".//a:xfrm/a:ext", NS)
    for el, key, v in ((off, "x", x), (off, "y", y), (ext, "cx", w), (ext, "cy", h)):
        if v is not None:
            el.set(key, str(int(v * EMU)))


def set_text(sp, value):
    ts = list(sp.iter(f"{{{NS['a']}}}t"))
    ts[0].text = value
    for t in ts[1:]:
        t.getparent().getparent().remove(t.getparent())


def header_shapes(tree):
    root = tree.find(".//p:spTree", NS)
    shapes = [el for el in root if el.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp", f"{{{P}}}pic", f"{{{P}}}grpSp",
                                              f"{{{P}}}graphicFrame")]
    eyebrow = next(sp for sp in shapes if sp_text(sp).startswith("RESULT"))
    title = shapes[shapes.index(eyebrow) + 1]
    return root, shapes, eyebrow, title


def reset(tree, eyebrow_text, title_text, subtitle_text):
    """Keep eyebrow and title (the author's styling), drop everything else, add subtitle."""
    root, shapes, eyebrow, title = header_shapes(tree)
    for sp in shapes:
        if sp is not eyebrow and sp is not title:
            root.remove(sp)
    for sp, value, key in ((eyebrow, eyebrow_text, "eyebrow"), (title, title_text, "title")):
        set_text(sp, value)
        place(sp, *HEADER[key])
    ids = Ids(tree)
    add(tree, text(ids, *HEADER["subtitle"], r(subtitle_text, BODY, 1600), name="Subtitle"))
    return ids


def ladder(tree, ids, current):
    """Six-rung tracker under the subtitle: done = mint, current = emerald, upcoming = outlined."""
    x, y, w, gap = 0.88, 2.30, 1.86, 0.10
    for k, label in enumerate(LADDER, 1):
        if k in current:
            fill, line, color = EMERALD, None, "FFFFFF"
        elif k < min(current):
            fill, line, color = MINT, None, EMERALD
        else:
            fill, line, color = "FFFFFF", LINE, "6B747B"
        add(tree, shape(ids(), "Ladder", x + (k - 1) * (w + gap), y, w, 0.34,
                        r(f"{k}  {label}", color, 1000, True, "Consolas"), fill=fill, line=line,
                        align="ctr", anchor="ctr"))


def card(tree, ids, x, y, w, h, label, note="", fill="FFFFFF"):
    add(tree, shape(ids(), "Card", x, y, w, h, geom="roundRect", fill=fill, line=LINE))
    add(tree, text(ids, x + 0.3, y + 0.20, w - 0.6, 0.3, r(label, TEAL, 1300, True, "Consolas"), name="Card label"))
    if note:
        add(tree, text(ids, x + 0.3, y + 0.52, w - 0.6, 0.26, r(note, MUTED, 1150, italic=True), name="Card note"))


def strip(tree, ids, y, runs):
    add(tree, shape(ids(), "Strip", 0.88, y, 11.57, 0.46, runs, geom="roundRect", fill=STRIP, anchor="ctr",
                    inset=0.22))


def source(tree, ids, line, y=6.98):
    add(tree, text(ids, 0.90, y, 11.6, 0.26, r(line, MUTED, 1000), name="Source"))


def bars(tree, ids, x, y, w, values, unit, maximum, colors, pitch=0.48, label_w=1.75):
    """Horizontal bars, one row per method, value label after the bar."""
    track = w - label_w - 0.85
    for k, (method, v) in enumerate(zip(METHODS, values)):
        ry = y + k * pitch
        rigor = method == "RIGOR"
        add(tree, text(ids, x, ry, label_w, 0.32, r(method, EMERALD if rigor else INK, 1400, rigor), anchor="ctr",
                       name="Method"))
        add(tree, shape(ids(), "Track", x + label_w, ry + 0.08, track, 0.16, fill="EEF1EF"))
        if v:
            add(tree, shape(ids(), "Bar", x + label_w, ry + 0.08, max(track * v / maximum, 0.03), 0.16,
                            fill=colors[k]))
        add(tree, text(ids, x + label_w + track + 0.08, ry, 0.80, 0.32,
                       r(f"{v}{unit}", colors[k] if rigor or colors[k] == RED else BODY, 1500, True), anchor="ctr",
                       name="Value"))


def build_a(tree):
    ids = reset(tree, "RESULT 1  /  NOTHING IS LOST", "First, nothing is lost: RIGOR keeps the schema.",
                "We climb from the easiest check to the hardest. Every number covers all 30 ontologies.")
    ladder(tree, ids, {1, 2})
    y, h = TOP, 3.45
    card(tree, ids, 0.88, y, 3.55, h, "1  VALID")
    add(tree, text(ids, 1.18, y + 0.70, 3.0, 1.0, r("30 / 30", INK, 5400, True, "Cambria"), name="Big"))
    add(tree, text(ids, 1.18, y + 1.75, 3.0, 1.2,
                   r("ontologies are valid OWL 2 DL and logically consistent, for every method.", BODY, 1500),
                   name="Claim"))
    card(tree, ids, 4.63, y, 7.82, h, "2  KEEPS THE SCHEMA", "share of database columns covered, per ontology")
    bars(tree, ids, 4.93, y + 1.05, 7.25, COVERAGE, "%", 100, ("9AA3A9", "9AA3A9", RED, EMERALD), pitch=0.55)
    strip(tree, ids, 6.43, r("Same retrieved context, no seed: ", INK, 1400, True)
          + r("Non-Iterative covers only a third of the columns. RIGOR's direct-mapping seed keeps them.",
              BODY, 1400))
    source(tree, ids, "Validity: RDFLib, consistency: HermiT (paper Sec. 4.7). Coverage: paper Table 9, averaged "
                      "over each method's ontologies.")


def build_b(tree):
    ids = reset(tree, "RESULT 2  /  MEANING AND ERRORS", "Then, meaning is added, with far fewer errors.",
                "Averages per ontology, over all 30 ontologies.")
    ladder(tree, ids, {3, 4})
    y, h = TOP, 3.45
    x, w = 0.88, 6.55
    card(tree, ids, x, y, w, h, "3  ADDS MEANING", "elements per ontology")
    cols = list(ADDED)
    col_x = [x + 2.30 + k * 1.40 for k in range(3)]
    for cx, label in zip(col_x, cols):
        add(tree, text(ids, cx, y + 0.85, 1.30, 0.46, r(label, MUTED, 1150), align="ctr", anchor="b", name="Column"))
    add(tree, shape(ids(), "Rule", x + 0.3, y + 1.38, w - 0.6, 0.012, fill="C5CCD0"))
    for k, method in enumerate(METHODS):
        ry = y + 1.50 + k * 0.45
        rigor = method == "RIGOR"
        if rigor:
            add(tree, shape(ids(), "RIGOR band", x + 0.18, ry - 0.04, w - 0.36, 0.42, geom="roundRect", fill=MINT))
        add(tree, text(ids, x + 0.3, ry, 2.0, 0.34, r(method, EMERALD if rigor else INK, 1450, rigor), anchor="ctr",
                       name="Method"))
        for cx, label in zip(col_x, cols):
            v = ADDED[label][k]
            color = EMERALD if rigor else (INK if v else "6B747B")
            add(tree, text(ids, cx, ry, 1.30, 0.34, r(f"{v:,}", color, 1600 if rigor else 1450, rigor or bool(v)),
                           align="ctr", anchor="ctr", name="Value"))
    x2, w2 = 7.63, 4.82
    card(tree, ids, x2, y, w2, h, "4  FEWER ERRORS", "OOPS! modeling pitfalls per ontology")
    bars(tree, ids, x2 + 0.30, y + 1.05, w2 - 0.45, PITFALLS, "", max(PITFALLS),
         ("9AA3A9", "9AA3A9", "9AA3A9", EMERALD), pitch=0.55, label_w=1.65)
    strip(tree, ids, 6.43, r("72%", INK, 1400, True) + r(" of the others' pitfalls: missing annotations.  "
                                                      "The critical one, a second domain: ", BODY, 1400)
          + r("zero in 6 of 9", EMERALD, 1400, True) + r(" RIGOR ontologies.", BODY, 1400))
    source(tree, ids, "Paper Table 8 (annotation assertions, prov:wasDerivedFrom, SubClassOf) and Table 7 (OOPS!; "
                      "P08 missing annotations, P19 multiple domains or ranges).")


def build_c(tree):
    """The author's BURR slide: renumber, cite, shift down for the ladder, add the relations line."""
    root, shapes, eyebrow, title = header_shapes(tree)
    set_text(eyebrow, "RESULT 3  /  EXTERNAL BENCHMARK")
    subtitle = next(sp for sp in shapes if sp_text(sp).startswith("BURR benchmark"))
    set_text(subtitle, "BURR [Laskowski et al., PACMMOD 2025]  ·  ISWC scenario  ·  highest F1 among "
                       "ontology-learning systems")
    footer = next(sp for sp in shapes if sp_text(sp).startswith("One scenario"))
    for sp in shapes:
        g = geometry(sp)
        if g and g[1] > 2.2 and sp is not footer:
            place(sp, y=g[1] + 0.42)
    place(footer, y=6.98)
    ids = Ids(tree)
    ladder(tree, ids, {5})
    strip(tree, ids, 6.43, r("Relations: 0.67 vs 0.11", INK, 1400, True)
          + r(" for W3C-Mapper / Ontop. Foreign keys become named relations, like hasPatientData.", BODY, 1400))


def build_d(tree):
    ids = reset(tree, "RESULT 4  /  LLM JUDGE AND HUMAN EXPERTS",
                "Finally, an LLM judge and human experts both rank RIGOR first.",
                "Both ranked four anonymized ontologies per table, from first to fourth.")
    ladder(tree, ids, {6})
    y, h, w = TOP, 3.45, 5.70
    x = 0.88
    card(tree, ids, x, y, w, h, "LLM-AS-A-JUDGE  ·  GPT-5.4", fill=PANEL)
    add(tree, text(ids, x + 0.3, y + 0.55, w - 0.6, 1.0, r("58 / 58", INK, 5400, True, "Cambria"), name="Big"))
    add(tree, text(ids, x + 0.3, y + 1.55, w - 0.6, 0.40, r("tables with RIGOR ranked first", EMERALD, 1800, True),
                   name="Claim"))
    add(tree, text(ids, x + 0.3, y + 2.15, w - 0.6, 0.32,
                   r("  ·  ".join(f"{db} {n}/{n}" for db, n in JUDGE), BODY, 1350), name="Split"))
    add(tree, text(ids, x + 0.3, y + 2.65, w - 0.6, 0.55,
                   r("Largest gains: ", INK, 1350, True) + r("clarity and domain enrichment.", BODY, 1350),
                   name="Gains"))
    x = 6.75
    card(tree, ids, x, y, w, h, "10 ONTOLOGY EXPERTS", fill=PANEL)
    add(tree, text(ids, x + 0.3, y + 0.55, w - 0.6, 1.0, r("28 / 40", INK, 5400, True, "Cambria"), name="Big"))
    add(tree, text(ids, x + 0.3, y + 1.55, w - 0.6, 0.40, r("rankings with RIGOR first (70%)", EMERALD, 1800, True),
                   name="Claim"))
    for k, (method, v) in enumerate(zip(METHODS, EXPERT_FIRST)):
        ry = y + 2.10 + k * 0.30
        rigor = method == "RIGOR"
        add(tree, text(ids, x + 0.3, ry, 1.75, 0.28, r(method, EMERALD if rigor else BODY, 1250, rigor),
                       anchor="ctr", name="Method"))
        add(tree, shape(ids(), "Track", x + 2.10, ry + 0.07, 2.60, 0.14, fill="E1E6E3"))
        add(tree, shape(ids(), "Bar", x + 2.10, ry + 0.07, 2.60 * v / 40, 0.14, fill=EMERALD if rigor else "9AA3A9"))
        add(tree, text(ids, x + 4.80, ry, 0.60, 0.28, r(str(v), EMERALD if rigor else BODY, 1250, rigor),
                       align="r", anchor="ctr", name="Count"))
    strip(tree, ids, 6.43, r("Direct Mapping: structure without meaning.  LLMs alone: meaning without structure.  ",
                             BODY, 1400) + r("RIGOR: both, and checked.", EMERALD, 1400, True))
    source(tree, ids, "Paper Table 1 (rankings), Table 10 (judge dimensions), Table 12 (expert agreement, Kendall's "
                      "W 0.39–0.72). Judge: 290 competency questions. Experts: 4 medical tables.")


def add_notes(work, slide_name, title):
    """Give a slide notes from speech.py, creating a notes slide if it has none (copied slides)."""
    rels_path = work / "ppt/slides/_rels" / f"{slide_name}.rels"
    m = re.search(r"notesSlides/(notesSlide\d+\.xml)", rels_path.read_text(encoding="utf-8"))
    notes_dir = work / "ppt/notesSlides"
    if not m:
        template = sorted(notes_dir.glob("notesSlide*.xml"))[0]
        n = max(int(re.findall(r"\d+", p.name)[0]) for p in notes_dir.glob("notesSlide*.xml")) + 1
        dest = notes_dir / f"notesSlide{n}.xml"
        shutil.copy(template, dest)
        rels = (notes_dir / "_rels" / f"{template.name}.rels").read_text(encoding="utf-8")
        (notes_dir / "_rels" / f"{dest.name}.rels").write_text(
            re.sub(r'Target="\.\./slides/slide\d+\.xml"', f'Target="../slides/{slide_name}"', rels), encoding="utf-8")
        ct = work / "[Content_Types].xml"
        ct.write_text(ct.read_text(encoding="utf-8").replace(
            "</Types>", f'<Override PartName="/ppt/notesSlides/{dest.name}" ContentType="application/vnd.'
                        'openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'), encoding="utf-8")
        rels_path.write_text(rels_path.read_text(encoding="utf-8").replace(
            "</Relationships>", '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/'
                                f'2006/relationships/notesSlide" Target="../notesSlides/{dest.name}"/></Relationships>'),
            encoding="utf-8")
        name = dest.name
    else:
        name = m.group(1)
    n_, s, start, end, _ = next(row for row in timeline() if row[1]["title"] == title)
    set_notes(notes_dir / name, notes_for(n_, s, start, end))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def order():
        pres = (work / "ppt/presentation.xml").read_text(encoding="utf-8")
        rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
                               (work / "ppt/_rels/presentation.xml.rels").read_text()))
        return [rels[x] for x in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', pres)]

    def find(tag):
        return next(n for n in order() if tag in (work / "ppt/slides" / n).read_text(encoding="utf-8"))

    r1, r2, r3 = find("RESULT 1"), find("RESULT 2"), find("RESULT 3")
    # Slide A: a copy of Result 1, inserted before it.
    prev = order()[order().index(r1) - 1]
    out = subprocess.run([sys.executable, str(PPTX_SCRIPTS / "add_slide.py"), str(work), r1, "--after", prev],
                         capture_output=True, text=True, cwd=PPTX_SCRIPTS, check=True)
    a = re.findall(r"(slide\d+\.xml)", out.stdout)[0]
    plan = ((a, build_a, "Result 1: nothing is lost"), (r1, build_b, "Result 2: meaning and errors"),
            (r2, build_c, "Result 3: BURR"), (r3, build_d, "Result 4: LLM judge and experts"))
    for name, build, title in plan:
        path = work / "ppt/slides" / name
        tree = etree.parse(str(path))
        build(tree)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
        add_notes(work, name, title)
    # Keep only the four result slides.
    keep = {a, r1, r2, r3}
    pres_path = work / "ppt/presentation.xml"
    pres = etree.parse(str(pres_path))
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
                           (work / "ppt/_rels/presentation.xml.rels").read_text()))
    lst = pres.find("p:sldIdLst", NS)
    for sld in list(lst):
        if rels[sld.get(f"{{{R}}}id")] not in keep:
            lst.remove(sld)
    for sec in pres.iter("{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst"):
        sec.getparent().remove(sec)
    pres.write(str(pres_path), xml_declaration=True, encoding="UTF-8", standalone=True)
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
