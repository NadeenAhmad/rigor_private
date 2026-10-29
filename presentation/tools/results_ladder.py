#!/usr/bin/env python3
"""Results as a ladder of five questions, easiest check first, on the author's edited deck.

    1 VALID, FEWER ERRORS  2 KEEPS THE SCHEMA  3 ADDS MEANING  4 GOLD MAPPINGS  5 JUDGES

  A  RESULT 1  validity in one sentence (Sec. 4.7), then OOPS! pitfalls per run (Table 7)
  B  RESULT 2  column coverage per run (Table 9)
  C  RESULT 3  labels, provenance and hierarchy per run (Table 8)
  D  RESULT 4  the author's BURR slide (one Claude run), shifted down for the tracker, plus the relations line
  E  RESULT 5  LLM judge and experts per database, both against competency questions (Tables 1, 6, 12)

Runs are shown per database x LLM (Direct Mapping: one ontology per database) with each method's
average. A and C are copies of the author's Result 1 slide; the output deck holds only these five
slides, with speaker notes from speech.py.

Usage: python3 results_ladder.py <in.pptx> <out.pptx>
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
LADDER = ("VALID, FEWER ERRORS", "KEEPS THE SCHEMA", "ADDS MEANING", "GOLD MAPPINGS", "JUDGES")
TOP = 2.85  # content starts below the ladder

COVERAGE = (91, 74, 34, 93)            # Table 9: mean column coverage per ontology, %
PITFALLS = (499, 307, 52, 8)           # Table 7: mean OOPS! pitfalls per ontology


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
    gap = 0.10
    x, y, w = 0.88, 2.30, (11.57 - gap * (len(LADDER) - 1)) / len(LADDER)
    for k, label in enumerate(LADDER, 1):
        if k in current:
            fill, line, color = EMERALD, None, "FFFFFF"
        elif not current or k < min(current):
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


# Per-run data for the heat tables: databases in talk order, LLMs Claude, Mistral, DeepSeek.
DBS = ("Liver Cancer", "eICU-CRD", "Chinook")
LLMS = ("Claude", "Mistral", "DeepSeek")
# Table 9 column coverage (%). Direct Mapping: one ontology per database.
COV_RUNS = {"Direct Mapping": (97, 99, 78),
            "LLM baseline": (94, 31, 89, 96, 68, 86, 100, 50, 52),
            "Non-Iterative": (41, 6, 12, 22, 28, 46, 58, 16, 78),
            "RIGOR": (91, 99, 87, 100, 99, 77, 95, 95, 95)}
# Table 7 OOPS! pitfall totals.
PIT_RUNS = {"Direct Mapping": (731, 665, 102),
            "LLM baseline": (422, 397, 403, 434, 440, 421, 95, 74, 81),
            "Non-Iterative": (29, 61, 18, 39, 220, 27, 24, 35, 14),
            "RIGOR": (20, 23, 3, 2, 2, 26, 0, 0, 0)}
# Table 1: LLM-judge average score (0-5, Claude ontologies) and expert first places (of 20).
JUDGE_SCORE = {"Direct Mapping": (2.05, 2.13, 0.13), "LLM baseline": (1.38, 0.28, 1.48),
               "Non-Iterative": (1.89, 0.67, 2.77), "RIGOR": (3.93, 3.64, 3.32)}
EXPERT_DB = {"Direct Mapping": (3, 2), "LLM baseline": (1, 4), "Non-Iterative": (2, 0), "RIGOR": (14, 14)}
GOOD2, GOOD1, MID, BAD1, BAD2 = "BFE3CF", "DDF0E5", "F1F3F2", "F6DAD3", "EDB8AC"


def shade(v, bands):
    """bands: four ascending cut points; returns the fill for v (higher = better)."""
    return (BAD2, BAD1, MID, GOOD1, GOOD2)[sum(v >= b for b in bands)]


def heat_table(tree, ids, x, y, label_w, col_w, row_h, runs, fill_for, fmt, averages=None, avg_fmt=str,
               cols=DBS, sub=LLMS, header=True, fill_by_col=None):
    """Rows = methods; columns = databases, each split by LLM (unless a row has one value per database).
    Returns the y below the last row."""
    n = len(sub) if sub else 1
    gx = x + label_w
    for g, db in enumerate(cols if header else ()):
        add(tree, text(ids, gx + g * n * col_w, y, n * col_w, 0.24, r(db, INK, 1100, True), align="ctr",
                       anchor="ctr", name="Database"))
        if sub:
            for k, llm in enumerate(sub):
                add(tree, text(ids, gx + (g * n + k) * col_w, y + 0.25, col_w, 0.20, r(llm, MUTED, 850),
                               align="ctr", anchor="ctr", name="LLM"))
    if averages and header:
        add(tree, text(ids, gx + len(cols) * n * col_w + 0.05, y + 0.12, 0.75, 0.26, r("Avg", MUTED, 1000, True),
                       align="ctr", anchor="ctr", name="Avg head"))
    top = y + ((0.50 if sub else 0.28) if header else 0)
    for k, (method, values) in enumerate(runs.items()):
        ry = top + k * row_h
        rigor = method == "RIGOR"
        add(tree, text(ids, x, ry, label_w, row_h, r(method, EMERALD if rigor else INK, 1200, rigor), anchor="ctr",
                       name="Method"))
        span = n if len(values) == len(cols) else 1
        for j, v in enumerate(values):
            cx = gx + j * span * col_w
            fill = fill_by_col(v, j * span) if fill_by_col else fill_for(v)
            add(tree, shape(ids(), "Cell", cx + 0.02, ry + 0.02, span * col_w - 0.04, row_h - 0.04,
                            r(fmt(v), INK, 1100, rigor), fill=fill, align="ctr", anchor="ctr"))
        if averages:
            add(tree, text(ids, gx + len(cols) * n * col_w + 0.05, ry, 0.75, row_h,
                           r(avg_fmt(averages[k]), EMERALD if rigor else INK, 1250, rigor), align="ctr",
                           anchor="ctr", name="Avg"))
    return top + len(runs) * row_h


# Table 8 per run: annotation assertions and prov:wasDerivedFrom assertions; hierarchy as inheritance richness
# (OntoQA: subclass axioms per class), from Table 8's SubClassOf and class counts. Direct Mapping and the LLM
# baseline have none of these in any run.
NI = "Non-Iterative RAG"
ANN_RUNS = {NI: (680, 134, 736, 877, 157, 743, 98, 14, 64),
            "RIGOR": (1029, 1137, 341, 1291, 1684, 435, 357, 362, 359)}
PROV_RUNS = {NI: (0,) * 9, "RIGOR": (336, 308, 22, 423, 335, 99, 81, 83, 75)}
IR_RUNS = {NI: (0.0,) * 9, "RIGOR": (0.70, 0.39, 0.13, 0.62, 0.49, 0.10, 2.12, 2.72, 4.07)}


def per_database(runs):
    """Shade each cell against the largest value of the same database (all methods and LLMs)."""
    peak = [max(max(v[g * 3:g * 3 + 3]) for v in runs.values()) or 1 for g in range(3)]

    def fill(v, j):
        if not v:
            return BAD2
        return (BAD1, MID, GOOD1, GOOD2)[min(3, int(4 * v / peak[j // 3] - 1e-9))]
    return fill


def build_a(tree):
    ids = reset(tree, "RESULT 1  /  VALID, WITH FEWER ERRORS",
                "Every ontology is valid. RIGOR's have far fewer modeling errors.",
                "All 30 ontologies are valid OWL 2 DL (RDFLib) and logically consistent (HermiT). So we look deeper.")
    ladder(tree, ids, {1})
    card(tree, ids, 0.88, TOP, 11.57, 3.45, "MODELING ERRORS FOUND BY OOPS!", "pitfalls per ontology, fewer is better")
    heat_table(tree, ids, 1.18, TOP + 0.85, 1.75, 0.92, 0.50, PIT_RUNS,
               lambda v: shade(-v, (-300, -100, -30, -10)), str, averages=PITFALLS)
    strip(tree, ids, 6.43, r("Same LLM, same database: ", BODY, 1400)
          + r("Mistral on eICU, 220 without RIGOR", RED, 1400, True) + r(", ", BODY, 1400)
          + r("2 with it.", EMERALD, 1400, True) + r("  On Chinook, RIGOR has zero with every LLM.", BODY, 1400))
    source(tree, ids, "Validity: RDFLib; consistency: HermiT (paper Sec. 4.7). Pitfalls: OOPS!, paper Table 7; 72% of "
                      "the other methods' pitfalls are missing annotations (P08).")


def build_b(tree):
    ids = reset(tree, "RESULT 2  /  NOTHING IS LOST", "Nothing is lost: RIGOR keeps the schema.",
                "Share of database columns covered, for every LLM and database.")
    ladder(tree, ids, {2})
    card(tree, ids, 0.88, TOP, 11.57, 3.45, "KEEPS THE SCHEMA", "columns covered, per ontology (%)")
    heat_table(tree, ids, 1.18, TOP + 0.85, 1.75, 0.92, 0.50, COV_RUNS,
               lambda v: shade(v, (20, 50, 75, 90)), str, averages=COVERAGE, avg_fmt=lambda v: f"{v}%")
    strip(tree, ids, 6.43, r("Same context, no seed: ", BODY, 1400) + r("Non-Iterative 6–78%", RED, 1400, True)
          + r(", depending on the LLM.  With RIGOR: ", BODY, 1400)
          + r("≥ 77% for every LLM, in every database.", EMERALD, 1400, True))
    source(tree, ids, "Coverage: a column counts as covered when an ontology class or property name is similar to it "
                      "(embedding similarity ≥ 0.55). Paper Sec. 4.6, Table 9.")


def build_meaning(tree):
    ids = reset(tree, "RESULT 3  /  MEANING IS ADDED", "Meaning is added: labels, provenance and hierarchy.",
                "Per ontology, for every LLM and database.")
    ladder(tree, ids, {3})
    card(tree, ids, 0.88, TOP, 11.57, 3.45, "ADDS MEANING")
    x_metric, x_table, col_w, row_h = 1.18, 2.75, 0.82, 0.28
    gx = x_table + 1.40
    y = TOP + 0.42
    # Header row (databases and LLMs) via an empty table, then the two methods that add nothing.
    heat_table(tree, ids, x_table, y, 1.40, col_w, row_h, {}, None, str, averages=[0], avg_fmt=str)
    y += 0.50
    for k, method in enumerate(("Direct Mapping", "LLM baseline")):
        add(tree, text(ids, x_table, y + k * 0.20, 1.40, 0.20, r(method, INK, 1100), anchor="ctr", name="Method"))
    add(tree, shape(ids(), "None", gx + 0.02, y + 0.01, 9 * col_w - 0.04, 0.38,
                    r("No labels, provenance or class hierarchy, in any run.", INK, 1150, italic=True),
                    fill=BAD1, align="ctr", anchor="ctr"))
    y += 0.46
    for metric, runs, fmt in (("Labels & comments", ANN_RUNS, lambda v: f"{v:,}"),
                              ("Provenance links", PROV_RUNS, lambda v: f"{v:,}"),
                              ("Class hierarchy", IR_RUNS, lambda v: f"{v:.2f}")):
        if runs is IR_RUNS:
            add(tree, text(ids, x_metric, y + 0.02, 1.55, 0.26, r(metric, TEAL, 1150, True), anchor="ctr",
                           name="Metric"))
            add(tree, text(ids, x_metric, y + row_h, 1.55, 0.24, r("inheritance richness", MUTED, 950, italic=True),
                           anchor="ctr", name="Metric note"))
        else:
            add(tree, text(ids, x_metric, y, 1.55, 2 * row_h, r(metric, TEAL, 1150, True), anchor="ctr",
                           name="Metric"))
        avgs = [sum(v) / len(v) for v in runs.values()]
        avg_fmt = (lambda v: f"{v:.2f}") if runs is IR_RUNS else (lambda v: f"{round(v):,}")
        y = heat_table(tree, ids, x_table, y, 1.40, col_w, row_h, runs, None, fmt, averages=avgs, avg_fmt=avg_fmt,
                       header=False, fill_by_col=per_database(runs)) + 0.05
    strip(tree, ids, 6.43, r("Only RIGOR adds labels, provenance ", BODY, 1400) + r("and", INK, 1400, True)
          + r(" a class hierarchy: ", BODY, 1400) + r("with every LLM, in every database.", EMERALD, 1400, True))
    source(tree, ids, "Labels & comments: annotation assertions; provenance: prov:wasDerivedFrom (paper Table 8). "
                      "Hierarchy: inheritance richness, subclass axioms per class [OntoQA: Tartir et al., KADASH 2005; "
                      "OntoMetrics: Lantow, KEOD 2016], from Table 8. Shading compares cells within each database.",
           y=6.94)


def build_c(tree):
    """The author's BURR slide: renumber, cite, shift down for the ladder, add the relations line."""
    root, shapes, eyebrow, title = header_shapes(tree)
    set_text(eyebrow, "RESULT 4  /  EXTERNAL BENCHMARK")
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
    ladder(tree, ids, {4})
    strip(tree, ids, 6.43, r("Relations: 0.67 vs 0.11", INK, 1400, True)
          + r(" for W3C-Mapper / Ontop. Foreign keys become named relations, like hasPatientData.", BODY, 1400))


def build_d(tree):
    ids = reset(tree, "RESULT 5  /  LLM JUDGE AND HUMAN EXPERTS",
                "An LLM judge and human experts both rank RIGOR first.",
                "Both ranked four anonymized ontologies per table, judged against that table's competency questions.")
    ladder(tree, ids, {5})
    add(tree, shape(ids(), "CQ strip", 0.88, TOP, 11.57, 0.52,
                    r("COMPETENCY QUESTIONS  ", TEAL, 1150, True, "Consolas")
                    + r("what the ontology must answer. 5 per table, 290 in total, expert-verified.  ",
                        BODY, 1150)
                    + r("e.g. “Which patients had abnormal physical exam findings?”",
                        INK, 1150, italic=True),
                    geom="roundRect", fill="FFFFFF", line=LINE, anchor="ctr", inset=0.22))
    y, h = TOP + 0.62, 2.83
    x, w = 0.88, 6.30
    card(tree, ids, x, y, w, h, "LLM-AS-A-JUDGE  ·  GPT-5.4", fill=PANEL)
    add(tree, text(ids, x + 0.3, y + 0.42, 2.4, 0.62, r("58 / 58", INK, 3400, True, "Cambria"), name="Big"))
    add(tree, text(ids, x + 2.55, y + 0.45, w - 2.8, 0.55, r("tables with RIGOR ranked first", EMERALD, 1500, True),
                   anchor="ctr", name="Claim"))
    heat_table(tree, ids, x + 0.25, y + 1.08, 1.45, 1.42, 0.34, JUDGE_SCORE,
               lambda v: shade(v, (1.0, 2.0, 3.0, 3.5)), lambda v: f"{v:.2f}", sub=None)
    x2, w2 = 7.33, 5.12
    card(tree, ids, x2, y, w2, h, "10 ONTOLOGY EXPERTS", fill=PANEL)
    add(tree, text(ids, x2 + 0.3, y + 0.42, 2.4, 0.62, r("28 / 40", INK, 3400, True, "Cambria"), name="Big"))
    add(tree, text(ids, x2 + 2.45, y + 0.45, w2 - 2.6, 0.55, r("rankings with RIGOR first (70%)", EMERALD, 1500,
                                                              True), anchor="ctr", name="Claim"))
    heat_table(tree, ids, x2 + 0.25, y + 1.08, 1.45, 1.55, 0.34, EXPERT_DB,
               lambda v: shade(v, (1, 3, 6, 10)), str, cols=DBS[:2], sub=None)
    strip(tree, ids, 6.43, r("Direct Mapping: structure without meaning.  LLMs alone: meaning without structure.  ",
                             BODY, 1400) + r("RIGOR: both, and checked.", EMERALD, 1400, True))
    source(tree, ids, "Judge: average score 0–5 per database, Claude ontologies, 58 tables. Experts: first places of "
                      "20 per database, 4 medical tables, Kendall's W 0.39–0.72. Paper Tables 1, 6, 12.")


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

    def copy(src, after):
        out = subprocess.run([sys.executable, str(PPTX_SCRIPTS / "add_slide.py"), str(work), src, "--after", after],
                             capture_output=True, text=True, cwd=PPTX_SCRIPTS, check=True)
        return re.findall(r"(slide\d+\.xml)", out.stdout)[0]

    a = copy(r1, order()[order().index(r1) - 1])   # errors, before coverage
    c = copy(r1, r1)                               # meaning, after coverage
    plan = ((a, build_a, "Result 1: valid, fewer errors"), (r1, build_b, "Result 2: keeps the schema"),
            (c, build_meaning, "Result 3: adds meaning"), (r2, build_c, "Result 4: BURR"),
            (r3, build_d, "Result 5: LLM judge and experts"))
    for name, build, title in plan:
        path = work / "ppt/slides" / name
        tree = etree.parse(str(path))
        build(tree)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
        add_notes(work, name, title)
    # Keep only the five result slides.
    keep = {a, r1, c, r2, r3}
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
