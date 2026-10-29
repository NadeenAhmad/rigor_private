#!/usr/bin/env python3
"""Results as a ladder of six questions, easiest check first, on the author's edited deck.

    1 VALID  2 KEEPS THE SCHEMA  3 ADDS MEANING  4 FEWER ERRORS  5 GOLD MAPPINGS  6 JUDGES

  A  RESULT 1  rungs 1-2  validity (Sec. 4.7) and column coverage (Table 9), averaged per ontology
  B  RESULT 2  rungs 3-4  added meaning (Table 8) and OOPS! pitfalls (Table 7), averaged per ontology
  C  RESULT 3  rung 5     the author's BURR slide, shifted down for the tracker, plus the relations line
  D  RESULT 4  rung 6     LLM judge and human experts (Tables 1, 10, 12), plus the closing sentence
  E  RESULT 5  (ladder done)  where the gains come from: across LLMs (Mistral Small, Tables 9 and 15)
                         and across domains (LLM-judge lead per database, Table 1)

A-D average over all 30 ontologies (each method's own runs); E breaks them down by LLM and domain.
A is a copy of the author's Result 1 slide and E a copy of Result 3; the output deck holds only these
five slides, with speaker notes from speech.py.

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
# E, across LLMs: Mistral Small 24B, mean column coverage over the three databases (Table 9), %.
MISTRAL = (("LLM baseline", 50), ("Non-Iterative", 17), ("RIGOR", 98))
CLAUDE_RIGOR, COST_MISTRAL, COST_CLAUDE = 95, "$0.12", "$33"   # Table 9; Table 15, all three databases
# E, across domains: LLM-judge average score, RIGOR minus the best other method (Table 1).
LEAD = (("Liver Cancer", 1.9, "private registry, unseen by any LLM"),
        ("eICU-CRD", 1.5, "large ICU schema, 559 columns"),
        ("Chinook", 0.6, "music store, a familiar domain"))


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


def bars(tree, ids, x, y, w, values, unit, maximum, colors, pitch=0.48, label_w=1.75, labels=METHODS):
    """Horizontal bars, one row per label, value label after the bar."""
    track = w - label_w - 0.85
    for k, (method, v) in enumerate(zip(labels, values)):
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
               cols=DBS, sub=LLMS):
    """Rows = methods; columns = databases, each split by LLM (unless a row has one value per database)."""
    n = len(sub) if sub else 1
    gx = x + label_w
    for g, db in enumerate(cols):
        add(tree, text(ids, gx + g * n * col_w, y, n * col_w, 0.24, r(db, INK, 1100, True), align="ctr",
                       anchor="ctr", name="Database"))
        if sub:
            for k, llm in enumerate(sub):
                add(tree, text(ids, gx + (g * n + k) * col_w, y + 0.25, col_w, 0.20, r(llm, MUTED, 850),
                               align="ctr", anchor="ctr", name="LLM"))
    if averages:
        add(tree, text(ids, gx + len(cols) * n * col_w + 0.05, y + 0.12, 0.75, 0.26, r("Avg", MUTED, 1000, True),
                       align="ctr", anchor="ctr", name="Avg head"))
    top = y + (0.50 if sub else 0.28)
    for k, (method, values) in enumerate(runs.items()):
        ry = top + k * row_h
        rigor = method == "RIGOR"
        add(tree, text(ids, x, ry, label_w, row_h, r(method, EMERALD if rigor else INK, 1200, rigor), anchor="ctr",
                       name="Method"))
        span = n if len(values) == len(cols) else 1
        for j, v in enumerate(values):
            cx = gx + j * span * col_w
            add(tree, shape(ids(), "Cell", cx + 0.02, ry + 0.02, span * col_w - 0.04, row_h - 0.04,
                            r(fmt(v), INK, 1100, rigor), fill=fill_for(v), align="ctr", anchor="ctr"))
        if averages:
            add(tree, text(ids, gx + len(cols) * n * col_w + 0.05, ry, 0.75, row_h,
                           r(avg_fmt(averages[k]), EMERALD if rigor else INK, 1250, True), align="ctr",
                           anchor="ctr", name="Avg"))


def build_a(tree):
    ids = reset(tree, "RESULT 1  /  NOTHING IS LOST", "First, nothing is lost: RIGOR keeps the schema.",
                "Every LLM, every database. We climb from the easiest check to the hardest.")
    ladder(tree, ids, {1, 2})
    y, h = TOP, 3.45
    card(tree, ids, 0.88, y, 2.45, h, "1  VALID")
    add(tree, text(ids, 1.13, y + 0.70, 2.1, 0.9, r("30 / 30", INK, 4000, True, "Cambria"), name="Big"))
    add(tree, text(ids, 1.13, y + 1.60, 2.0, 1.6,
                   r("ontologies are valid OWL 2 DL and logically consistent, for every method and LLM.", BODY,
                     1350), name="Claim"))
    card(tree, ids, 3.48, y, 8.97, h, "2  KEEPS THE SCHEMA", "share of database columns covered, per ontology (%)")
    heat_table(tree, ids, 3.73, y + 0.85, 1.45, 0.70, 0.50, COV_RUNS,
               lambda v: shade(v, (20, 50, 75, 90)), str, averages=COVERAGE, avg_fmt=lambda v: f"{v}%")
    strip(tree, ids, 6.43, r("Without the seed, coverage swings with the LLM: ", BODY, 1400)
          + r("Non-Iterative 6–78%.", RED, 1400, True)
          + r("  With RIGOR, every LLM keeps ≥ 77% in every database.", EMERALD, 1400, True))
    source(tree, ids, "Validity: RDFLib; consistency: HermiT. Coverage: a column counts as covered when an ontology "
                      "class or property name is similar to it (embedding similarity ≥ 0.55). Paper Sec. "
                      "4.6–4.7, Table 9.")


def build_b(tree):
    ids = reset(tree, "RESULT 2  /  MEANING AND ERRORS", "Then, meaning is added, with far fewer errors.",
                "What each method adds, and the modeling errors per ontology, for every LLM and database.")
    ladder(tree, ids, {3, 4})
    y, h = TOP, 3.45
    x, w = 0.88, 3.80
    card(tree, ids, x, y, w, h, "3  ADDS MEANING", "elements per ontology (avg)")
    cols = (("Labels", "Labels & comments"), ("Provenance", "Provenance links"), ("Hierarchy", "Subclass axioms"))
    col_x = [x + 1.40 + k * 0.78 for k in range(3)]
    for cx, (short, _) in zip(col_x, cols):
        add(tree, text(ids, cx - 0.06, y + 0.95, 0.90, 0.26, r(short, MUTED, 1000), align="ctr", name="Column"))
    add(tree, shape(ids(), "Rule", x + 0.25, y + 1.25, w - 0.5, 0.012, fill="C5CCD0"))
    for k, method in enumerate(METHODS):
        ry = y + 1.35 + k * 0.44
        rigor = method == "RIGOR"
        if rigor:
            add(tree, shape(ids(), "RIGOR band", x + 0.15, ry - 0.03, w - 0.30, 0.40, geom="roundRect", fill=MINT))
        add(tree, text(ids, x + 0.25, ry, 1.2, 0.34, r(method, EMERALD if rigor else INK, 1150, rigor),
                       anchor="ctr", name="Method"))
        for cx, (_, key) in zip(col_x, cols):
            v = ADDED[key][k]
            color = EMERALD if rigor else (INK if v else "6B747B")
            add(tree, text(ids, cx - 0.06, ry, 0.90, 0.34, r(f"{v:,}", color, 1300, rigor or bool(v)),
                           align="ctr", anchor="ctr", name="Value"))
    add(tree, text(ids, x + 0.25, y + 3.12, w - 0.5, 0.28, r("Same pattern for every LLM and database.",
                                                            MUTED, 1000, italic=True), name="Note"))
    x2, w2 = 4.83, 7.62
    card(tree, ids, x2, y, w2, h, "4  FEWER ERRORS", "OOPS! modeling pitfalls per ontology")
    heat_table(tree, ids, x2 + 0.22, y + 0.85, 1.30, 0.58, 0.50, PIT_RUNS,
               lambda v: shade(-v, (-300, -100, -30, -10)), str, averages=PITFALLS)
    strip(tree, ids, 6.43, r("Without validation, the smaller LLM fails hardest: ", BODY, 1400)
          + r("Mistral on eICU, 220.", RED, 1400, True) + r("  With RIGOR: ", BODY, 1400)
          + r("2, and zero on Chinook for every LLM.", EMERALD, 1400, True))
    source(tree, ids, "Paper Table 8 (annotations, prov:wasDerivedFrom, SubClassOf) and Table 7 (OOPS! totals). 72% "
                      "of the other methods' pitfalls are missing annotations (P08).")

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
                "An LLM judge and human experts both rank RIGOR first.",
                "Blind rankings of four anonymized ontologies per table, broken down by database.")
    ladder(tree, ids, {6})
    y, h = TOP, 3.45
    x, w = 0.88, 6.30
    card(tree, ids, x, y, w, h, "LLM-AS-A-JUDGE  ·  GPT-5.4", fill=PANEL)
    add(tree, text(ids, x + 0.3, y + 0.50, 2.6, 0.75, r("58 / 58", INK, 4000, True, "Cambria"), name="Big"))
    add(tree, text(ids, x + 2.85, y + 0.62, w - 3.1, 0.55, r("tables with RIGOR ranked first", EMERALD, 1600, True),
                   anchor="ctr", name="Claim"))
    add(tree, text(ids, x + 0.3, y + 1.28, w - 0.6, 0.24, r("average score, 0–5 (Claude ontologies)", MUTED, 1050,
                                                           italic=True), name="Note"))
    heat_table(tree, ids, x + 0.25, y + 1.52, 1.45, 1.42, 0.38, JUDGE_SCORE,
               lambda v: shade(v, (1.0, 2.0, 3.0, 3.5)), lambda v: f"{v:.2f}", sub=None)
    x2, w2 = 7.33, 5.12
    card(tree, ids, x2, y, w2, h, "10 ONTOLOGY EXPERTS", fill=PANEL)
    add(tree, text(ids, x2 + 0.3, y + 0.50, 2.6, 0.75, r("28 / 40", INK, 4000, True, "Cambria"), name="Big"))
    add(tree, text(ids, x2 + 2.80, y + 0.62, w2 - 3.0, 0.55, r("rankings with RIGOR first (70%)", EMERALD, 1600, True),
                   anchor="ctr", name="Claim"))
    add(tree, text(ids, x2 + 0.3, y + 1.28, w2 - 0.6, 0.24, r("first places, of 20 per database", MUTED, 1050,
                                                             italic=True), name="Note"))
    heat_table(tree, ids, x2 + 0.25, y + 1.52, 1.45, 1.55, 0.38, EXPERT_DB,
               lambda v: shade(v, (1, 3, 6, 10)), str, cols=DBS[:2], sub=None)
    strip(tree, ids, 6.43, r("Direct Mapping: structure without meaning.  LLMs alone: meaning without structure.  ",
                             BODY, 1400) + r("RIGOR: both, and checked.", EMERALD, 1400, True))
    source(tree, ids, "Paper Table 1 (scores, rankings; judge: 58 tables, 290 competency questions) and Table 12 "
                      "(expert agreement, Kendall's W 0.39–0.72; 4 medical tables).")


def build_e(tree):
    ids = reset(tree, "RESULT 5  /  WHERE THE GAINS COME FROM",
                "RIGOR's gains do not depend on model size or memorized data.",
                "Is it just a strong LLM? Is it memorized data? We broke the results down by LLM and by domain.")
    ladder(tree, ids, set())
    y, h, w = TOP, 3.45, 5.70
    x = 0.88
    card(tree, ids, x, y, w, h, "ACROSS LLMs  \u00b7  MISTRAL SMALL (24B)",
         "column coverage, average over the three databases")
    bars(tree, ids, x + 0.30, y + 1.00, w - 0.40, [v for _, v in MISTRAL], "%", 100, ("9AA3A9", RED, EMERALD),
         pitch=0.52, label_w=1.60, labels=[m for m, _ in MISTRAL])
    add(tree, text(ids, x + 0.30, y + 2.65, w - 0.6, 0.62,
                   r("On par with Claude under RIGOR (", BODY, 1350) + r(f"{CLAUDE_RIGOR}%", INK, 1350, True)
                   + r("), for ", BODY, 1350) + r(COST_MISTRAL, EMERALD, 1350, True)
                   + r(f" instead of {COST_CLAUDE} for all three databases.", BODY, 1350), name="Claim"))
    x = 6.75
    card(tree, ids, x, y, w, h, "ACROSS DOMAINS  \u00b7  LLM JUDGE",
         "RIGOR's lead over the best other method (score 0\u20135)")
    track = 2.30
    for k, (db, lead, note) in enumerate(LEAD):
        ry = y + 1.00 + k * 0.72
        top = k == 0
        add(tree, text(ids, x + 0.30, ry, 1.75, 0.30, r(db, EMERALD if top else INK, 1400, True), anchor="ctr",
                       name="Database"))
        add(tree, text(ids, x + 0.30, ry + 0.30, 2.6, 0.26, r(note, MUTED, 1100, italic=True), name="Note"))
        add(tree, shape(ids(), "Track", x + 2.55, ry + 0.08, track, 0.16, fill="EEF1EF"))
        add(tree, shape(ids(), "Bar", x + 2.55, ry + 0.08, track * lead / 2.0, 0.16,
                        fill=EMERALD if top else "6FA88F"))
        add(tree, text(ids, x + 2.55 + track + 0.08, ry, 0.70, 0.32, r(f"+{lead}", EMERALD, 1500, True),
                       anchor="ctr", name="Lead"))
    strip(tree, ids, 6.43, r("A small open model catches up, and the lead grows where the LLM knows least: ", BODY, 1400)
          + r("grounding and validation make the difference.", EMERALD, 1400, True))
    source(tree, ids, "Coverage: paper Table 9 (Claude 95%, DeepSeek 86% with RIGOR). Cost: Table 15. Judge scores: "
                      "Table 1 (Claude ontologies; best other: Direct Mapping, Direct Mapping, Non-Iterative).")


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
    # Slide E: a copy of Result 3 (the judge slide), inserted after it.
    out = subprocess.run([sys.executable, str(PPTX_SCRIPTS / "add_slide.py"), str(work), r3, "--after", r3],
                         capture_output=True, text=True, cwd=PPTX_SCRIPTS, check=True)
    e = re.findall(r"(slide\d+\.xml)", out.stdout)[0]
    plan = ((a, build_a, "Result 1: nothing is lost"), (r1, build_b, "Result 2: meaning and errors"),
            (r2, build_c, "Result 3: BURR"), (r3, build_d, "Result 4: LLM judge and experts"),
            (e, build_e, "Result 5: where the gains come from"))
    for name, build, title in plan:
        path = work / "ppt/slides" / name
        tree = etree.parse(str(path))
        build(tree)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
        add_notes(work, name, title)
    # Keep only the four result slides.
    keep = {a, r1, r2, r3, e}
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
