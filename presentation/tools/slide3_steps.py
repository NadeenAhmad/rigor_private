#!/usr/bin/env python3
"""Slide 3 ("What we want to build") as separate reveal slides, for a PDF without animations.

Works on the author's edited deck, where the reveal is a run of slides titled
"WHAT WE WANT TO BUILD": the first shows the empty target, the next four add one evidence source
each (schema, documentation, Disease Ontology, LLM knowledge) and highlight what they add, the
last one shows the complete target with "What this enables".

On every slide of the run:
  - the input table gets the admission_id primary-key row;
  - the target code is the final slide-9 example (code_excerpts.DIFF_ROWS, right column), on a fixed
    row grid so nothing moves between slides; rows of later steps stay empty;
  - each enriched row carries a right-aligned gloss in its own text box (no tab stops, which
    Google Slides does not keep).
Every reveal state is verified with the OWL API (code_excerpts.py verify).

Usage: python3 slide3_steps.py <in.pptx> <out.pptx>
"""
import copy
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import (A, DIFF_ROWS, row_step, EMU, GLOSS_COLORS, NS, P, add, code_line, move, next_id,  # noqa: E402
                           para, run_xml, simple_box)
from slide9_diff import geometry, shape_id, text  # noqa: E402

TITLE = "WHAT WE WANT TO BUILD"
HIGHLIGHT = "FFE599"  # the author's highlight for what a step adds
SIZE, GLOSS_SIZE = 1050, 1000
ROW_SHIFT = 0.38      # input-table row pitch

PANEL_BOTTOM = 7.12
CODE_TOP, CODE_BOTTOM, ROW_PAD = 3.50, 7.00, 0.07


def gloss(row):
    _, code, kind, group = row
    if kind == "changed":
        return "schema (foreign key)"
    return {"schema": "source column", "docs": "documentation", "vocab": "Disease Ontology",
            "llm": "LLM knowledge"}.get(group, "")


def highlighted(row, step):
    """Rows a step adds and highlights: step 1 marks only the schema-derived meaning (the renamed
    relation and provenance), not the plain structure; later steps mark everything they add."""
    _, code, kind, group = row
    if row_step(row) != step:
        return False
    return step > 1 or kind == "changed" or (kind == "added" and group == "schema")


def add_highlight(line_sp):
    """Highlight every run of a code line except its leading indentation."""
    runs = list(line_sp.iter(f"{{{A}}}r"))
    if runs and not runs[0].find("a:t", NS).text.strip():
        runs = runs[1:]
    for r in runs:
        rpr = r.find("a:rPr", NS)
        hl = etree.Element(f"{{{A}}}highlight")
        etree.SubElement(hl, f"{{{A}}}srgbClr", val=HIGHLIGHT)
        rpr.find("a:solidFill", NS).addnext(hl)


def find(shapes, test):
    return next(sp for sp in shapes if (g := geometry(sp)) and test(g, sp))


def edit(tree, step, final):
    root = tree.find(".//p:spTree", NS)
    shapes = [el for el in root if el.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp")]

    # Input table: rows down one pitch, admission_id PK as the new first row.
    card = find(shapes, lambda g, sp: g[0] < 1.0 and 2.0 < g[1] < 2.5 and g[3] > 2)
    table = find(shapes, lambda g, sp: 1.0 < g[0] < 1.2 and 2.7 < g[1] < 3.0 and g[3] > 1.5)
    first_name = next(sp for sp in shapes if text(sp) == "patient_id")
    first_type = next(sp for sp in shapes if text(sp) == "FK")
    row_y = geometry(first_name)[1]
    for sp in shapes:
        g = geometry(sp)
        if g and g[0] < 4.2 and row_y - 0.01 <= g[1] < row_y + 1.2 and text(sp):
            move(tree, shape_id(sp), y=g[1] + ROW_SHIFT)
    for src, label in ((first_name, "admission_id"), (first_type, "PK")):
        new = copy.deepcopy(src)
        new.find(".//p:cNvPr", NS).set("id", str(next_id(tree)))
        new.find(".//p:cNvPr", NS).set("name", f"Input row {label}")
        new.find(".//a:t", NS).text = label
        src.addprevious(new)
        move(tree, shape_id(new), y=row_y)
    for sp in (card, table):
        move(tree, shape_id(sp), h=geometry(sp)[3] + ROW_SHIFT)

    # "What this enables" (last slide only) moves below the taller input card.
    enables = [sp for sp in shapes if (g := geometry(sp)) and g[0] < 1.2 and g[1] > 5.0]
    for sp in enables:
        g = geometry(sp)
        move(tree, shape_id(sp), y=g[1] + 0.30, h=g[3] - 0.15 if g[3] > 1.5 else None)

    # Target panel and code box: taller, same geometry on every slide.
    panel = find(shapes, lambda g, sp: 5.0 < g[0] < 5.4 and g[2] > 7 and g[3] > 4)
    move(tree, shape_id(panel), h=PANEL_BOTTOM - geometry(panel)[1])
    box = find(shapes, lambda g, sp: 5.3 < g[0] < 5.6 and g[2] > 6.5 and g[3] > 2.5)
    bx, _, bw, _ = geometry(box)
    for sp in shapes:  # old code lines
        g = geometry(sp)
        if g and sp is not box and bx - 0.01 <= g[0] and g[0] + g[2] <= bx + bw + 0.01 and g[1] >= CODE_TOP - 0.2:
            root.remove(sp)
    move(tree, shape_id(box), y=CODE_TOP, h=CODE_BOTTOM - CODE_TOP)
    if step == 0:
        return

    pitch = (CODE_BOTTOM - CODE_TOP - 2 * ROW_PAD) / len(DIFF_ROWS)
    y = CODE_TOP + ROW_PAD
    for row in DIFF_ROWS:
        if row_step(row) <= step:
            line = code_line(next_id(tree), bx, y, bw - 1.9, pitch, row[1], "", None, size=SIZE)
            if not final and highlighted(row, step):
                add_highlight(line)
            add(tree, line)
            if g := gloss(row):
                color = GLOSS_COLORS[row[3] or "schema"]
                add(tree, simple_box(next_id(tree), "Gloss", bx + bw - 1.97, y, 1.75, pitch,
                                     paras=para(run_xml("← " + g, color, False, GLOSS_SIZE, font="Calibri",
                                                        italic=True), algn="r")))
        y += pitch


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    pres = etree.parse(str(work / "ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target")
            for r in etree.parse(str(work / "ppt/_rels/presentation.xml.rels")).getroot()}
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    order = [rels[s.get(f"{{{R}}}id")] for s in pres.find("p:sldIdLst", NS)]
    run = [t for t in order if TITLE in (work / "ppt" / t).read_text(encoding="utf-8")]
    if len(run) != 6:
        raise SystemExit(f"expected 6 '{TITLE}' slides, found {len(run)}")
    for k, target in enumerate(run):
        path = work / "ppt" / target
        tree = etree.parse(str(path))
        edit(tree, step=min(k, 4), final=k == 5)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print("wrote", dst, "slides:", ", ".join(run))


if __name__ == "__main__":
    main(*sys.argv[1:3])
