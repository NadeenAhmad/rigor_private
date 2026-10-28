#!/usr/bin/env python3
"""Future-work slide on the author's edited deck, without tab stops (Google Slides drops them).

  - Card 1: the instance table is one text box per cell, so columns stay aligned; header "BM_DEP".
  - Card 2: the timing note goes; one line says what data and label suggest together.
  - Card 3: the candidate is one fixed text box per code line (code_excerpts.FUTURE, verified with the
    OWL API), so nothing overflows; the note sits below the code with its own space.
  - Cards are taller; the "Also next" strip and the footer move down to make room.

Usage: python3 future_v2.py <in.pptx> <out.pptx>   (writes a deck with only this slide)
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import FUTURE, NS, P, code_line, move, run_xml, simple_box  # noqa: E402
from results import BODY, EMERALD, INK, MUTED, RED, Ids, add, shape  # noqa: E402
from slide9_diff import geometry, shape_id, text  # noqa: E402

CARD_TOP, CARD_BOTTOM = 2.40, 5.82
ROWS = [("P01", "yes", "1"), ("P02", "yes", "1"), ("P03", "no", "0"), ("P04", "yes", "1"), ("P05", "no", "0")]
BAND = "E6F2EC"


def r(t, color=INK, size=1200, bold=False, font="Calibri", italic=False):
    return run_xml(t, color, bold, size, font=font, italic=italic)


def by_text(shapes, start):
    return next(sp for sp in shapes if text(sp).startswith(start))


def set_runs(sp, runs_xml):
    """Replace a shape's paragraphs with one paragraph of the given runs, keeping its paragraph props."""
    body = sp.find("p:txBody", NS)
    paras = body.findall("a:p", NS)
    ppr = paras[0].find("a:pPr", NS)
    for p in paras:
        body.remove(p)
    p = etree.SubElement(body, f"{{{NS['a']}}}p")
    if ppr is not None:
        p.append(ppr)
    for run in etree.fromstring(f'<x xmlns:a="{NS["a"]}">{runs_xml}</x>'):
        p.append(run)


def edit(tree):
    root = tree.find(".//p:spTree", NS)
    shapes = [el for el in root if el.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp")]
    ids = Ids(tree)
    cards = sorted((sp for sp in shapes if (g := geometry(sp)) and g[3] > 3 and 2.2 < g[1] < 2.6),
                   key=lambda sp: geometry(sp)[0])
    arrows = [sp for sp in shapes if (g := geometry(sp)) and g[2] < 0.6 and 0.3 < g[3] < 0.6 and 3.5 < g[1] < 4.2]
    table = [sp for sp in shapes if text(sp).startswith(("PATIENT", "P0"))]
    caption1 = by_text(shapes, "Synthetic illustration")
    timing = by_text(shapes, "Timing is evidence")
    code = by_text(shapes, "Individual:")
    note = by_text(shapes, "No causes axiom")
    strip = by_text(shapes, "ALSO NEXT")
    footer = by_text(shapes, "Instance rows are synthetic")
    for sp in table + [timing, code]:
        root.remove(sp)

    for card in cards:
        move(tree, shape_id(card), y=CARD_TOP, h=CARD_BOTTOM - CARD_TOP)
    for arrow in arrows:
        move(tree, shape_id(arrow), y=(CARD_TOP + CARD_BOTTOM) / 2 - geometry(arrow)[3] / 2)
    move(tree, shape_id(strip), y=6.00)
    move(tree, shape_id(footer), y=6.62)

    # Subtitle as the author edited it: key terms in bold, RIGOR bold italic.
    subtitle = by_text(shapes, "Today RIGOR")
    set_runs(subtitle, "".join(r(t, "3A434C" if b else MUTED, 1700, b, italic=i) for t, b, i in (
        ("Today ", False, False), ("RIGOR", True, True), (" reads ", False, False), ("schemas", True, False),
        (" and ", False, False), ("documentation", True, False), (". ", False, False),
        ("Instance data", True, False), (" could propose ", False, False), ("new knowledge", True, False),
        (".", False, False))))

    # Card 1: instance table, one text box per cell.
    cx, _, cw, _ = geometry(cards[0])
    cols = [cx + 0.22, cx + 1.20, cx + 2.30]
    y = 3.32
    for col, label in zip(cols, ("PATIENT", "CISPLATIN", "BM_DEP")):
        add(tree, shape(ids(), "Table header", col, y, 1.0, 0.26, r(label, MUTED, 950, True, "Consolas"),
                        anchor="ctr"))
    for k, row in enumerate(ROWS):
        ry = 3.64 + k * 0.32
        if row[1] == "yes":
            add(tree, shape(ids(), "Row band", cx + 0.15, ry, cw - 0.30, 0.28, fill=BAND))
        for col, cell in zip(cols, row):
            add(tree, shape(ids(), "Cell", col, ry, 1.0, 0.28, r(cell, INK, 1150, False, "Consolas"), anchor="ctr"))
    move(tree, shape_id(caption1), y=5.30, h=0.42)
    set_runs(caption1, r("Synthetic rows: patients given cisplatin show bone marrow depression (BM_DEP = 1).",
                         MUTED, 1100))

    # Card 2: one bridging line where the timing note was.
    x2, _, w2, _ = geometry(cards[1])
    add(tree, shape(ids(), "Bridge", x2 + 0.22, 4.95, w2 - 0.40, 0.62,
                    r("Data and label point the same way: a link worth checking.", BODY, 1200)))

    # Card 3: candidate as fixed per-line text boxes inside a code panel, note below it.
    x3, _, w3, _ = geometry(cards[2])
    px, py, pw = x3 + 0.15, 3.28, w3 - 0.30
    lines = FUTURE.splitlines()
    pitch = 0.205
    ph = len(lines) * pitch + 0.16
    add(tree, simple_box(ids(), "Code panel", px, py, pw, ph, fill="F6F8FA", line="D0D7DE", geom="roundRect"))
    for k, line in enumerate(lines):
        add(tree, code_line(ids(), px, py + 0.08 + k * pitch, pw - 0.05, pitch, line, "", None, size=1050))
    move(tree, shape_id(note), y=py + ph + 0.10, h=0.42)
    set_runs(note, r("No causal claim: ", RED, 1100, True)
             + r("gemcitabine, radiotherapy or cirrhosis can also suppress the marrow.", BODY, 1100))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    pres_path = work / "ppt/presentation.xml"
    pres = etree.parse(str(pres_path))
    rels = {m[0]: m[1] for m in re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
                                           (work / "ppt/_rels/presentation.xml.rels").read_text())}
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    lst = pres.find("p:sldIdLst", NS)
    target = None
    for sld in list(lst):
        name = rels[sld.get(f"{{{R}}}id")]
        if "FUTURE WORK" in (work / "ppt/slides" / name).read_text(encoding="utf-8"):
            target = name
        else:
            lst.remove(sld)
    for sec in pres.iter("{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst"):
        sec.getparent().remove(sec)
    pres.write(str(pres_path), xml_declaration=True, encoding="UTF-8", standalone=True)
    path = work / "ppt/slides" / target
    tree = etree.parse(str(path))
    edit(tree)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
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
