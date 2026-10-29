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

CARD_TOP, CARD_BOTTOM = 2.40, 5.92
SHIFT = 0.62  # everything in the cards moves down to make room for the validation strip
# Synthetic rows over two tables joined by admission_id (no names; patients P1-P3, admissions A1-A3).
TABLES = [
    ("treatment", ("admission_id", "drug"), (1.15, 1.65),
     [("A1", "cisplatin"), ("A2", "nivolumab"), ("A3", "cisplatin")], "admission_id"),
    ("admission", ("admission_id", "patient_id", "BM_DEP"), (1.15, 0.95, 0.70),
     [("A1", "P1", "1"), ("A2", "P2", "0"), ("A3", "P3", "1")], None),
]
JOINED = {"A1", "A3"}  # cisplatin admissions, highlighted in both tables
BAND, TABLE_FILL = "D7EDE2", "E9EEEB"


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
    move(tree, shape_id(strip), y=6.08)
    move(tree, shape_id(footer), y=6.68)

    # Subtitle as the author edited it: key terms in bold, RIGOR bold italic.
    subtitle = by_text(shapes, "Today RIGOR")
    set_runs(subtitle, "".join(r(t, "3A434C" if b else MUTED, 1700, b, italic=i) for t, b, i in (
        ("Today ", False, False), ("RIGOR", True, True), (" reads ", False, False), ("schemas", True, False),
        (" and ", False, False), ("documentation", True, False), (". ", False, False),
        ("Instance data", True, False), (" could propose ", False, False), ("new knowledge", True, False),
        (".", False, False))))

    # Card 1: the instance data spans two tables, drawn like admission / PatientData on slide 2.
    cx, _, cw, _ = geometry(cards[0])
    set_runs(by_text(shapes, "What patient rows show"), r("What the rows show, across two tables", MUTED, 1300))
    root.remove(caption1)  # the footer already says the rows are synthetic
    bx, bw = cx + 0.18, cw - 0.36
    y = 3.22
    for name, cols, widths, rows, fk in TABLES:
        h = 0.32 + 0.22 + len(rows) * 0.205 + 0.07
        add(tree, simple_box(ids(), f"Table {name}", bx, y, bw, h, fill=TABLE_FILL, geom="roundRect", shadow=True))
        add(tree, shape(ids(), "Table name", bx + 0.15, y + 0.03, bw - 0.3, 0.27, r(name, INK, 1450, True),
                        anchor="ctr"))
        add(tree, shape(ids(), "Rule", bx + 0.15, y + 0.32, bw - 0.30, 0.012, fill="9AA3A9"))
        x = bx + 0.15
        for col, w in zip(cols, widths):
            label = r(col, EMERALD, 1050, True) + (r("  FK", MUTED, 900) if col == fk else "")
            add(tree, shape(ids(), "Column", x, y + 0.35, w, 0.20, label, anchor="ctr"))
            x += w
        for k, row in enumerate(rows):
            ry = y + 0.56 + k * 0.205
            if row[0] in JOINED:
                add(tree, shape(ids(), "Row band", bx + 0.08, ry, bw - 0.16, 0.19, fill=BAND))
            x = bx + 0.15
            for cell, w in zip(row, widths):
                add(tree, shape(ids(), "Cell", x, ry, w, 0.19, r(cell, INK, 1050, False, "Consolas"), anchor="ctr"))
                x += w
        y += h + 0.10

    # Card 2: one bridging line where the timing note was.
    x2, _, w2, _ = geometry(cards[1])
    add(tree, shape(ids(), "Bridge", x2 + 0.22, 4.95, w2 - 0.40, 0.62,
                    r("Data and label point the same way: a link worth checking.", BODY, 1200)))

    # Card 3: candidate as fixed per-line text boxes inside a code panel, note below it.
    x3, _, w3, _ = geometry(cards[2])
    px, py, pw = x3 + 0.15, 3.28, w3 - 0.30
    lines = FUTURE.splitlines()
    pitch = 0.19
    ph = len(lines) * pitch + 0.16
    add(tree, simple_box(ids(), "Code panel", px, py, pw, ph, fill="F6F8FA", line="D0D7DE", geom="roundRect"))
    for k, line in enumerate(lines):
        add(tree, code_line(ids(), px, py + 0.08 + k * pitch, pw - 0.05, pitch, line, "", None, size=1050))
    move(tree, shape_id(note), y=py + ph + 0.10, h=0.42)
    set_runs(note, r("No causal claim: ", RED, 1100, True)
             + r("gemcitabine, radiotherapy or cirrhosis can also suppress the marrow.", BODY, 1100))

    # Next steps in order: validation first (strip under the title), then the data (the three cards).
    for sp in tree.find(".//p:spTree", NS):
        g = geometry(sp) if sp.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp") else None
        if g and 2.3 <= g[1] < 6.0 and sp is not strip:
            move(tree, shape_id(sp), y=g[1] + SHIFT)
    title = by_text(shapes, "Next:")
    ts = list(title.iter(f"{{{NS['a']}}}t"))
    ts[0].text = "Next: check inside the loop, then let the data speak."
    for t in ts[1:]:
        t.getparent().getparent().remove(t.getparent())
    move(tree, shape_id(strip), y=1.82, h=0.60)
    set_runs(strip, r("FIRST, VALIDATORS IN THE LOOP   ", EMERALD, 1250, True)
             + r("Syntax, consistency and pitfall checks feed back to the Gen-LLM, as in ", BODY, 1250)
             + r("NeOn-GPT", INK, 1250, True) + r(" [Fathallah et al., ESWC 2024] and ", MUTED, 1250)
             + r("Extended NeOn-GPT", INK, 1250, True) + r(" [Fathallah et al., SWJ 2026]. ", MUTED, 1250)
             + r("Now running in the ", BODY, 1250)
             + r("RIGOR App", INK, 1250, True) + r(" [Fathallah, Nayyeri & Staab, ISWC 2026 Demos].", MUTED, 1250))
    move(tree, shape_id(subtitle), y=2.55)
    set_runs(subtitle, r("THEN, LET THE DATA SPEAK   ", EMERALD, 1300, True)
             + "".join(r(t, "3A434C" if b else MUTED, 1500, b, italic=i) for t, b, i in (
                 ("Today ", False, False), ("RIGOR", True, True), (" reads ", False, False), ("schemas", True, False),
                 (" and ", False, False), ("documentation", True, False), (". ", False, False),
                 ("Instance data", True, False), (" could propose ", False, False), ("new knowledge", True, False),
                 (".", False, False))))
    add(tree, shape(ids(), "Privacy", 0.90, CARD_BOTTOM + SHIFT + 0.08, 11.5, 0.28,
                    r("Only aggregate statistics would leave the hospital: ", EMERALD, 1300, True)
                    + r("patient data stays where it is.", BODY, 1300), anchor="ctr"))
    move(tree, shape_id(footer), y=7.02)


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
