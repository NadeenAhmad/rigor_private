#!/usr/bin/env python3
"""Future-work slide: instance data can propose new knowledge.

On the author's edited deck (the slide whose eyebrow starts with FUTURE WORK; the evaluation-loop step has
its own slide): the data shows a link the documentation does not contain (every cisplatin admission has
bone marrow depression), and RIGOR proposes it as a candidate association for expert review.
Three cards: the data (two tables joined on admission_id) -> the pattern -> the candidate knowledge
(code_excerpts.FUTURE, verified with the OWL API). Speaker notes come from speech.py.

Usage: python3 future_v3.py <in.pptx> <out.pptx>
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import A, FUTURE, NS, P, code_line, run_xml, simple_box  # noqa: E402
from results import BODY, EMERALD, INK, MUTED, TEAL, Ids, add, set_notes, shape  # noqa: E402
from slide9_diff import text  # noqa: E402
from speech import notes_for, timeline  # noqa: E402

EMU = 914400
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
LINE, BAND, TABLE_FILL, AMBER = "D9DEDB", "D7EDE2", "E9EEEB", "93600F"
TABLES = [
    ("treatment", ("admission_id", "drug"), (1.15, 1.60),
     [("A1", "cisplatin"), ("A2", "nivolumab"), ("A3", "cisplatin")], "admission_id"),
    ("admission", ("admission_id", "patient_id", "BM_DEP"), (1.15, 0.95, 0.80),
     [("A1", "P1", "1"), ("A2", "P2", "0"), ("A3", "P3", "1")], None),
]
JOINED = {"A1", "A3"}
CARDS = ((0.88, 3.55), (4.79, 3.30), (8.45, 4.00))   # x, width
CARD_TOP, CARD_H = 2.52, 3.90


def r(t, color=INK, size=1200, bold=False, font="Calibri", italic=False):
    return run_xml(t, color, bold, size, font=font, italic=italic)


def box(ids, x, y, w, h, runs, align="l", anchor="t", name="Text"):
    return shape(ids(), name, x, y, w, h, runs, align=align, anchor=anchor)


def raw(xml):
    return etree.fromstring(f'<x xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R_NS}">{xml}</x>')[0]


def xfrm(x, y, w, h, rot=0):
    rot_attr = f' rot="{rot}"' if rot else ""
    return (f'<a:xfrm{rot_attr}><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
            f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>')


def arrow(ids, x, y):
    return raw(f'<p:sp><p:nvSpPr><p:cNvPr id="{ids()}" name="Arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr>{xfrm(x, y, 0.30, 0.36)}<a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>'
               f'<a:solidFill><a:srgbClr val="{EMERALD}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>')


def card(tree, ids, x, w, label, note=""):
    add(tree, shape(ids(), "Card", x, CARD_TOP, w, CARD_H, geom="roundRect", fill="FFFFFF", line=LINE))
    add(tree, box(ids, x + 0.25, CARD_TOP + 0.18, w - 0.5, 0.30, r(label, TEAL, 1300, True, "Consolas"),
                  name="Card label"))
    if note:
        add(tree, box(ids, x + 0.25, CARD_TOP + 0.48, w - 0.5, 0.26, r(note, MUTED, 1150, italic=True), name="Note"))


def tables(tree, ids, x, w):
    bx, bw, y = x + 0.15, w - 0.30, CARD_TOP + 0.82
    for name, cols, widths, rows, fk in TABLES:
        h = 0.36 + 0.26 + len(rows) * 0.245 + 0.08
        add(tree, simple_box(ids(), f"Table {name}", bx, y, bw, h, fill=TABLE_FILL, geom="roundRect", shadow=True))
        add(tree, shape(ids(), "Table name", bx + 0.15, y + 0.03, bw - 0.3, 0.30, r(name, INK, 1600, True),
                        anchor="ctr"))
        add(tree, shape(ids(), "Rule", bx + 0.15, y + 0.36, bw - 0.30, 0.012, fill="9AA3A9"))
        cx = bx + 0.15
        for col, cw in zip(cols, widths):
            label = r(col, EMERALD, 1150, True) + (r(" FK", MUTED, 950) if col == fk else "")
            add(tree, shape(ids(), "Column", cx, y + 0.39, cw, 0.24, label, anchor="ctr"))
            cx += cw
        for k, row in enumerate(rows):
            ry = y + 0.64 + k * 0.245
            if row[0] in JOINED:
                add(tree, shape(ids(), "Row band", bx + 0.08, ry, bw - 0.16, 0.225, fill=BAND))
            cx = bx + 0.15
            for cell, cw in zip(row, widths):
                add(tree, shape(ids(), "Cell", cx, ry, cw, 0.225, r(cell, INK, 1250, False, "Consolas"), anchor="ctr"))
                cx += cw
        y += h + 0.10


def pattern(tree, ids, x, w):
    y = CARD_TOP + 0.85
    add(tree, box(ids, x + 0.25, y, w - 0.5, 0.80,
                  r("Every cisplatin admission has bone marrow depression.", INK, 1600, True), name="Pattern"))
    y += 1.00
    for label, hits, total in (("cisplatin", 2, 2), ("other drugs", 0, 1)):
        add(tree, box(ids, x + 0.25, y, 1.30, 0.34, r(label, INK, 1300, False, "Consolas"), anchor="ctr",
                      name="Group"))
        add(tree, shape(ids(), "Track", x + 1.55, y + 0.10, 0.85, 0.15, fill="E1E6E3"))
        if hits:
            add(tree, shape(ids(), "Bar", x + 1.55, y + 0.10, 0.85 * hits / total, 0.15, fill=EMERALD))
        add(tree, box(ids, x + 2.45, y, 0.75, 0.34, r(f"{hits} of {total}", EMERALD if hits else BODY, 1300, True),
                      anchor="ctr", name="Count"))
        y += 0.44
    add(tree, box(ids, x + 0.25, y, w - 0.5, 0.26, r("admissions with BM_DEP = 1", MUTED, 1100, italic=True),
                  name="Axis note"))
    add(tree, shape(ids(), "Not documented", x + 0.15, CARD_TOP + 3.06, w - 0.30, 0.70,
                    r("Not in the documentation: ", EMERALD, 1300, True)
                    + r("a candidate for new knowledge.", BODY, 1300),
                    geom="roundRect", fill="E8F6EF", align="ctr", anchor="ctr", inset=0.12))


def candidate(tree, ids, x, w):
    add(tree, box(ids, x + 0.25, CARD_TOP + 0.48, w - 0.5, 0.28, r("A reviewable hypothesis, not a fact", MUTED, 1250,
                                                                   italic=True), name="Note"))
    lines = FUTURE.splitlines()
    px, py, pw, pitch = x + 0.15, CARD_TOP + 0.88, w - 0.30, 0.225
    add(tree, simple_box(ids(), "Code panel", px, py, pw, len(lines) * pitch + 0.14, fill="F6F8FA", line="D0D7DE",
                         geom="roundRect"))
    for k, line in enumerate(lines):
        add(tree, code_line(ids(), px, py + 0.07 + k * pitch, pw - 0.05, pitch, line, "", None, size=1100))
    add(tree, shape(ids(), "Candidate", x + 0.15, CARD_TOP + 3.16, w - 0.30, 0.60,
                    r("A PATTERN, NOT A CAUSE  ", AMBER, 1200, True, "Consolas")
                    + r("an expert reviews it before it joins the ontology", BODY, 1250),
                    geom="roundRect", fill="FFF7E8", align="ctr", anchor="ctr", inset=0.10))


def edit(tree):
    root = tree.find(".//p:spTree", NS)
    shapes = [el for el in root if el.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp", f"{{{P}}}pic", f"{{{P}}}grpSp")]
    eyebrow = next(sp for sp in shapes if text(sp).startswith("FUTURE WORK"))
    title = shapes[shapes.index(eyebrow) + 1]
    for sp in shapes:
        if sp is not eyebrow and sp is not title:
            root.remove(sp)
    for t in list(eyebrow.iter(f"{{{A}}}t"))[1:]:
        t.getparent().getparent().remove(t.getparent())
    list(eyebrow.iter(f"{{{A}}}t"))[0].text = "FUTURE WORK"
    ts = list(title.iter(f"{{{A}}}t"))
    ts[0].text = "Next: let the data speak."
    for t in ts[1:]:
        t.getparent().getparent().remove(t.getparent())
    ids = Ids(tree)
    add(tree, box(ids, 0.90, 1.84, 11.5, 0.40, "".join(r(t, "3A434C" if b else MUTED, 1650, b, italic=i) for t, b, i in (
        ("Today ", False, False), ("RIGOR", True, True), (" reads ", False, False), ("schemas", True, False),
        (" and ", False, False), ("documentation", True, False), (". ", False, False), ("Instance data", True, False),
        (" can reveal links the documentation does not contain, and propose ", False, False),
        ("new knowledge", True, False), (".", False, False))), anchor="ctr", name="Subtitle"))
    (x1, w1), (x2, w2), (x3, w3) = CARDS
    card(tree, ids, x1, w1, "1  THE DATA", "two tables, joined on admission_id")
    tables(tree, ids, x1, w1)
    card(tree, ids, x2, w2, "2  THE PATTERN", "what the joined rows show")
    pattern(tree, ids, x2, w2)
    card(tree, ids, x3, w3, "3  CANDIDATE KNOWLEDGE")
    candidate(tree, ids, x3, w3)
    mid = CARD_TOP + CARD_H / 2 - 0.18
    add(tree, arrow(ids, x1 + w1 + 0.03, mid))
    add(tree, arrow(ids, x2 + w2 + 0.03, mid))
    add(tree, box(ids, 0.90, 6.56, 11.6, 0.26,
                  r("Rows are synthetic. In practice, RIGOR would use aggregate statistics only, so patient data stays "
                    "in the hospital.", MUTED, 1000), name="Source"))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    pres_path = work / "ppt/presentation.xml"
    pres = etree.parse(str(pres_path))
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
                           (work / "ppt/_rels/presentation.xml.rels").read_text()))
    lst = pres.find("p:sldIdLst", NS)
    target = None
    for sld in list(lst):
        name = rels[sld.get(f"{{{R_NS}}}id")]
        if target is None and "FUTURE WORK" in (work / "ppt/slides" / name).read_text(encoding="utf-8"):
            target = name
        else:
            lst.remove(sld)
    for sec in pres.iter("{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst"):
        sec.getparent().remove(sec)
    pres.write(str(pres_path), xml_declaration=True, encoding="UTF-8", standalone=True)
    rels_xml = (work / "ppt/slides/_rels" / f"{target}.rels").read_text(encoding="utf-8")
    path = work / "ppt/slides" / target
    tree = etree.parse(str(path))
    edit(tree)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    notes = re.search(r"notesSlides/(notesSlide\d+\.xml)", rels_xml)
    if notes:
        n, s, start, end, _ = next(row for row in timeline() if row[1]["title"] == "Future work")
        set_notes(work / "ppt/notesSlides" / notes.group(1), notes_for(n, s, start, end))
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
