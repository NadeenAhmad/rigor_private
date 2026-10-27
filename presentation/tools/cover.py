#!/usr/bin/env python3
"""Cover slide: white left side (wordmark, title, authors, affiliations, links) and a
deep-emerald panel on the right third with the table-to-graph motif and the venue.

Usage: python3 cover.py <in.pptx> <out.pptx>
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from results import (A, BODY, EMERALD, INK, MUTED, NS, P, TEAL, Ids, delete, move, set_runs,  # noqa: E402
                     set_size, shape, sp_by_id)

EMU = 914400
PANEL_X, SLIDE_W, SLIDE_H = 8.85, 13.333, 7.5
PAPER, GRID = "FAFBFA", "E6ECE8"          # graph-paper panel
GOLD_TINT, GOLD_LINE = "FFF3D1", "E0B34A"  # links the BM_DEP column to its label

AUTHOR_LINES = [
    [("Nadeen Fathallah", "1", True), ("Mojtaba Nayyeri", "1", False), ("Athish A Yogi", "1", False)],
    [("Ratan Bahadur Thapa", "1", False), ("Hans-Michael Tautenhahn", "3", False)],
    [("Anton Schnurpel", "3", False), ("Steffen Staab", "1,2", False)],
]
AFFILIATIONS = [
    ("1", "Institute for Artificial Intelligence, University of Stuttgart, Germany"),
    ("2", "University of Southampton, United Kingdom"),
    ("3", "Universitätsklinikum Leipzig, Germany"),
]


def r(text, color, size, bold=False, font="Calibri", sup=False, italic=False, spc=None):
    base = ' baseline="30000"' if sup else ""
    space = f' spc="{spc}"' if spc else ""
    return (f'<a:r><a:rPr lang="en-US" sz="{size}" b="{int(bold)}" i="{int(italic)}"{base}{space} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:rPr>'
            f'<a:t xml:space="preserve">{escape(text)}</a:t></a:r>')


def text_box(sid, name, x, y, w, h, paragraphs, line_pts=None, anchor="t"):
    """paragraphs: list of run-XML strings, one per paragraph."""
    spacing = f'<a:lnSpc><a:spcPts val="{line_pts}"/></a:lnSpc>' if line_pts else ""
    paras = "".join(f'<a:p><a:pPr>{spacing}</a:pPr>{p}</a:p>' for p in paragraphs)
    return etree.fromstring(
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
        f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}">'
        f'<a:noAutofit/></a:bodyPr><a:lstStyle/>{paras}</p:txBody></p:sp>')


def edit_cover(tree):
    ids = Ids(tree)
    sp_tree = tree.find(".//p:spTree", NS)
    add = sp_tree.append

    # Right panel: near-white "graph paper", behind everything else.
    background = [shape(ids(), "Cover panel", PANEL_X, 0, SLIDE_W - PANEL_X, SLIDE_H, fill=PAPER)]
    step = 0.30
    gx = PANEL_X + step
    while gx < SLIDE_W - 0.01:
        background.append(shape(ids(), "Grid line", gx, 0, 0.007, SLIDE_H, fill=GRID))
        gx += step
    gy = step
    while gy < SLIDE_H - 0.01:
        background.append(shape(ids(), "Grid line", PANEL_X, gy, SLIDE_W - PANEL_X, 0.007, fill=GRID))
        gy += step
    for el in reversed(background):
        sp_tree.insert(2, el)

    # Left: eyebrow, wordmark, title.
    set_runs(tree, 21, [("RESEARCH PAPER", None, None)])
    move(tree, 21, x=0.87, y=0.62, w=7.4)
    add(text_box(ids(), "Wordmark", 0.84, 1.08, 7.4, 0.95, [r("RIGOR", EMERALD, 5400, True, "Cambria", spc=300)]))
    move(tree, 22, x=0.87, y=2.12, w=7.6, h=1.62)
    set_size(tree, 22, 3000)

    # Authors: three balanced lines, true superscripts, presenter in bold.
    lines = []
    for line in AUTHOR_LINES:
        parts = []
        for i, (name, aff, presenter) in enumerate(line):
            if i:
                parts.append(r("   ·   ", "6B747B", 1500))
            parts.append(r(name, INK if presenter else BODY, 1500, presenter))
            parts.append(r(aff, EMERALD, 1500, True, sup=True))
        lines.append("".join(parts))
    delete(tree, 23, 24)
    add(text_box(ids(), "Authors", 0.87, 4.18, 7.6, 1.05, lines, line_pts=2300))
    add(text_box(ids(), "Affiliations", 0.87, 5.38, 7.6, 0.80,
                 [r(n, EMERALD, 1200, True, sup=True) + r(" " + a, MUTED, 1200) for n, a in AFFILIATIONS],
                 line_pts=1700))

    # Links footer.
    set_runs(tree, 25, [("PAPER  ", MUTED, True), ("arXiv:2506.01232", TEAL, False),
                        ("        CODE  ", MUTED, True), ("github.com/NadeenAhmad/RIGOR", TEAL, False)])
    move(tree, 25, x=0.87, y=6.62, w=7.6, h=0.30)
    set_size(tree, 25, 1250)

    # Panel content: the running example, drawn as vectors (no bitmap).
    pic = next(p for p in tree.iter(f"{{{P}}}pic"))
    pic.getparent().remove(pic)
    delete(tree, 27, 28)
    diagram(tree, ids, add)
    add(text_box(ids(), "Venue", 9.35, 6.32, 3.4, 0.62,
                 [r("ISWC 2026", EMERALD, 1400, True, "Consolas"), r("Bari, Italy", INK, 1300)], line_pts=2000))


def line(sid, x1, y1, x2, y2, color, width_pt=1.25, arrow=False, dash=None):
    x, y, w, h = min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1)
    flip = (' flipH="1"' if x2 < x1 else "") + (' flipV="1"' if y2 < y1 else "")
    tail = '<a:tailEnd type="triangle" w="med" len="med"/>' if arrow else ""
    dash_xml = f'<a:prstDash val="{dash}"/>' if dash else ""
    return etree.fromstring(
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="{sid}" name="Line {sid}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{flip}><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
        f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        f'<a:ln w="{int(width_pt * 12700)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{dash_xml}{tail}</a:ln>'
        f'</p:spPr></p:sp>')


def diagram(tree, ids, add):
    """admission table -> RIGOR -> ontology fragment, with the BM_DEP row and its label in gold."""
    x0, w = 9.35, 3.50
    # Table card.
    top = 1.00
    add(shape(ids(), "Table card", x0, top, w, 1.68, geom="roundRect", fill="FFFFFF", line="C9D1CD",
              descr="relational table admission"))
    head = shape(ids(), "Table head", x0, top, w, 0.44, rr("admission", "FFFFFF", 1300, True, "Consolas"),
                 geom="round2SameRect", fill=INK, anchor="ctr", inset=0.16)
    add(head)
    rows = [("patient_id", "FK → PatientData", False), ("BM_DEP", "INT", True), ("ecog_stage", "INT", False)]
    for k, (col, typ, gold) in enumerate(rows):
        y = top + 0.50 + k * 0.38
        if gold:
            add(shape(ids(), "Gold row", x0 + 0.08, y, w - 0.16, 0.34, geom="roundRect", fill=GOLD_TINT,
                      line=GOLD_LINE))
        add(shape(ids(), "Column", x0 + 0.20, y, 1.8, 0.34, rr(col, INK, 1250, gold, "Consolas"), anchor="ctr"))
        add(shape(ids(), "Type", x0 + 1.70, y, w - 1.90, 0.34, rr(typ, MUTED, 1100, False, "Consolas"),
                  align="r", anchor="ctr"))
    # RIGOR step: the flow runs down the left, into the Admission class.
    ny, nh, nw = 3.78, 0.46, 1.36
    fx = x0 + nw / 2
    add(line(ids(), fx, top + 1.70, fx, 2.98, EMERALD, 1.5))
    add(shape(ids(), "RIGOR pill", fx - 0.62, 2.98, 1.24, 0.38, rr("RIGOR", "FFFFFF", 1300, True, "Consolas"),
              geom="roundRect", fill=EMERALD, align="ctr", anchor="ctr"))
    add(line(ids(), fx, 3.36, fx, ny - 0.03, EMERALD, 1.5, arrow=True))
    # Ontology: two classes and the relation.
    add(shape(ids(), "Class Admission", x0, ny, nw, nh, rr("Admission", "FFFFFF", 1350, True), geom="roundRect",
              fill=EMERALD, align="ctr", anchor="ctr", descr="ontology class"))
    pd = shape(ids(), "Class PatientData", x0 + w - nw, ny, nw, nh, rr("PatientData", EMERALD, 1350, True),
               geom="roundRect", fill="FFFFFF", line=EMERALD, align="ctr", anchor="ctr", descr="ontology class")
    pd.find(".//a:ln", NS).set("w", "19050")
    add(pd)
    add(line(ids(), x0 + nw + 0.02, ny + nh / 2, x0 + w - nw - 0.04, ny + nh / 2, TEAL, 1.25, arrow=True))
    gap_cx = x0 + w / 2
    add(no_wrap(shape(ids(), "Relation label", gap_cx - 0.95, ny - 0.30, 1.9, 0.24,
                      rr("hasPatientData", TEAL, 1050, True, "Consolas"), align="ctr")))
    # Annotations hanging off Admission.
    spine_x = x0 + 0.22
    chips = [('rdfs:label "Bone Marrow Depression"', True), ('rdfs:comment "ECOG scale 0–5"', False),
             ("prov:wasDerivedFrom src:ecog_stage", False)]
    first, gap = ny + nh + 0.26, 0.46
    add(line(ids(), spine_x, ny + nh, spine_x, first + (len(chips) - 1) * gap + 0.17, "8E989E", 1.0))
    for k, (text, gold) in enumerate(chips):
        y = first + k * gap
        add(line(ids(), spine_x, y + 0.17, spine_x + 0.20, y + 0.17, "8E989E", 1.0))
        width = 0.55 * 10.5 / 72 * len(text) + 0.30
        add(no_wrap(shape(ids(), "Annotation chip", spine_x + 0.20, y, width, 0.34, chip_runs(text),
                          geom="roundRect", fill=GOLD_TINT if gold else "FFFFFF",
                          line=GOLD_LINE if gold else "C9D1CD", anchor="ctr", inset=0.12, descr="annotation")))


def no_wrap(sp):
    sp.find(".//a:bodyPr", NS).set("wrap", "none")
    return sp


def rr(text, color, size, bold=False, font="Calibri"):
    return r(text, color, size, bold, font)


def chip_runs(text):
    """Colour a one-line annotation: key in slate, literal in brown, IRI in teal."""
    key, _, value = text.partition(" ")
    color = "A3401E" if value.startswith('"') else TEAL
    return r(key + " ", MUTED, 1050, False, "Consolas") + r(value, color, 1050, False, "Consolas")


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    slide = work / "ppt/slides/slide1.xml"
    tree = etree.parse(str(slide))
    edit_cover(tree)
    tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
    # The bitmap motif is replaced by the vector diagram: drop its relationship and file.
    rels = work / "ppt/slides/_rels/slide1.xml.rels"
    rels.write_text(re.sub(r'<Relationship [^>]*media/image1\.png"/>', "", rels.read_text()))
    (work / "ppt/media/image1.png").unlink(missing_ok=True)
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
