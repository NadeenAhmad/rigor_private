#!/usr/bin/env python3
"""Future-work slide: close the loop, then let the data explain a cryptic column.

On the author's edited deck (the slide whose eyebrow starts with FUTURE WORK):
  - title "Next: close the loop, and let the data speak.";
  - strip: the evaluation tools (OOPS!, the reasoner) move into the loop, live in the RIGOR App demo;
  - three cards, data -> clues -> answer: the two instance tables with a "?" on BM_DEP (the question
    from the motivation slide), two clues read from the data (a 0/1 flag; 1 exactly with cisplatin),
    and the proposed label with provenance (code_excerpts.FUTURE, verified with the OWL API);
  - citations in the footer (paper Sec. 5; NeOn-GPT, Extended NeOn-GPT, RIGOR App).
Speaker notes come from speech.py. Writes a deck with only this slide.

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
LINE, STRIP, BAND, TABLE_FILL, AMBER = "D9DEDB", "E6F2EC", "D7EDE2", "E9EEEB", "93600F"
ICON = "image3.png"   # the graph-validation icon from the author's architecture figure
TABLES = [
    ("treatment", ("admission_id", "drug"), (1.15, 1.60),
     [("A1", "cisplatin"), ("A2", "nivolumab"), ("A3", "cisplatin")], "admission_id"),
    ("admission", ("admission_id", "patient_id", "BM_DEP"), (1.15, 0.95, 0.80),
     [("A1", "P1", "1"), ("A2", "P2", "0"), ("A3", "P3", "1")], None),
]
JOINED = {"A1", "A3"}
CARDS = ((0.88, 3.55), (4.79, 3.45), (8.60, 3.85))   # x, width
CARD_TOP, CARD_H = 3.00, 3.55


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


def magnifier(ids, x, y, size=0.34):
    """A plain magnifier in the deck's emerald: a ring and a 45-degree handle."""
    ring = raw(f'<p:sp><p:nvSpPr><p:cNvPr id="{ids()}" name="Magnifier ring"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr>{xfrm(x, y, size * 0.72, size * 0.72)}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
               f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="28575"><a:solidFill>'
               f'<a:srgbClr val="{EMERALD}"/></a:solidFill></a:ln></p:spPr></p:sp>')
    hx, hy = x + size * 0.62, y + size * 0.52
    handle = raw(f'<p:sp><p:nvSpPr><p:cNvPr id="{ids()}" name="Magnifier handle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                 f'<p:spPr>{xfrm(hx, hy, size * 0.12, size * 0.40, rot=-2700000)}<a:prstGeom prst="roundRect">'
                 f'<a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{EMERALD}"/></a:solidFill><a:ln><a:noFill/>'
                 f'</a:ln></p:spPr></p:sp>')
    return [ring, handle]


def picture(ids, rid, x, y, w, h, descr):
    return raw(f'<p:pic><p:nvPicPr><p:cNvPr id="{ids()}" name="Icon" descr="{descr}"/><p:cNvPicPr>'
               f'<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill>'
               f'<a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{xfrm(x, y, w, h)}'
               f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')


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
        h = 0.32 + 0.22 + len(rows) * 0.205 + 0.07
        add(tree, simple_box(ids(), f"Table {name}", bx, y, bw, h, fill=TABLE_FILL, geom="roundRect", shadow=True))
        add(tree, shape(ids(), "Table name", bx + 0.15, y + 0.03, bw - 0.3, 0.27, r(name, INK, 1450, True),
                        anchor="ctr"))
        add(tree, shape(ids(), "Rule", bx + 0.15, y + 0.32, bw - 0.30, 0.012, fill="9AA3A9"))
        cx = bx + 0.15
        for col, cw in zip(cols, widths):
            label = r(col, EMERALD, 1050, True) + (r("  FK", MUTED, 900) if col == fk else "")
            add(tree, shape(ids(), "Column", cx, y + 0.35, cw, 0.20, label, anchor="ctr"))
            if col == "BM_DEP":   # the open question from the motivation slide
                add(tree, shape(ids(), "Question", cx + 0.58, y + 0.34, 0.22, 0.22, r("?", "FFFFFF", 1100, True),
                                geom="ellipse", fill=AMBER, align="ctr", anchor="ctr"))
            cx += cw
        for k, row in enumerate(rows):
            ry = y + 0.56 + k * 0.205
            if row[0] in JOINED:
                add(tree, shape(ids(), "Row band", bx + 0.08, ry, bw - 0.16, 0.19, fill=BAND))
            cx = bx + 0.15
            for cell, cw in zip(row, widths):
                add(tree, shape(ids(), "Cell", cx, ry, cw, 0.19, r(cell, INK, 1050, False, "Consolas"), anchor="ctr"))
                cx += cw
        y += h + 0.10


def clues(tree, ids, x, w):
    y = CARD_TOP + 0.62
    for k, (head, body) in enumerate((
            ("Only 0 and 1", "A yes/no flag, not a score."),
            ("1 exactly when cisplatin was given", "Cisplatin is a chemotherapy known to suppress the bone marrow."))):
        cy = y + k * 1.42
        add(tree, shape(ids(), "Clue", x + 0.15, cy, w - 0.30, 1.30, geom="roundRect", fill="F6F8FA", line=LINE))
        for el in magnifier(ids, x + 0.28, cy + 0.14):
            add(tree, el)
        add(tree, box(ids, x + 0.72, cy + 0.10, w - 1.0, 0.50, r(head, INK, 1400, True), anchor="ctr", name="Clue head"))
        add(tree, box(ids, x + 0.30, cy + 0.66, w - 1.25 if k == 0 else w - 0.6, 0.58, r(body, BODY, 1200),
                      name="Clue body"))
        if k == 0:   # BM_DEP values in the admission rows: one 0, two 1s
            base = cy + 1.12
            for j, (label, count) in enumerate((("0", 1), ("1", 2))):
                bx = x + w - 0.85 + j * 0.30
                add(tree, shape(ids(), "Value bar", bx, base - 0.20 * count, 0.20, 0.20 * count,
                                fill=EMERALD if label == "1" else "9AA3A9"))
                add(tree, box(ids, bx - 0.05, base + 0.01, 0.30, 0.16, r(label, MUTED, 900, False, "Consolas"),
                              align="ctr", name="Value label"))


def answer(tree, ids, x, w):
    add(tree, box(ids, x + 0.25, CARD_TOP + 0.62, w - 0.5, 0.34,
                  r("BM_DEP", EMERALD, 1700, True, "Consolas") + r("   →", MUTED, 1700), name="Reveal key"))
    add(tree, box(ids, x + 0.25, CARD_TOP + 0.98, w - 0.5, 0.42, r("Bone marrow depression", INK, 2000, True,
                                                                   "Cambria"), name="Reveal"))
    lines = FUTURE.splitlines()
    px, py, pw, pitch = x + 0.15, CARD_TOP + 1.55, w - 0.30, 0.19
    add(tree, simple_box(ids(), "Code panel", px, py, pw, len(lines) * pitch + 0.14, fill="F6F8FA", line="D0D7DE",
                         geom="roundRect"))
    for k, line in enumerate(lines):
        add(tree, code_line(ids(), px, py + 0.07 + k * pitch, pw - 0.05, pitch, line, "", None, size=1000))
    add(tree, shape(ids(), "Candidate", x + 0.15, CARD_TOP + 2.92, w - 0.30, 0.46,
                    r("CANDIDATE  ", AMBER, 1150, True, "Consolas") + r("an expert confirms it first", BODY, 1200),
                    geom="roundRect", fill="FFF7E8", align="ctr", anchor="ctr"))


def edit(tree, icon_rid):
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
    ts[0].text = "Next: close the loop, and let the data speak."
    for t in ts[1:]:
        t.getparent().getparent().remove(t.getparent())
    ids = Ids(tree)

    # First: the evaluation tools move into the loop.
    add(tree, shape(ids(), "Loop strip", 0.88, 1.80, 11.57, 0.62, geom="roundRect", fill=STRIP))
    add(tree, picture(ids, icon_rid, 1.02, 1.86, 0.50, 0.50, "Graph validation icon"))
    add(tree, shape(ids(), "Loop text", 1.62, 1.80, 10.70, 0.62,
                    r("FIRST, MOVE THE EVALUATION TOOLS INTO THE LOOP   ", EMERALD, 1250, True)
                    + r("OOPS! and the reasoner only checked our final ontologies. Next, they flag each error while "
                        "RIGOR builds, so the generator fixes it before merging.  ", BODY, 1250)
                    + r("Live in our RIGOR App demo.", INK, 1250, True, italic=True), anchor="ctr"))
    # Then: the data explains a cryptic column.
    add(tree, box(ids, 0.90, 2.56, 11.5, 0.34,
                  r("THEN, LET THE DATA SPEAK   ", EMERALD, 1300, True)
                  + r("When documentation is missing, instance data can explain ", BODY, 1450)
                  + r("cryptic columns", INK, 1450, True) + r(".", BODY, 1450), anchor="ctr", name="Data lead"))
    (x1, w1), (x2, w2), (x3, w3) = CARDS
    card(tree, ids, x1, w1, "1  THE DATA")
    add(tree, box(ids, x1 + 0.25, CARD_TOP + 0.48, w1 - 0.5, 0.28,
                  r("What does ", BODY, 1250, italic=True) + r("BM_DEP", AMBER, 1250, True, "Consolas")
                  + r(" mean?", BODY, 1250, italic=True), name="Question"))
    tables(tree, ids, x1, w1)
    card(tree, ids, x2, w2, "2  THE CLUES")
    clues(tree, ids, x2, w2)
    card(tree, ids, x3, w3, "3  THE ANSWER")
    answer(tree, ids, x3, w3)
    mid = CARD_TOP + CARD_H / 2 - 0.18
    add(tree, arrow(ids, x1 + w1 + 0.03, mid))
    add(tree, arrow(ids, x2 + w2 + 0.03, mid))
    add(tree, box(ids, 0.90, 6.66, 11.6, 0.42,
                  r("Rows are synthetic. Instance-level signals for cryptic columns: paper Sec. 5. Tool feedback in the "
                    "loop: NeOn-GPT [Fathallah et al., ESWC 2024]; Extended NeOn-GPT [Fathallah et al., SWJ 2026]; "
                    "RIGOR App [Fathallah, Nayyeri & Staab, ISWC 2026 Demos].", MUTED, 1000), name="Source"))


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
    if not (work / "ppt/media" / ICON).exists():
        raise SystemExit(f"icon {ICON} not found in the deck")
    rels_path = work / "ppt/slides/_rels" / f"{target}.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    if "rIdIcon" not in rels_xml:
        rels_xml = rels_xml.replace("</Relationships>", '<Relationship Id="rIdIcon" Type="http://schemas.openxmlformats'
                                    f'.org/officeDocument/2006/relationships/image" Target="../media/{ICON}"/>'
                                    "</Relationships>")
        rels_path.write_text(rels_xml, encoding="utf-8")
    path = work / "ppt/slides" / target
    tree = etree.parse(str(path))
    edit(tree, "rIdIcon")
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
