#!/usr/bin/env python3
"""Cover slide: white left side (wordmark, title, authors, affiliations, links) and a
deep-emerald panel on the right third with the table-to-graph motif and the venue.

Usage: python3 cover.py <in.pptx> <out.pptx>
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from restyle import TITLE_STYLES, recolor_motif  # noqa: E402
from results import (A, BODY, EMERALD, INK, MUTED, NS, P, TEAL, Ids, delete, move, set_runs,  # noqa: E402
                     set_size, shape, sp_by_id)

ROOT = Path(__file__).resolve().parents[2]
ORIGINAL = ROOT / "RIGOR_ISWC2026_Slide20_Full_Comparison_23_Slides.pptx"
EMU = 914400
PANEL_X, SLIDE_W, SLIDE_H = 8.85, 13.333, 7.5
GOLD_TEXT, MINT = "F9DD8F", "CFE8DC"

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

    # Right panel, behind everything else (first drawable child of the shape tree).
    panel = shape(ids(), "Cover panel", PANEL_X, 0, SLIDE_W - PANEL_X, SLIDE_H, fill=EMERALD)
    sp_tree.insert(2, panel)

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

    # Panel content: motif, its two labels, venue.
    pic = next(p for p in tree.iter(f"{{{P}}}pic"))
    off, ext = pic.find(".//a:off", NS), pic.find(".//a:ext", NS)
    w = 3.70
    h = w * 880 / 984
    x, y = PANEL_X + (SLIDE_W - PANEL_X - w) / 2, 1.55
    off.set("x", str(int(x * EMU))); off.set("y", str(int(y * EMU)))
    ext.set("cx", str(int(w * EMU))); ext.set("cy", str(int(h * EMU)))
    for sid, label, cx in ((27, "relational database", x + 0.21 * w), (28, "ontology", x + 0.76 * w)):
        set_runs(tree, sid, [(label, MINT, False)])
        set_size(tree, sid, 1200)
        move(tree, sid, x=cx - 0.95, y=y + h + 0.12, w=1.9, h=0.26)
        para = sp_by_id(tree, sid).find(".//a:p", NS)
        ppr = para.find("a:pPr", NS)
        if ppr is None:
            ppr = etree.Element(f"{{{A}}}pPr")
            para.insert(0, ppr)
        ppr.set("algn", "ctr")
    add(text_box(ids(), "Venue", PANEL_X + 0.45, 6.30, SLIDE_W - PANEL_X - 0.8, 0.62,
                 [r("ISWC 2026", GOLD_TEXT, 1400, True, "Consolas"),
                  r("Bari, Italy", "FFFFFF", 1300)], line_pts=2000))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    slide = work / "ppt/slides/slide1.xml"
    tree = etree.parse(str(slide))
    edit_cover(tree)
    tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
    # The motif sits on the emerald panel, so recolour it from the original with the emerald palette.
    with zipfile.ZipFile(ORIGINAL) as z:
        (work / "ppt/media/image1.png").write_bytes(z.read("ppt/media/image1.png"))
    recolor_motif(work / "ppt/media/image1.png", TITLE_STYLES["emerald"])
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
