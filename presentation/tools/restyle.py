#!/usr/bin/env python3
"""Restyle the RIGOR ISWC 2026 deck in place: new palette, fonts, constant backgrounds,
step-tracker states and a few text edits. Layout (positions, sizes, shapes) is untouched.

Usage: python3 restyle.py <original.pptx> <output.pptx>

Palette ("lab notebook / editorial light"):
  white FFFFFF background · panels F3F5F4 · hairlines D9DEDB
  ink 1F2328 · body 3A434C · secondary 5B6770
  emerald 0B6E4F (primary accent) · teal 2F6B78 / 3A7D8C (secondary) · vermilion B23A1E (errors)
  Title and closing slides: emerald background, white text, gold F6CF6A highlight.
"""
import copy
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P}
EMU = 914400

# Slide groups by their original background.
EMERALD_SLIDES = {1, 23}
DARK_SLIDES = {2, 3, 5, 9, 12, 13, 14, 15, 16, 17, 19, 20, 21}
LIGHT_SLIDES = {4, 6, 7, 8, 10, 11, 18, 22}
TRACKER_SLIDES = {5: 0, 6: 1, 7: 2, 8: 3, 9: 4, 10: 5, 11: 6}  # slide -> current step (0 = none)

EMERALD, EMERALD_TINT, TEAL, TEAL_LINE = "0B6E4F", "E6F2EC", "2F6B78", "3A7D8C"
INK, BODY, MUTED, HAIRLINE, PANEL = "1F2328", "3A434C", "5B6770", "D9DEDB", "F3F5F4"

ORANGE = {"F9AB5E", "F4A261", "FF9F52", "FFA159", "FFA057"}
BLUE = {"84ADF7", "7799F0", "7AA1F4", "7BA4F4", "4D71DE"}
MINT = {"8DE1B6"}

# Navy panel levels (dark slides): deeper nesting was lighter navy, becomes darker grey.
NAVY_LEVELS = [
    ({"0E1726", "0A1427", "0D182B", "132038", "132641", "18273E", "182740", "1A2B45",
      "1D2F4C", "18314C", "193548"}, "FFFFFF"),
    ({"1F3957", "21314A", "233754"}, PANEL),
    ({"253B57", "263959", "263A56", "263E62"}, "E9EEEB"),
    ({"2E5369", "334D66", "3D495F"}, "DDE5E1"),
    ({"455B7A", "546B8C"}, "D3DAD6"),
]
NAVY_LINES = {"263A56", "304C70", "455B7A", "182740", "1D2F4C", "435E84", "42597E", "384F76",
              "485B81", "4C6491", "5A7092", "576A8A", "465C7F"}
GREY_LINES = {"8899B2", "7F8CA1", "C7D4E8"}

TEXT_DARK_TO_LIGHT = {
    **{c: INK for c in ("F7F9FD", "EDF1F6", "EAF1FA", "EBF2FA", "EBEFF5")},
    **{c: BODY for c in ("C7D4E8", "C9D2DE", "B5C3D8")},
    **{c: MUTED for c in ("AEBCCF", "ACB9CD", "9AA6B8", "93A0B3", "7F8CA1")},
}


def navy_level(c):
    for colors, out in NAVY_LEVELS:
        if c in colors:
            return out
    return None


def map_dark(c, role):
    """Dark slide -> white slide."""
    if role == "bg":
        return "FFFFFF"
    if c in ORANGE:
        return EMERALD
    if c in MINT:
        return "1E8A5E"
    if c in BLUE:
        return TEAL if role == "text" else TEAL_LINE
    if role == "text":
        return TEXT_DARK_TO_LIGHT.get(c)
    if role == "line":
        if c in NAVY_LINES:
            return HAIRLINE
        if c in GREY_LINES:
            return "8E989E"
    return navy_level(c)


LIGHT_MAP = {
    "F2F4F7": PANEL, "FFFFFF": "FFFFFF", "E4EBFF": EMERALD_TINT, "E3EAFC": EMERALD_TINT,
    "4D71DE": EMERALD, "5279E7": EMERALD, "7799F0": "3F8F6F",
    "D68B4B": TEAL, "F4A261": TEAL_LINE,
    "FFECE9": "FBE9E4", "AA3737": "B23A1E",
    "111A28": INK, "111B2B": INK, "5D6879": MUTED, "5C697D": MUTED,
    "DFE5ED": HAIRLINE, "DAE2EE": HAIRLINE, "C4D3F1": "B7D7C7",
}


def map_light(c, role):
    if role == "bg":
        return "FFFFFF"
    if c == "EDF1F6" and role == "text":
        return "FFFFFF"
    return LIGHT_MAP.get(c)


EMERALD_DEEP, EMERALD_PANEL, EMERALD_EDGE, GOLD = "095C42", "0D7A58", "4FAE88", "F6CF6A"


def map_emerald(c, role):
    """Dark slide -> emerald slide (title and closing)."""
    if role == "bg" or c == "0E1726":
        return EMERALD
    if c in ORANGE:
        return GOLD
    if role == "text":
        if c in ("F7F9FD", "EDF1F6", "EAF1FA", "EBF2FA", "EBEFF5", "FFFFFF"):
            return "FFFFFF"
        if c in BLUE:
            return "BFE6D3"
        return "D3EDE0"
    if role == "line":
        return "8FD3B4" if c in BLUE else EMERALD_EDGE
    if c in {"1A2B45", "18273E"}:
        return EMERALD_DEEP
    if c in {"233754", "21314A", "263A56"}:
        return EMERALD_PANEL
    if c in BLUE:
        return "8FD3B4"
    return None


def role_of(el):
    tags = [etree.QName(a).localname for a in el.iterancestors()]
    if "bg" in tags:
        return "bg"
    if any(t in ("rPr", "defRPr", "endParaRPr") for t in tags):
        return "text"
    if "ln" in tags:
        return "line"
    return "fill"


FONT_MAP = {"IBM Plex Sans": "Calibri", "Aptos": "Calibri", "Arial": "Calibri",
            "IBM Plex Mono": "Consolas"}
TITLE_FONT, TITLE_MIN_SZ = "Cambria", 2600


def restyle_fonts(tree):
    for el in tree.iter(f"{{{A}}}latin", f"{{{A}}}ea", f"{{{A}}}cs", f"{{{A}}}sym"):
        tf = el.get("typeface")
        if tf in FONT_MAP:
            el.set("typeface", FONT_MAP[tf])
    # Large text (titles, hero numbers) gets the serif heading font.
    for rpr in tree.iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr"):
        if int(rpr.get("sz", "0")) >= TITLE_MIN_SZ:
            latin = rpr.find("a:latin", NS)
            if latin is None:
                latin = etree.SubElement(rpr, f"{{{A}}}latin")
            if latin.get("typeface") != "Consolas":
                latin.set("typeface", TITLE_FONT)


def set_fill(sppr_or_rpr, color, tag="solidFill"):
    sf = sppr_or_rpr.find(f"a:{tag}", NS)
    if sf is None:
        return
    for child in list(sf):
        sf.remove(child)
    etree.SubElement(sf, f"{{{A}}}srgbClr", val=color)


def shape_text(sp):
    return "".join(t.text or "" for t in sp.iter(f"{{{A}}}t")).strip()


def shape_box(sp):
    off, ext = sp.find(".//a:off", NS), sp.find(".//a:ext", NS)
    return tuple(int(v) / EMU for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")))


def restyle_tracker(tree, current):
    """Step tracker: done = emerald tint, current = solid emerald, upcoming = hollow grey."""
    tree_sps = list(tree.iter(f"{{{P}}}sp"))
    for sp in tree_sps:
        m = re.fullmatch(r"(\d)\s+(ORDER|SEED|RETRIEVE|GENERATE|VALIDATE|MERGE)", shape_text(sp))
        if not m:
            continue
        step = int(m.group(1))
        x, y, _, _ = shape_box(sp)
        box = next(b for b in tree_sps if b is not sp and b.find(".//a:prstGeom[@prst='roundRect']", NS) is not None
                   and abs(shape_box(b)[0] - (x - 0.06)) < 0.05 and abs(shape_box(b)[1] - (y - 0.04)) < 0.05)
        if current and step < current:
            fill, line, text, bold = EMERALD_TINT, EMERALD_TINT, EMERALD, "1"
        elif step == current:
            fill, line, text, bold = EMERALD, EMERALD, "FFFFFF", "1"
        else:
            fill, line, text, bold = "FFFFFF", HAIRLINE, "8A949B", "1"
        sppr = box.find("p:spPr", NS)
        set_fill(sppr, fill)
        set_fill(sppr.find("a:ln", NS), line)
        for rpr in sp.iter(f"{{{A}}}rPr"):
            rpr.set("b", bold)
            set_fill(rpr, text)


def clone_text_shape(tree, source_sp, new_id, x, y, w, h, runs, align=None):
    """Copy a text shape's style, then reposition it and replace its text with runs
    [(text, color, bold, size)]."""
    sp = copy.deepcopy(source_sp)
    cnv = sp.find(".//p:cNvPr", NS)
    cnv.set("id", str(new_id))
    cnv.set("name", f"Added {new_id}")
    off, ext = sp.find(".//a:off", NS), sp.find(".//a:ext", NS)
    off.set("x", str(int(x * EMU))); off.set("y", str(int(y * EMU)))
    ext.set("cx", str(int(w * EMU))); ext.set("cy", str(int(h * EMU)))
    txbody = sp.find("p:txBody", NS)
    paras = txbody.findall("a:p", NS)
    p0 = paras[0]
    for extra in paras[1:]:
        txbody.remove(extra)
    template_r = p0.find("a:r", NS)
    for child in list(p0):
        if etree.QName(child).localname in ("r", "br", "fld"):
            p0.remove(child)
    ppr = p0.find("a:pPr", NS)
    if align:
        if ppr is None:
            ppr = etree.Element(f"{{{A}}}pPr")
            p0.insert(0, ppr)
        ppr.set("algn", align)
    end = p0.find("a:endParaRPr", NS)
    for text, color, bold, size in runs:
        r = copy.deepcopy(template_r)
        rpr = r.find("a:rPr", NS)
        rpr.set("b", "1" if bold else "0")
        if size:
            rpr.set("sz", str(size))
        set_fill(rpr, color)
        r.find("a:t", NS).text = text
        if end is not None:
            end.addprevious(r)
        else:
            p0.append(r)
    source_sp.getparent().append(sp)
    return sp


def max_id(tree):
    return max(int(e.get("id")) for e in tree.iter(f"{{{P}}}cNvPr"))


def find_sp(tree, text_start):
    return next(sp for sp in tree.iter(f"{{{P}}}sp") if shape_text(sp).startswith(text_start))


def text_edits(slide_no, tree):
    if slide_no == 5:  # Result preview under the step tracker
        src = find_sp(tree, "We need to preserve structure")
        clone_text_shape(tree, src, max_id(tree) + 1, 0.94, 6.86, 11.44, 0.32, [
            ("RESULT PREVIEW   ", EMERALD, True, 1300),
            ("Ranked first in 70% of expert rankings  ·  best F1 on the BURR benchmark", MUTED, False, 1300),
        ])
    if slide_no == 10:  # Make the hypothetical status of the examples visible
        src = find_sp(tree, "STEP 5 / VALIDATE")
        clone_text_shape(tree, src, max_id(tree) + 1, 7.67, 0.49, 5.0, 0.34,
                         [("ILLUSTRATIVE ERRORS, NOT OBSERVED OUTPUTS", MUTED, True, None)], align="r")
    if slide_no == 18:
        for t in tree.iter(f"{{{A}}}t"):
            if t.text and "independent LLM judge" in t.text:
                t.text = t.text.replace("independent LLM judge", "independent evaluation judge (GPT-5.4)")
    if slide_no in (12, 13, 14, 15, 16, 17, 20):  # Eyebrow labels: monospace like every other slide
        eyebrow = find_sp(tree, "RIGOR  /  REASSEMBLING" if slide_no != 20 else "RESULT 2")
        for rpr in eyebrow.iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr"):
            latin = rpr.find("a:latin", NS)
            if latin is None:
                latin = etree.SubElement(rpr, f"{{{A}}}latin")
            latin.set("typeface", "Consolas")
    if slide_no == 7:  # Labels that overflow their boxes even at Consolas width
        for start, size in (("INPUT  /  admission", 1500), ("OUTPUT  /  DMr", 1500), ("FK → PatientData", 1250)):
            for rpr in find_sp(tree, start).iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr"):
                rpr.set("sz", str(size))


def restyle_slide(path, slide_no, unmapped):
    tree = etree.parse(str(path))
    mapper = map_emerald if slide_no in EMERALD_SLIDES else map_dark if slide_no in DARK_SLIDES else map_light
    for el in tree.iter(f"{{{A}}}srgbClr"):
        old, role = el.get("val").upper(), role_of(el)
        new = mapper(old, role)
        if new is None:
            unmapped.add((slide_no, old, role))
        else:
            el.set("val", new)
    restyle_fonts(tree)
    if slide_no in TRACKER_SLIDES:
        restyle_tracker(tree, TRACKER_SLIDES[slide_no])
    text_edits(slide_no, tree)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def recolor_motif(png):
    """Title-slide motif (navy table + graph, orange highlight) -> emerald + gold."""
    im = Image.open(png).convert("RGBA")
    px = im.load()
    for yy in range(im.height):
        for xx in range(im.width):
            r, g, b, a = px[xx, yy]
            if a == 0:
                continue
            if r > 180 and r > g > b:  # orange highlight
                px[xx, yy] = (0xF6, 0xCF, 0x6A, a)
            elif b > 200:  # light-blue edges and node rims
                t = min(1.0, max(0.0, (b - 200) / 55))
                px[xx, yy] = (int(150 + 50 * t), int(215 + 25 * t), int(185 + 30 * t), a)
            else:  # navy table, nodes, placeholders: keep relative lightness on a deep-emerald ramp
                t = min(1.0, max(0.0, ((r + g + b) / 3 - 14) / 100))
                lo, hi = (0x05, 0x45, 0x31), (0x5F, 0xB5, 0x92)
                px[xx, yy] = tuple(int(lo[i] + (hi[i] - lo[i]) * t) for i in range(3)) + (a,)
    im.save(png)


def drop_embedded_fonts(root):
    pres = root / "ppt/presentation.xml"
    tree = etree.parse(str(pres))
    for lst in tree.findall("p:embeddedFontLst", NS):
        lst.getparent().remove(lst)
    tree.write(str(pres), xml_declaration=True, encoding="UTF-8", standalone=True)
    rels = root / "ppt/_rels/presentation.xml.rels"
    rtree = etree.parse(str(rels))
    for rel in list(rtree.getroot()):
        if rel.get("Type", "").endswith("/font"):
            rtree.getroot().remove(rel)
    rtree.write(str(rels), xml_declaration=True, encoding="UTF-8", standalone=True)
    shutil.rmtree(root / "ppt/fonts", ignore_errors=True)
    ct = root / "[Content_Types].xml"
    ct.write_text(re.sub(r'<Default[^>]*Extension="fntdata"[^>]*/>', "", ct.read_text()))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    order = re.findall(r'r:id="(rId\d+)"', re.search(
        r"<p:sldIdLst>(.*?)</p:sldIdLst>", (work / "ppt/presentation.xml").read_text()).group(1))
    rels = {rel.get("Id"): rel.get("Target") for rel in etree.parse(
        str(work / "ppt/_rels/presentation.xml.rels")).getroot()}
    unmapped = set()
    for n, rid in enumerate(order, 1):
        restyle_slide(work / "ppt" / rels[rid], n, unmapped)
    for extra in [*(work / "ppt/slideMasters").glob("*.xml"), *(work / "ppt/slideLayouts").glob("*.xml"),
                  *(work / "ppt/notesSlides").glob("*.xml")]:
        tree = etree.parse(str(extra))
        restyle_fonts(tree)
        for el in tree.iter(f"{{{A}}}srgbClr"):
            el.set("val", {"111A28": INK, "F2F4F7": "FFFFFF"}.get(el.get("val"), el.get("val")))
        tree.write(str(extra), xml_declaration=True, encoding="UTF-8", standalone=True)
    for theme in (work / "ppt/theme").glob("*.xml"):
        t = theme.read_text()
        t = re.sub(r'(<a:majorFont><a:latin typeface=")[^"]*"', r'\1Cambria"', t)
        t = re.sub(r'(<a:minorFont><a:latin typeface=")[^"]*"', r'\1Calibri"', t)
        theme.write_text(t)
    recolor_motif(work / "ppt/media/image1.png")
    drop_embedded_fonts(work)
    if unmapped:
        for u in sorted(unmapped):
            print("UNMAPPED", u)
        sys.exit(1)
    Path(dst).unlink(missing_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        files = sorted(f for f in work.rglob("*") if f.is_file())
        files.sort(key=lambda f: f.name != "[Content_Types].xml")
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print("wrote", dst)


if __name__ == "__main__":
    main(*sys.argv[1:3])
