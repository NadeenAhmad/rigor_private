#!/usr/bin/env python3
"""Results R1-R3 on the author's edited deck, reduced to one message each.

R1  Only RIGOR adds meaning, with far fewer modeling errors:
    left, what each method adds (paper Table 8: annotations, provenance, SubClassOf; same in all runs);
    right, OOPS! modeling errors for the Claude ontologies (paper Table 7 totals).
R2  BURR: unchanged.
R3  An LLM judge and human experts both rank RIGOR first (paper Tables 1 and 12).

Writes a deck with only these three slides; speaker notes come from speech.py.

Usage: python3 results_v3.py <in.pptx> <out.pptx>
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import run_xml  # noqa: E402
from results import (BODY, EMERALD, INK, MUTED, NS, P, TEAL, Ids, add, set_notes, shape)  # noqa: E402
from speech import notes_for, timeline  # noqa: E402

LINE, PANEL, MINT = "D9DEDB", "F3F5F4", "E8F6EF"
HEADER = {"eyebrow": (0.88, 0.56, 11.76, 0.34), "title": (0.88, 1.00, 11.77, 0.73),
          "subtitle": (0.90, 1.80, 11.68, 0.36)}

METHODS = ["Direct Mapping", "LLM baseline", "Non-Iterative", "RIGOR"]
# Table 8: which methods' ontologies carry each element (identical across databases and LLMs).
ADDS = {"Labels & comments": (False, False, True, True),
        "Provenance": (False, False, False, True),
        "Class hierarchy": (False, False, False, True)}
# Table 7: OOPS! pitfall totals, Claude ontologies (Direct Mapping needs no LLM).
ERRORS = {"Liver Cancer": (731, 422, 29, 20), "eICU-CRD": (665, 434, 39, 2), "Chinook": (102, 95, 24, 0)}
BAR_COLORS = ("C5CCD0", "9AA3A9", "6B747B", EMERALD)
# Table 1: first-place counts.
JUDGE = (("Liver Cancer", 16, 16), ("eICU-CRD", 31, 31), ("Chinook", 11, 11))
EXPERT_FIRST = (5, 5, 2, 28)  # of 40 (Liver Cancer + eICU-CRD)


def r(text, color=INK, size=1400, bold=False, font="Calibri", italic=False):
    return run_xml(text, color, bold, size, font=font, italic=italic)


def text(ids, x, y, w, h, runs, align="l", anchor="t", name="Text"):
    return shape(ids(), name, x, y, w, h, runs, align=align, anchor=anchor)


def reset(tree, eyebrow, title, subtitle):
    """Keep the slide's header style: rewrite eyebrow and title, drop everything else."""
    root = tree.find(".//p:spTree", NS)
    shapes = [el for el in root if el.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp", f"{{{P}}}pic", f"{{{P}}}graphicFrame",
                                              f"{{{P}}}grpSp")]
    texts = {}
    for sp in shapes:
        t = "".join(x.text or "" for x in sp.iter(f"{{{NS['a']}}}t")).strip()
        if t.startswith("RESULT") and "eyebrow" not in texts:
            texts["eyebrow"] = sp
        elif "eyebrow" in texts and "title" not in texts and t:
            texts["title"] = sp
    for sp in shapes:
        if sp is not texts["eyebrow"] and sp is not texts["title"]:
            root.remove(sp)
    for key, value in (("eyebrow", eyebrow), ("title", title)):
        sp = texts[key]
        ts = list(sp.iter(f"{{{NS['a']}}}t"))
        ts[0].text = value
        for t in ts[1:]:
            t.getparent().getparent().remove(t.getparent())
        x, y, w, h = HEADER[key]
        off, ext = sp.find(".//a:xfrm/a:off", NS), sp.find(".//a:xfrm/a:ext", NS)
        off.set("x", str(int(x * 914400))); off.set("y", str(int(y * 914400)))
        ext.set("cx", str(int(w * 914400))); ext.set("cy", str(int(h * 914400)))
    ids = Ids(tree)
    add(tree, text(ids, *HEADER["subtitle"], r(subtitle, BODY, 1600), name="Subtitle"))
    return ids


def card(ids, x, y, w, h, label, fill="FFFFFF"):
    return [shape(ids(), "Card", x, y, w, h, geom="roundRect", fill=fill, line=LINE),
            text(ids, x + 0.3, y + 0.22, w - 0.6, 0.3, r(label, TEAL, 1300, True, "Consolas"), name="Card label")]


def dot(ids, x, y, yes):
    d = 0.30
    if yes:
        return shape(ids(), "Yes", x, y, d, d, r("✓", "FFFFFF", 1300, True), geom="ellipse", fill=EMERALD,
                     align="ctr", anchor="ctr")
    el = shape(ids(), "No", x, y, d, d, geom="ellipse", line="9AA3A9")
    return el


def build_r1(tree):
    ids = reset(tree, "RESULT 1  /  MEANING AND MODELING ERRORS",
                "Only RIGOR adds meaning, with far fewer modeling errors.",
                "Ontologies generated with Claude for all three databases. Mistral and DeepSeek show the same pattern.")
    # Left: what each method adds.
    x0, y0, w0, h0 = 0.88, 2.40, 5.45, 4.25
    for el in card(ids, x0, y0, w0, h0, "WHAT EACH METHOD ADDS"):
        add(tree, el)
    cols = list(ADDS)
    col_x = [x0 + 2.35 + k * 1.03 for k in range(3)]
    for cx, label in zip(col_x, cols):
        add(tree, text(ids, cx - 0.12, y0 + 0.70, 1.0, 0.50, r(label, MUTED, 1150), align="ctr", name="Column"))
    row_y = [y0 + 1.40 + k * 0.68 for k in range(4)]
    for k, (method, ry) in enumerate(zip(METHODS, row_y)):
        rigor = method == "RIGOR"
        if rigor:
            add(tree, shape(ids(), "RIGOR band", x0 + 0.15, ry - 0.14, w0 - 0.30, 0.58, geom="roundRect", fill=MINT))
        add(tree, text(ids, x0 + 0.30, ry, 2.0, 0.30, r(method, EMERALD if rigor else INK, 1600, True),
                       anchor="ctr", name="Method"))
        for cx, label in zip(col_x, cols):
            add(tree, dot(ids, cx + 0.38 - 0.15 + 0.0, ry, ADDS[label][k]))
    # Right: OOPS! modeling errors, grouped bars on one linear scale.
    x1, y1, w1, h1 = 6.55, 2.40, 5.90, 4.25
    for el in card(ids, x1, y1, w1, h1, "MODELING ERRORS FOUND BY OOPS!"):
        add(tree, el)
    add(tree, text(ids, x1 + 0.3, y1 + 0.52, 3.0, 0.26, r("fewer is better", MUTED, 1150, italic=True),
                   name="Note"))
    lx = x1 + 0.30
    for k, (method, color) in enumerate(zip(METHODS, BAR_COLORS)):
        cx = lx + k * 1.36
        add(tree, shape(ids(), "Legend key", cx, y1 + 0.93, 0.16, 0.16, fill=color))
        add(tree, text(ids, cx + 0.22, y1 + 0.86, 1.15, 0.28,
                       r(method, EMERALD if method == "RIGOR" else BODY, 1100, method == "RIGOR"), name="Legend"))
    base, top = y1 + h1 - 0.62, y1 + 1.55
    scale = (base - top) / max(max(v) for v in ERRORS.values())
    group_w, bar_w, gap = 1.75, 0.30, 0.08
    gx0 = x1 + 0.33
    add(tree, shape(ids(), "Baseline", gx0 - 0.05, base, 3 * group_w + 0.05, 0.012, fill="9AA3A9"))
    for g, (db, values) in enumerate(ERRORS.items()):
        gx = gx0 + g * group_w + (group_w - (4 * bar_w + 3 * gap)) / 2 - 0.05
        for k, (v, color) in enumerate(zip(values, BAR_COLORS)):
            bx = gx + k * (bar_w + gap)
            h = v * scale
            if h > 0:
                add(tree, shape(ids(), "Bar", bx, base - h, bar_w, h, fill=color))
            rigor = k == 3
            add(tree, text(ids, bx - 0.15, base - h - 0.27, bar_w + 0.30, 0.24,
                           r(str(v), EMERALD if rigor else BODY, 1300 if rigor else 1050, rigor), align="ctr",
                           name="Value"))
        add(tree, text(ids, gx0 + g * group_w - 0.05, base + 0.10, group_w, 0.28, r(db, INK, 1200, True),
                       align="ctr", name="Database"))
    add(tree, text(ids, 0.90, 6.88, 11.6, 0.26,
                   r("Sources: paper Table 8 (labels, provenance, hierarchy in every ontology of each method) and "
                     "Table 7 (OOPS! totals).", MUTED, 1050), name="Source"))


def build_r3(tree):
    ids = reset(tree, "RESULT 3  /  LLM JUDGE AND HUMAN EXPERTS",
                "An LLM judge and human experts both rank RIGOR first.",
                "Both ranked four anonymized ontologies per table, from first to fourth.")
    y, h, w = 2.40, 4.25, 5.65
    # Left: LLM judge.
    x = 0.88
    for el in card(ids, x, y, w, h, "LLM-AS-A-JUDGE  ·  GPT-5.4", fill=PANEL):
        add(tree, el)
    add(tree, text(ids, x + 0.3, y + 0.75, w - 0.6, 1.05, r("58 / 58", INK, 6000, True, "Cambria"), name="Big"))
    add(tree, text(ids, x + 0.3, y + 1.85, w - 0.6, 0.40, r("tables with RIGOR ranked first", EMERALD,
                                                           1900, True), name="Claim"))
    for k, (db, a, b) in enumerate(JUDGE):
        ry = y + 2.55 + k * 0.40
        add(tree, text(ids, x + 0.3, ry, 2.4, 0.30, r(db, BODY, 1500), name="Database"))
        add(tree, text(ids, x + 2.7, ry, 1.2, 0.30, r(f"{a} / {b}", EMERALD, 1500, True), name="Count"))
    # Right: human experts.
    x = 6.80
    for el in card(ids, x, y, w, h, "10 ONTOLOGY EXPERTS", fill=PANEL):
        add(tree, el)
    add(tree, text(ids, x + 0.3, y + 0.75, w - 0.6, 1.05, r("28 / 40", INK, 6000, True, "Cambria"), name="Big"))
    add(tree, text(ids, x + 0.3, y + 1.85, w - 0.6, 0.40, r("rankings with RIGOR first (70%)", EMERALD, 1900, True),
                   name="Claim"))
    for k, (method, v) in enumerate(zip(METHODS, EXPERT_FIRST)):
        ry = y + 2.55 + k * 0.38
        rigor = method == "RIGOR"
        add(tree, text(ids, x + 0.3, ry, 1.75, 0.30, r(method, EMERALD if rigor else BODY, 1350, rigor),
                       anchor="ctr", name="Method"))
        add(tree, shape(ids(), "Track", x + 2.10, ry + 0.08, 2.55, 0.15, fill="E1E6E3"))
        add(tree, shape(ids(), "Bar", x + 2.10, ry + 0.08, 2.55 * v / 40, 0.15, fill=EMERALD if rigor else "9AA3A9"))
        add(tree, text(ids, x + 4.75, ry, 0.65, 0.30, r(str(v), EMERALD if rigor else BODY, 1350, rigor),
                       align="r", anchor="ctr", name="Count"))
    add(tree, text(ids, 0.90, 6.88, 11.6, 0.26,
                   r("LLM judge: 58 tables, 290 competency questions. Experts: 4 medical tables, significant agreement "
                     "(Kendall's W 0.39–0.72). Sources: paper Tables 1 and 12.", MUTED, 1050), name="Source"))


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
    keep = {}
    for sld in list(lst):
        name = rels[sld.get(f"{{{R}}}id")]
        xml = (work / "ppt/slides" / name).read_text(encoding="utf-8")
        tag = next((t for t in ("RESULT 1", "RESULT 2", "RESULT 3") if t in xml), None)
        if tag:
            keep[tag] = name
        else:
            lst.remove(sld)
    for sec in pres.iter("{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst"):
        sec.getparent().remove(sec)
    pres.write(str(pres_path), xml_declaration=True, encoding="UTF-8", standalone=True)
    rows = {row[1]["title"]: row for row in timeline()}
    for tag, build, title in (("RESULT 1", build_r1, "Result 1: meaning and modeling errors"),
                              ("RESULT 3", build_r3, "Result 3: LLM judge and experts")):
        path = work / "ppt/slides" / keep[tag]
        tree = etree.parse(str(path))
        build(tree)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
        notes = re.search(r'notesSlides/(notesSlide\d+\.xml)',
                          (work / "ppt/slides/_rels" / f"{keep[tag]}.rels").read_text())
        if notes:
            n, s, start, end, _ = rows[title]
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
