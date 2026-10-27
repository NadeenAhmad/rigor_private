#!/usr/bin/env python3
"""Phase 4: restructure the results as R1 (meaning without losing structure) -> R2 (BURR)
-> R3 (experts), and keep the raw OOPS! table as backup slide B1.

Usage: python3 results.py <in.pptx> <out.pptx>

Every number traces to the RIGOR paper (arXiv:2506.01232v2):
  Table 7 (OOPS! pitfalls), Table 8 (annotations, SubClassOf, provenance),
  Table 9 (schema coverage), Table 13 (BURR), Tables 1 and 12 (experts, LLM judge).
"""
import copy
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}
EMU = 914400
ADD_SLIDE = Path("/root/.claude/skills/synced/954bdbed-f7d8-4e99-9a67-2334e0f54e1f_96ca6ae3-be58-43c5-a447-6fed00e21488"
                 "/pptx/scripts/add_slide.py")

EMERALD, TEAL, INK, BODY, MUTED, GREY, RED = "0B6E4F", "2F6B78", "1F2328", "3A434C", "5B6770", "6B747B", "B23A1E"

# R1 matrix (methods x what their ontologies contain). Coverage = Table 9 column-coverage range
# over the 3 databases x 3 LLMs; counts = runs with any annotations / SubClassOf + provenance (Table 8).
R1_CELLS = {
    # row: (coverage, labels & comments, hierarchy + provenance)
    "DM": ("0.78–0.99", (False, "0 / 3"), (False, "0 / 3")),
    "BL": ("0.31–1.00", (False, "0 / 9"), (False, "0 / 9")),
    "NI": ("0.06–0.78", (True, "9 / 9"), (False, "0 / 9")),
    "RI": ("0.77–1.00", (True, "9 / 9"), (True, "9 / 9")),
}


# ---- helpers -----------------------------------------------------------------------------

def sp_by_id(tree, sid):
    return next(sp for sp in tree.iter(f"{{{P}}}sp", f"{{{P}}}cxnSp")
                if sp.find(".//p:cNvPr", NS).get("id") == str(sid))


def move(tree, sid, x=None, y=None, w=None, h=None):
    sp = sp_by_id(tree, sid)
    off, ext = sp.find(".//a:off", NS), sp.find(".//a:ext", NS)
    for el, attr, val in ((off, "x", x), (off, "y", y), (ext, "cx", w), (ext, "cy", h)):
        if val is not None:
            el.set(attr, str(int(val * EMU)))


def set_runs(tree, sid, runs):
    """Replace a text shape's text with runs [(text, color, bold)] keeping its first run's style.
    A color/bold of None keeps the template's value."""
    sp = sp_by_id(tree, sid)
    body = sp.find("p:txBody", NS)
    paras = body.findall("a:p", NS)
    p0 = paras[0]
    for extra in paras[1:]:
        body.remove(extra)
    template = p0.find("a:r", NS)
    for child in list(p0):
        if etree.QName(child).localname in ("r", "br", "fld"):
            p0.remove(child)
    end = p0.find("a:endParaRPr", NS)
    for text, color, bold in runs:
        r = copy.deepcopy(template)
        rpr = r.find("a:rPr", NS)
        if bold is not None:
            rpr.set("b", "1" if bold else "0")
        if color is not None:
            fill = rpr.find("a:solidFill", NS)
            if fill is None:
                fill = etree.Element(f"{{{A}}}solidFill")
                rpr.insert(0, fill)
            for c in list(fill):
                fill.remove(c)
            etree.SubElement(fill, f"{{{A}}}srgbClr", val=color)
        r.find("a:t", NS).text = text
        if end is not None:
            end.addprevious(r)
        else:
            p0.append(r)


def delete(tree, *ids):
    for sid in ids:
        sp = sp_by_id(tree, sid)
        sp.getparent().remove(sp)


def set_size(tree, sid, size):
    for rpr in sp_by_id(tree, sid).iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr"):
        rpr.set("sz", str(size))


def run(text, color, size, bold=False, font="Calibri"):
    return (f'<a:r><a:rPr lang="en-US" sz="{size}" b="{int(bold)}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:rPr>'
            f'<a:t xml:space="preserve">{escape(text)}</a:t></a:r>')


def shape(sid, name, x, y, w, h, runs_xml="", geom="rect", fill=None, line=None,
          align="l", anchor="t", inset=0.0, descr=""):
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    line_xml = (f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
                else "<a:ln><a:noFill/></a:ln>")
    adj = '<a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst>' if geom == "roundRect" else "<a:avLst/>"
    ins = int(inset * EMU)
    txbox = ' txBox="1"' if runs_xml and not fill else ""
    descr_attr = f' descr="{escape(descr)}"' if descr else ""
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"{descr_attr}/>'
        f'<p:cNvSpPr{txbox}/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/><a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}">{adj}</a:prstGeom>{fill_xml}{line_xml}</p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="{ins}" tIns="0" rIns="{ins}" bIns="0" anchor="{anchor}">'
        f'<a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="{align}"/>{runs_xml}</a:p></p:txBody></p:sp>')


class Ids:
    def __init__(self, tree):
        self.n = max(int(e.get("id")) for e in tree.iter(f"{{{P}}}cNvPr"))

    def __call__(self):
        self.n += 1
        return self.n


def add(tree, el):
    tree.find(".//p:spTree", NS).append(el)


def set_notes(path, paragraphs):
    """Replace the body placeholder text of a notes slide with plain paragraphs."""
    tree = etree.parse(str(path))
    for sp in tree.iter(f"{{{P}}}sp"):
        ph = sp.find(".//p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") == "body":
            body = sp.find("p:txBody", NS)
            template = body.find(".//a:r/a:rPr", NS)
            for p in body.findall("a:p", NS):
                body.remove(p)
            for text in paragraphs:
                p = etree.SubElement(body, f"{{{A}}}p")
                if text:
                    r = etree.SubElement(p, f"{{{A}}}r")
                    if template is not None:
                        r.append(copy.deepcopy(template))
                    etree.SubElement(r, f"{{{A}}}t").text = text
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


# ---- R1: meaning without losing structure (rebuilt from the pitfalls slide) ------------------

def edit_r1(tree):
    """Open table: hairline rules, RIGOR band, coverage drawn as range bars on a 0-1 scale."""
    set_runs(tree, 946, [("RESULT 1  /  MEANING AND STRUCTURE", None, None)])
    set_runs(tree, 947, [("Only RIGOR adds meaning without losing structure.", None, None)])
    set_runs(tree, 948, [("What the generated ontologies contain  ·  3 databases × 3 LLMs  ·  "
                          "Direct Mapping: one ontology per database", None, None)])
    delete(tree, *range(949, 993))  # the old boxed grid, cells and strip box
    ids = Ids(tree)
    x0, x1 = 0.90, 12.43
    cols = {"method": 1.05, "cov": 3.55, "bar": 4.75, "lab": 8.25, "hier": 10.45}
    bar_w = 2.75
    head_y, row_y, row_h = 2.08, 2.80, 0.62

    def label(x, y, w, text, sub):
        add(tree, shape(ids(), "Col head", x, y, w, 0.26, run(text, TEAL, 1150, True, "Consolas")))
        add(tree, shape(ids(), "Col sub", x, y + 0.27, w, 0.22, run(sub, MUTED, 1050)))

    label(cols["method"], head_y, 2.3, "METHOD", "")
    label(cols["cov"], head_y, 4.3, "COLUMNS KEPT", "final schema coverage, range over the runs")
    label(cols["lab"], head_y, 2.1, "LABELS & COMMENTS", "runs that have them")
    label(cols["hier"], head_y, 2.1, "HIERARCHY + PROVENANCE", "runs that have them")
    add(tree, shape(ids(), "Head rule", x0, row_y - 0.08, x1 - x0, 0.012, fill="9AA3A9"))

    names = {"DM": "Direct Mapping", "BL": "LLM baseline", "NI": "Non-Iterative", "RI": "RIGOR"}
    for r, (row, (coverage, labels, hierarchy)) in enumerate(R1_CELLS.items()):
        y = row_y + r * row_h
        rigor = row == "RI"
        if rigor:
            add(tree, shape(ids(), "RIGOR band", x0, y + 0.04, x1 - x0, row_h - 0.08, geom="roundRect",
                            fill="E6F2EC"))
        elif r:
            add(tree, shape(ids(), "Row rule", x0, y, x1 - x0, 0.01, fill="D9DEDB"))
        mid = y + row_h / 2
        add(tree, shape(ids(), "Method", cols["method"], mid - 0.17, 2.4, 0.34,
                        run(names[row], EMERALD if rigor else INK, 1600, True), anchor="ctr"))
        lo, hi = (float(v) for v in coverage.split("–"))
        tone = EMERALD if rigor else RED if row == "NI" else "8E989E"
        add(tree, shape(ids(), "Coverage", cols["cov"], mid - 0.17, 1.15, 0.34,
                        run(coverage, EMERALD if rigor else RED if row == "NI" else INK, 1500, rigor or row == "NI"),
                        anchor="ctr"))
        add(tree, shape(ids(), "Bar track", cols["bar"], mid - 0.06, bar_w, 0.12, geom="roundRect", fill="EEF1EF"))
        add(tree, shape(ids(), "Bar range", cols["bar"] + lo * bar_w, mid - 0.06, max(0.06, (hi - lo) * bar_w), 0.12,
                        geom="roundRect", fill=tone, descr=f"coverage {coverage}"))
        for key, (ok, count) in (("lab", labels), ("hier", hierarchy)):
            add(tree, shape(ids(), "Marker", cols[key], mid - 0.09, 0.18, 0.18, geom="ellipse",
                            fill=EMERALD if ok else "FFFFFF", line=EMERALD if ok else GREY,
                            descr="present" if ok else "absent"))
            add(tree, shape(ids(), "Count", cols[key] + 0.30, mid - 0.17, 1.4, 0.34,
                            run(count, EMERALD if rigor else INK if ok else GREY, 1500, rigor), anchor="ctr"))
    end = row_y + 4 * row_h
    add(tree, shape(ids(), "End rule", x0, end + 0.02, x1 - x0, 0.012, fill="9AA3A9"))
    for k, t in ((0, "0"), (0.5, "0.5"), (1, "1")):  # scale ticks under the bars, like a chart axis
        add(tree, shape(ids(), "Bar tick", cols["bar"] + k * bar_w - 0.25, end + 0.05, 0.5, 0.2,
                        run(t, MUTED, 1000), align="ctr"))

    move(tree, 993, x=x0 + 0.15, y=end + 0.40, w=x1 - x0 - 0.3, h=0.36)
    set_runs(tree, 993, [("MODELING PITFALLS (OOPS!)   ", EMERALD, True),
                         ("Fewer than every other method in 9/9 like-for-like runs  ·  "
                          "zero on Chinook with all three LLMs", INK, False)])
    set_size(tree, 993, 1400)
    move(tree, 994, y=6.45)
    move(tree, 995, y=6.75)
    set_runs(tree, 994, [("Coverage = embedding-based proxy (cosine ≥ 0.55); Direct Mapping is complete by construction. "
                          "Like-for-like = same database and same LLM.", None, None)])
    set_runs(tree, 995, [("Sources: paper Tables 7 (pitfalls), 8 (labels, SubClassOf, provenance), 9 (coverage). "
                          "Raw pitfall counts: backup slide B1.", None, None)])


R1_NOTES = [
    "[9:05–10:00]  SPOKEN SCRIPT (interim; final wording in Phase 6)",
    "First: does RIGOR really add meaning, and does it keep the schema? This table shows what each "
    "method's ontologies contain. Direct mapping keeps the columns but adds no meaning. Single-pass "
    "retrieval adds labels, but drops columns: in one run it keeps only six percent of them. RIGOR is the "
    "only method that produces class hierarchies and provenance in all nine runs, while keeping 77 to 100 "
    "percent of the columns. And its ontologies have the fewest modeling pitfalls in every like-for-like "
    "comparison: zero on Chinook.",
    "",
    "PRESENTER / POINTER: Read across the RIGOR row last. Point at Non-Iterative's 0.06 first.",
    "AUDIENCE TAKEAWAY: Only RIGOR adds hierarchy and provenance and still keeps the schema.",
    "SOURCE AND SCOPE: Paper Tables 7, 8, 9. Coverage is an embedding-based proxy; Direct Mapping is "
    "complete by construction. Disjointness (1/9 runs) is deliberately left for Q&A.",
]


# ---- R2: BURR, explained before the scores -------------------------------------------------

def edit_r2(tree):
    ids = Ids(tree)
    set_runs(tree, 975, [("RESULT 2  /  EXTERNAL BENCHMARK", None, None)])
    set_runs(tree, 976, [("RIGOR keeps the link to the data.", None, None)])
    set_runs(tree, 977, [("BURR benchmark (Laskowski et al., 2025)  ·  ISWC scenario  ·  highest F1 among the "
                          "compared ontology-learning systems", None, None)])

    # Left card: what BURR checks.
    x0, y0, w0, h0 = 0.92, 2.34, 5.30, 3.35
    add(tree, shape(ids(), "BURR card", x0, y0, w0, h0, geom="roundRect", fill="F3F5F4", line="D9DEDB"))
    add(tree, shape(ids(), "BURR card label", x0 + 0.25, y0 + 0.20, w0 - 0.5, 0.30,
                    run("WHAT BURR CHECKS", TEAL, 1300, True, "Consolas")))
    add(tree, shape(ids(), "BURR card lead", x0 + 0.25, y0 + 0.56, w0 - 0.5, 0.34,
                    run("Every ontology element must map back to the database.", INK, 1500)))
    rows = [("CLASSES", "person", "Class: Person"),
            ("RELATIONS", "paper.author_id", "ObjectProperty: author"),
            ("ATTRIBUTES", "person.name", "DataProperty: name")]
    for i, (kind, db, onto) in enumerate(rows):
        y = y0 + 1.04 + i * 0.46
        add(tree, shape(ids(), f"BURR kind {i}", x0 + 0.25, y, 1.10, 0.36,
                        run(kind, TEAL, 1050, True, "Consolas"), anchor="ctr"))
        add(tree, shape(ids(), f"BURR db {i}", x0 + 1.35, y, 1.42, 0.36,
                        run(db, INK, 1050, False, "Consolas"), geom="roundRect", fill="FFFFFF", line="D0D7DE",
                        anchor="ctr", inset=0.08, descr="database element"))
        add(tree, shape(ids(), f"BURR arrow {i}", x0 + 2.79, y, 0.26, 0.36,
                        run("→", MUTED, 1400), align="ctr", anchor="ctr"))
        keyword, name = onto.split(" ")
        add(tree, shape(ids(), f"BURR onto {i}", x0 + 3.07, y, 2.03, 0.36,
                        run(keyword + " ", EMERALD, 1050, True, "Consolas") + run(name, INK, 1050, False, "Consolas"),
                        geom="roundRect", fill="FFFFFF", line="D0D7DE", anchor="ctr", inset=0.08,
                        descr="ontology element"))
    # Matching rule: BURR §4, Eq. 3 (elements match when their mappings are equal).
    add(tree, shape(ids(), "BURR card how", x0 + 0.25, y0 + 2.42, w0 - 0.5, 0.64,
                    run("Elements match when they map to the same database elements, whatever their "
                        "names. F1 per element type.", BODY, 1300)))
    add(tree, shape(ids(), "BURR card note", x0 + 0.25, y0 + 3.10, w0 - 0.5, 0.20,
                    run("Illustrative elements  ·  ISWC gold mapping: D2RQ example by Bizer & Seaborne",
                        MUTED, 1000)))

    # Right: the existing F1 table, compressed into the right half.
    tx, tw = 6.42, 5.98
    move(tree, 978, x=tx, w=tw)
    move(tree, 983, x=tx + 0.22, w=tw - 0.44)
    move(tree, 999, x=tx + 0.12, w=tw - 0.24)
    move(tree, 979, x=tx + 0.25, w=2.0)
    set_size(tree, 979, 1150)
    num_x = {"C": tx + 2.55, "R": tx + 3.70, "A": tx + 4.85}
    for sid, col in ((980, "C"), (981, "R"), (982, "A")):
        move(tree, sid, x=num_x[col] - 0.05, w=1.15)
        set_size(tree, sid, 1150)
    for name_id, tag_id, c_id, r_id, a_id in ((984, 985, 986, 987, 988), (989, 990, 991, 992, 993),
                                              (994, 995, 996, 997, 998), (1000, 1001, 1002, 1003, 1004)):
        name_y = float(sp_by_id(tree, name_id).find(".//a:off", NS).get("y")) / EMU
        move(tree, name_id, x=tx + 0.27, y=name_y - 0.08, w=2.2)
        move(tree, tag_id, x=tx + 0.27, y=name_y + 0.20, w=2.0, h=0.22)
        for sid, col in ((c_id, "C"), (r_id, "R"), (a_id, "A")):
            move(tree, sid, x=num_x[col], w=1.05)
    delete(tree, 1005, 1006)  # strip removed at the author's request
    add(tree, shape(ids(), "BURR caption", 0.93, 5.92, 11.47, 0.27,
                    run("One scenario  ·  9 tables · 46 columns · 11 FKs  ·  RIGOR with Claude  ·  its "
                        "prov:wasDerivedFrom links exported as a D2RQ mapping  ·  Source: paper Table 13",
                        MUTED, 1050)))


R2_NOTES = [
    "[10:00–11:05]  SPOKEN SCRIPT (interim; final wording in Phase 6)",
    "Second: is that link back to the data real? BURR is an external benchmark that asks exactly this. "
    "A system must say which part of the database each ontology element comes from: which table becomes "
    "a class, which foreign key becomes a relation, which column becomes an attribute. BURR compares that "
    "mapping with a gold-standard mapping and reports F1 for each type. RIGOR's provenance annotations give "
    "us this mapping directly; we export them as D2RQ. On the ISWC scenario, RIGOR has the highest F1 for "
    "classes, relations and attributes among the compared ontology-learning systems, including both "
    "rule-based ones. It is one scenario, with Claude.",
    "",
    "PRESENTER / POINTER: Walk the three mapping rows on the left first, then land on the RIGOR row.",
    "AUDIENCE TAKEAWAY: Semantic enrichment did not break the correspondence that OBDA and KG population need.",
    "SOURCE AND SCOPE: Paper Table 13, Strategy 8. BURR = Laskowski et al., PACMMOD 2025. Matching rule: "
    "BURR Section 4, Eq. 3. Left-card elements are illustrative, not the gold standard's actual names.",
    "",
    "DO NOT SAY: 'BURR found LLMs trail rule-based systems'. BURR's results show Gemini 2.5 Pro ahead of "
    "the rule-based systems overall (rule-based win only on NPD).",
    "IF ASKED ABOUT GEMINI 2.5 PRO (BURR Table 3, ISWC .65 / .68 / .56): In BURR it is a general-purpose LLM "
    "prompted with few-shot BURR-style examples and a checklist to write the D2RQ mapping directly. It is a "
    "mapping generator, not an ontology-learning system that builds an enriched ontology, so we compare "
    "with the ontology-learning systems. Against it, RIGOR's class F1 is higher, relations are on par, and "
    "attributes are lower; RIGOR additionally produces labels, hierarchy and provenance. (It is prompted, "
    "not fine-tuned.)",
]


# ---- R3 and B1 ---------------------------------------------------------------------------

def edit_r3(tree):
    set_runs(tree, 913, [("RESULT 3  /  HUMAN AND LLM EVALUATION", None, None)])
    set_runs(tree, 914, [("Experts prefer RIGOR's ontologies.", None, None)])


def edit_b1(tree):
    set_runs(tree, 946, [("BACKUP  /  B1", None, None)])
    set_runs(tree, 947, [("Raw OOPS! pitfall counts (paper Table 7).", None, None)])


def slide_order(work):
    pres = work / "ppt/presentation.xml"
    rels = {r.get("Id"): r.get("Target") for r in etree.parse(str(work / "ppt/_rels/presentation.xml.rels")).getroot()}
    tree = etree.parse(str(pres))
    lst = tree.find("p:sldIdLst", NS)
    return tree, lst, {rels[s.get(f"{{{R}}}id")].split("/")[-1]: s for s in lst}


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    slides = work / "ppt/slides"

    # B1 = copy of the original pitfalls slide, appended after the closing slide.
    out = subprocess.run([sys.executable, str(ADD_SLIDE), str(work), "slide20.xml", "--after", "slide23.xml"],
                         capture_output=True, text=True, check=True).stdout
    b1 = re.search(r"(slide\d+\.xml)", out.split("Created")[-1]).group(1)

    for name, fn in (("slide20.xml", edit_r1), ("slide21.xml", edit_r2), ("slide19.xml", edit_r3), (b1, edit_b1)):
        tree = etree.parse(str(slides / name))
        fn(tree)
        tree.write(str(slides / name), xml_declaration=True, encoding="UTF-8", standalone=True)

    for slide, paragraphs in (("slide20.xml", R1_NOTES), ("slide21.xml", R2_NOTES)):
        rel = etree.parse(str(slides / "_rels" / f"{slide}.rels")).getroot()
        target = next(r.get("Target") for r in rel if r.get("Type").endswith("/notesSlide"))
        set_notes(slides / target, paragraphs)
    rel = etree.parse(str(slides / "_rels/slide19.xml.rels")).getroot()
    n19 = slides / next(r.get("Target") for r in rel if r.get("Type").endswith("/notesSlide"))
    n19.write_text(n19.read_text(encoding="utf-8").replace(
        "First, what do ontology experts think?", "Finally, what do ontology experts think?"), encoding="utf-8")

    # Order: ... 18, R1 (20), R2 (21), R3 (19), 22, 23, B1.
    tree, lst, by_file = slide_order(work)
    s19 = by_file["slide19.xml"]
    lst.remove(s19)
    by_file["slide21.xml"].addnext(s19)
    tree.write(str(work / "ppt/presentation.xml"), xml_declaration=True, encoding="UTF-8", standalone=True)

    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print("wrote", dst, "(B1 =", b1 + ")")


if __name__ == "__main__":
    main(*sys.argv[1:3])
