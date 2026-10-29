#!/usr/bin/env python3
"""Readability and numbering pass on the author's final deck (RIGOR_ISWC2026_4.pptx).

  * Small text is enlarged (nothing on a talk slide below 10.5 pt except citations, which stay 10 pt);
    per-line code listings are re-spaced so the larger lines still fit their panels.
  * Step 2's code panel loses its tab-aligned glosses (tabs wrap in PowerPoint and vanish in Google
    Slides); each gloss becomes its own right-aligned text box on the line it explains.
  * Slide numbers count talk slides, not physical slides: every reveal of a slide shows one number.
    Title and closing slides are unnumbered. On the architecture reveal the "k / 6" counter moves under
    the progress dashes so the corner is free for the number.
  * Speaker notes come from speech_final.py.

Usage: python3 polish_final.py <author.pptx> <out.pptx>
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
import speech_final  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
EMU = 914400
MUTED = "6B747B"


def emu(v):
    return int(round(v * EMU))


def inch(v):
    return int(v) / EMU


def shapes(tree):
    return list(tree.getroot().find("p:cSld/p:spTree", NS))


def sid(el):
    c = el.find(".//p:cNvPr", NS)
    return c.get("id") if c is not None else None


def xfrm(el):
    return el.find(".//a:xfrm", NS)


def geom(el):
    x = xfrm(el)
    if x is None:
        return None
    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
    return inch(off.get("x")), inch(off.get("y")), inch(ext.get("cx")), inch(ext.get("cy"))


def set_geom(el, x=None, y=None, w=None, h=None):
    off, ext = xfrm(el).find("a:off", NS), xfrm(el).find("a:ext", NS)
    if x is not None:
        off.set("x", str(emu(x)))
    if y is not None:
        off.set("y", str(emu(y)))
    if w is not None:
        ext.set("cx", str(emu(w)))
    if h is not None:
        ext.set("cy", str(emu(h)))


def text(el):
    return "".join(t.text or "" for t in el.iter(f"{{{A}}}t"))


def resize(el, mapping):
    """Map run sizes (rPr and endParaRPr) through `mapping` (old sz -> new sz)."""
    for tag in ("rPr", "endParaRPr", "defRPr"):
        for rp in el.iter(f"{{{A}}}{tag}"):
            sz = rp.get("sz")
            if sz and int(sz) in mapping:
                rp.set("sz", str(mapping[int(sz)]))


def by_id(tree, i):
    return next(el for el in shapes(tree) if sid(el) == str(i))


# ---- per-line code listings -------------------------------------------------------------------

def regrid(tree, region, y0_old, pitch_old, y0_new, pitch_new, sizes, blanks, last_row, grow=(), blank_frac=0.5):
    """Re-space a per-line listing. `region` = (x0, x1, y_top, y_bottom) selects the line shapes
    (code lines, glosses, tags, highlight bars). Rows keep their order; `blanks` are the blank lines of
    the complete listing (taken from the last reveal, so every reveal uses the same rows) and get
    `blank_frac` of a pitch. `grow` = ids of panels whose bottom follows the last row."""
    x0, x1, yt, yb = region
    picked = []
    for el in shapes(tree):
        g = geom(el)
        if g and x0 - 0.01 <= g[0] <= x1 and yt - 0.01 <= g[1] <= yb and g[3] < pitch_old * 1.5:
            picked.append((el, g))
    k = pitch_new / pitch_old

    def new_row(r):
        return r - sum(1 - blank_frac for b in blanks if b < r)

    bottom = y0_new + new_row(last_row) * pitch_new + pitch_new  # same on every reveal
    for el, (x, y, w, h) in picked:
        r = round((y - y0_old) / pitch_old)
        ny = y0_new + new_row(r) * pitch_new + (y - (y0_old + r * pitch_old)) * k
        set_geom(el, y=ny, h=h * k)
        resize(el, sizes)
    for i, pad in grow:
        el = by_id(tree, i)
        x, y, w, h = geom(el)
        set_geom(el, h=bottom + pad - y)
    return bottom


def slides_3_to_8(tree):
    """'What we want': code 10.5 -> 11.5 pt, glosses 10 -> 11 pt; panel and card follow."""
    ids = {sid(el) for el in shapes(tree)}
    if not any(text(el).startswith("Class:") for el in shapes(tree)):
        return
    first = min((el for el in shapes(tree) if text(el).strip() == "Class: Admission"), key=lambda e: geom(e)[1])
    code_panel = max((el for el in shapes(tree) if geom(el) and not text(el).strip()
                      and abs(geom(el)[1] - 3.50) < 0.03 and geom(el)[2] > 6), key=lambda e: geom(e)[2])
    card = max((el for el in shapes(tree) if geom(el) and not text(el).strip()
                and abs(geom(el)[1] - 2.25) < 0.03 and geom(el)[2] > 7), key=lambda e: geom(e)[2])
    bottom = regrid(tree, (5.40, 12.2, 3.55, 6.95), geom(first)[1], 0.1775, 3.54, 0.19,
                    {1050: 1150, 1000: 1100}, blanks=[6], last_row=18, grow=[(sid(code_panel), 0.06)])
    x, y, w, h = geom(card)
    set_geom(card, h=bottom + 0.14 - y)
    for el in shapes(tree):  # the 'What this enables' box on the last reveal shares the card bottom
        g = geom(el)
        if g and abs(g[0] - 0.88) < 0.02 and abs(g[1] - 5.56) < 0.02:
            set_geom(el, h=bottom + 0.14 - g[1])
    assert ids == {sid(el) for el in shapes(tree)}


def slides_17_18(tree):
    """Step 4 BEFORE/AFTER: code 11 -> 11.5 pt, source tags 8.5 -> 10 pt."""
    bottom = regrid(tree, (0.65, 12.3, 2.50, 5.80), 2.54, 0.1775, 2.53, 0.19, {1100: 1150, 850: 1000}, blanks=[6],
                    last_row=18)
    for el in shapes(tree):
        g = geom(el)
        if not g or text(el).strip():
            continue
        if abs(g[1] - 2.46) < 0.02 and g[2] > 3:  # code panels
            set_geom(el, h=bottom + 0.07 - g[1])
        elif abs(g[1] - 2.02) < 0.02 and g[2] > 3.4:  # cards
            set_geom(el, h=bottom + 0.16 - g[1])


# ---- Step 2: glosses out of the code text ----------------------------------------------------

def slide_12(tree):
    code = next(el for el in shapes(tree) if text(el).startswith("Class: Admission"))
    x, y, w, h = geom(code)
    body = code.find("p:txBody", NS)
    bp = body.find("a:bodyPr", NS)
    top = y + int(bp.get("tIns", "45720")) / EMU
    right = x + w - int(bp.get("rIns", "91440")) / EMU
    paras = body.findall("a:p", NS)
    glosses = []
    for i, p in enumerate(paras):
        runs = p.findall("a:r", NS)
        cut = next((j for j, r in enumerate(runs) if "\t" in (r.findtext("a:t", namespaces=NS) or "")
                    or "←" in (r.findtext("a:t", namespaces=NS) or "")), None)
        if cut is not None:
            note = "".join(r.findtext("a:t", namespaces=NS) or "" for r in runs[cut:])
            glosses.append((i, note.replace("\t", "").strip(), copy.deepcopy(runs[-1].find("a:rPr", NS))))
            for r in runs[cut:]:
                p.remove(r)
            last = p.findall("a:r", NS)[-1].find("a:t", NS)
            last.text = last.text.rstrip()
        for r in p.findall("a:r", NS):  # one code size throughout
            rp = r.find("a:rPr", NS)
            if rp.find("a:latin", NS) is not None and rp.find("a:latin", NS).get("typeface") == "Consolas":
                rp.set("sz", "1400")
    pitch = 14 * 1.2 * 1.14285 / 72  # sz x single spacing x lnSpc, in inches
    nid = max(int(sid(el)) for el in shapes(tree) if sid(el)) + 1
    tree_root = tree.getroot().find("p:cSld/p:spTree", NS)
    for i, note, rpr in glosses:
        rpr.set("sz", "1200")
        box = textbox(nid, "Gloss", right - 2.3, top + i * pitch, 2.3, pitch, note, rpr, "r")
        tree_root.append(box)
        nid += 1


def textbox(i, name, x, y, w, h, s, rpr, algn="l"):
    sp = etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{i}" name="{name} {i}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/>'
        f'<a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr">'
        f'<a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="{algn}"/><a:r><a:t/></a:r></a:p></p:txBody></p:sp>')
    r = sp.find(".//a:r", NS)
    r.insert(0, rpr)
    r.find("a:t", NS).text = s
    return sp


# ---- small text elsewhere ----------------------------------------------------------------------

TABLE_SIZES = {850: 1050, 950: 1050, 1000: 1100, 1100: 1200, 1150: 1200}


def results_tables(tree):
    for el in shapes(tree):
        g = geom(el)
        if g and g[1] > 6.8:  # citation lines stay 10 pt
            continue
        resize(el, TABLE_SIZES)


def burr(tree):
    for el in shapes(tree):
        g = geom(el)
        if g and g[1] > 2.2 and g[1] < 7.0:
            resize(el, {1130: 1200})  # mapping boxes are full at 10.5 pt


def validate_cards(tree):
    """Step 5: the explanations run to four lines; both cards grow by 0.2 in, the code line widens."""
    for el in shapes(tree):
        g = geom(el)
        if text(el).startswith("ObjectProperty: BM_DEP"):
            set_geom(el, w=4.80)
        if g and abs(g[1] - 1.93) < 0.02 and abs(g[3] - 2.70) < 0.02:
            set_geom(el, h=2.90)
        elif g and abs(g[1] - 3.775) < 0.02 and abs(g[3] - 0.71) < 0.02:
            set_geom(el, h=0.90)


def judges_footer(tree):
    """Result 5: the corrected citation wraps to two lines; keep it clear of the slide number."""
    for el in shapes(tree):
        if text(el).startswith("Quality dimensions"):
            set_geom(el, y=6.95, w=11.40)


def ladder_only(tree):
    for el in shapes(tree):
        if re.match(r"^\d  [A-Z ,]+$", text(el)):
            resize(el, {1000: 1100})


def future_data(tree):
    resize(tree.getroot(), {950: 1100, 1150: 1200})


# ---- corrections the author approved ---------------------------------------------------------

# (physical slides, paragraph must contain, {exact run text: new run text})
FIXES = [
    (range(31, 36), "Direct Mapping [Goldberger", {"[Goldberger et al. 2000]  ": "[Sequeda et al., WWW 2012]  "}),
    (range(29, 36), "PhysioNet [Goldberger", {"[Goldberger et al. 2000]  ": "[Pollard et al., Sci. Data 2018]  "}),
    (range(36, 37), "valid OWL2DL", {"All 30 ontologies are": "All 30 ontologies", " valid OW": " parse without errors ",
                                      "L": "", "2DL ": "", "(RDFLib) and ": "(RDFLib) and are "}),
    (range(36, 37), "OPPS!", {"OPPS!": "OOPS!"}),
    (range(36, 37), "220 without RIGOR", {"Mistral on eICU, 220 without RIGOR": "Mistral on eICU, 220 with Non-Iterative RAG",
                                           "2 with it.": "2 with RIGOR."}),
    # Step 5, structural card: the LLM turns BM_DEP into a link, but the seed has it as a number (punning).
    (range(19, 22), "reuses a core property", {"Later table ": "The LLM links ", "lab_result": "BM_DEP",
                                               "reuses": "to a disease", " a core property:": ":"}),
    (range(19, 22), "hasPatientData  Domain: LabResult", {" hasPatientData  ": " BM_DEP  ", "Domain:": "Range:",
                                                          " LabResult": " BoneMarrowDepression"}),
    (range(19, 22), "Merged with PatientData", {
        "Merged with ": "But the seed already has ",
        "PatientData Domain: Admission": "DataProperty: BM_DEP",
        ", OWL infers that every lab result is an admission (pitfall P19). Graph checks allow ":
            ", a number. One property can’t be both a number and a link: OWL 2 DL forbids it. Graph checks ",
        "one domain per property": "reject it"}),
    (range(21, 22), "CORRECTION", ["CORRECTION   ", "DataProperty:", " BM_DEP  ", "Range:", " xsd:integer", "    |    ",
                                   "ObjectProperty:", " hasPatientData  ", "Range:", " PatientData"]),
    # Step 5, semantic card: duplicate class, said plainly.
    (range(20, 22), "already has PatientData, now", {"Valid OWL, but the ": "A new class, but the ", ", now ": ": ",
                                                    "two classes": "two names", " mean the ": " for the ",
                                                    ". The ": ". Valid OWL, so no rule complains. The "}),
    (range(16, 17), "patient_id,BM_DEP", {"admission → query: admission + patient_id,BM_DEP, ecog_stage+ PatientData":
                                          "admission → query: admission + patient_id, BM_DEP, ecog_stage + PatientData"}),
    (range(41, 42), "Quality dimensions", {
        "Quality dimensions: accuracy, completeness, conciseness, adaptability, clarity, consistency, ":
            "Quality dimensions: accuracy, completeness, conciseness, adaptability, clarity, consistency "
            "[Vrandečić, Handbook on Ontologies 2009]; ",
        "domain enrichment ": "domain enrichment, adapted from ontology enrichment ",
        "[Vrandečić, Handbook on Ontologies 2009; OntoQA: Tartir et al., KADASH 2005].": "[Petasis et al., LNCS 2011]."}),
    (range(32, 35), "3 LLM METHOD × 3 DATABASE × 3 LLM", {"METHOD": "METHODS", "DATABASE": "DATABASES", "LLM": "LLMS"}),
    (range(33, 35), "1 METHOD × 3 DATABASE =", {"DATABASE": "DATABASES"}),
    (range(42, 44), "tools in the loop", {"Next: Move the evaluation tools in the loop ":
                                          "Next: Move the evaluation tools into the loop"}),
]


def fix(tree, phys):
    for span, where, runs in FIXES:
        if phys not in span:
            continue
        hit = False
        for p in tree.iter(f"{{{A}}}p"):
            rs = p.findall(".//a:r", NS)
            if where not in "".join(r.findtext("a:t", namespaces=NS) or "" for r in rs):
                continue
            hit = True
            if isinstance(runs, list):  # positional replacement
                assert len(runs) == len(rs), (phys, where)
                for r, new in zip(rs, runs):
                    r.find("a:t", NS).text = new
                continue
            for i, r in enumerate(rs):
                t = r.find("a:t", NS)
                if t.text in runs and not (t.text == "LLM" and i == 1):  # keep the first 'LLM' of '3 LLM METHOD'
                    t.text = runs[t.text]
        assert hit, (phys, where)


# ---- slide numbers -----------------------------------------------------------------------------

def number(tree, n, arch=False):
    root = tree.getroot().find("p:cSld/p:spTree", NS)
    if arch:  # move the reveal counter under the progress dashes (top right)
        for el in shapes(tree):
            if re.fullmatch(r"\d / 6", text(el).strip()):
                set_geom(el, x=12.52 - geom(el)[2], y=0.70)
                resize(el, {1400: 1100})
    rpr = etree.fromstring(f'<a:rPr xmlns:a="{A}" lang="en-US" sz="1100" b="0" dirty="0"><a:solidFill>'
                           f'<a:srgbClr val="{MUTED}"/></a:solidFill><a:latin typeface="Consolas"/>'
                           f'<a:cs typeface="Consolas"/></a:rPr>')
    nid = max(int(sid(el)) for el in shapes(tree) if sid(el)) + 1
    root.append(textbox(nid, "Slide number", 12.50, 7.08, 0.45, 0.26, str(n), rpr, "r"))


# ---- driver ------------------------------------------------------------------------------------

EDITS = {range(3, 9): slides_3_to_8, range(17, 19): slides_17_18, range(12, 13): slide_12,
         range(36, 40): results_tables, range(41, 42): results_tables, range(40, 41): burr,
         range(44, 46): future_data, range(19, 22): validate_cards}


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    order = speech_final.slide_order(work)
    talk_no = speech_final.talk_numbers()
    last = talk_no[-1]
    for phys, (path, n) in enumerate(zip(order, talk_no), 1):
        tree = etree.parse(str(path))
        fix(tree, phys)
        for span, fn in EDITS.items():
            if phys in span:
                fn(tree)
        if phys in range(40, 41):
            ladder_only(tree)
        if phys == 41:
            judges_footer(tree)
        if 1 < n < last:
            number(tree, n, arch=speech_final.TALK[n - 1]["title"].startswith("Reassembling"))
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    speech_final.apply_notes(work)
    speech_final.repack(work, dst)
    shutil.rmtree(work)
    speech_final.summary()


if __name__ == "__main__":
    main(*sys.argv[1:])
