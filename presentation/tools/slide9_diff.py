#!/usr/bin/env python3
"""Slide 9 (step 4, generate) on the author's edited slide: BEFORE and AFTER share one row grid.

The left panel shows the direct mapping with empty rows; the right panel repeats it row for row,
fills the empty rows with the Gen-LLM's insertions and highlights renamed names. Inserted and
renamed lines are highlighted and tagged with their evidence source; unchanged lines stay plain.
Text comes from code_excerpts.DIFF_ROWS (verified with the OWL API). The speaker notes are
rewritten from speech.py.

Usage: python3 slide9_diff.py <in.pptx> <out.pptx>   (edits the slide titled "STEP 4 / GENERATE")
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import (DIFF_ROWS, EMU, GLOSS_COLORS, NS, P, add, code_line, move, next_id, para,  # noqa: E402
                           run_xml, simple_box)
from results import set_notes  # noqa: E402
from speech import notes_for, timeline  # noqa: E402

TAG_LABELS = {"schema": "SCHEMA", "docs": "DOCS", "vocab": "ONTOLOGY", "llm": "LLM"}
NEW_FILL = "DCF2E6"            # mint: content the Gen-LLM inserted or renamed
SIZE = 1100                    # code size, both panels
CHAR_W = 0.5498 * SIZE / 100 / 72  # Consolas advance width in inches
LINE_INSET = 0.16              # left inset of code_line text boxes

PANEL_TOP, PANEL_BOTTOM = 2.02, 6.12
CODE_TOP, CODE_BOTTOM, ROW_PAD = 2.46, 6.00, 0.08


def geometry(sp):
    off, ext = sp.find(".//a:xfrm/a:off", NS), sp.find(".//a:xfrm/a:ext", NS)
    if off is None or ext is None:
        return None
    return tuple(int(v) / EMU for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")))


def text(sp):
    return "".join(t.text or "" for t in sp.iter(f"{{{NS['a']}}}t")).strip()


def shape_id(sp):
    return sp.find(".//p:cNvPr", NS).get("id")


def set_first_run(sp, value):
    runs = sp.findall(".//a:r", NS)
    runs[0].find("a:t", NS).text = value
    for r in runs[1:]:
        r.getparent().remove(r)


def find_slide(work):
    for path in sorted((work / "ppt/slides").glob("slide*.xml")):
        if "STEP 4 / GENERATE" in path.read_text(encoding="utf-8"):
            return path
    raise SystemExit("no slide titled STEP 4 / GENERATE")


def edit(tree):
    tree_root = tree.find(".//p:spTree", NS)
    shapes = [el for el in tree_root if el.tag in (f"{{{P}}}sp", f"{{{P}}}cxnSp")]
    by_text = {text(sp): sp for sp in shapes}
    subtitle = next(sp for t, sp in by_text.items() if t.startswith("We prompt the"))
    before = next(sp for t, sp in by_text.items() if t.startswith("BEFORE"))
    after = next(sp for t, sp in by_text.items() if t.startswith("AFTER"))
    gen_label = by_text["Gen-LLM"]
    # The two outer panels: tall boxes below the subtitle; the right one is the widest right of centre
    # (the old inner code box is also tall, but narrower).
    panels = [sp for sp in shapes if (g := geometry(sp)) and g[2] > 3 and g[3] > 3 and g[1] > 1.5]
    left = min(panels, key=lambda sp: geometry(sp)[0])
    right = max((sp for sp in panels if geometry(sp)[0] > 4), key=lambda sp: geometry(sp)[2])
    lx, ly, lw, lh = geometry(left)
    rx, ry, rw, rh = geometry(right)

    # Remove everything inside the two panels (code, tags, caption) except the two headers.
    keep = {id(before), id(after), id(left), id(right)}
    for sp in shapes:
        g = geometry(sp)
        if g and id(sp) not in keep and any(px <= g[0] and g[0] + g[2] <= px + pw + 0.01 and
                                            py <= g[1] and g[1] + g[3] <= py + ph + 0.01
                                            for px, py, pw, ph in ((lx, ly, lw, lh), (rx, ry, rw, rh))):
            tree_root.remove(sp)

    # Frame: subtitle up a little, taller panels, headers, arrow and label centred on the panels.
    move(tree, shape_id(subtitle), y=1.62)
    for panel in (left, right):
        move(tree, shape_id(panel), y=PANEL_TOP, h=PANEL_BOTTOM - PANEL_TOP)
    set_first_run(before, "BEFORE  /  DIRECT MAPPING")
    move(tree, shape_id(before), w=lw - 0.40)
    for header in (before, after):
        move(tree, shape_id(header), y=PANEL_TOP + 0.13)
    mid = (PANEL_TOP + PANEL_BOTTOM) / 2
    arrow = next(sp for sp in tree_root if (g := geometry(sp)) and lx + lw < g[0] < rx and g[3] > 0.3
                 and ly < g[1] < ly + lh and sp is not gen_label)
    move(tree, shape_id(arrow), y=mid - 0.33)
    move(tree, shape_id(gen_label), y=mid + 0.12)

    # Code boxes at identical heights, one row grid for both.
    boxes = ((lx + 0.16, lw - 0.32), (rx + 0.33, rw - 0.60))
    for bx, bw in boxes:
        add(tree, simple_box(next_id(tree), "Code panel", bx, CODE_TOP, bw, CODE_BOTTOM - CODE_TOP,
                             fill="F6F8FA", line="D0D7DE", geom="roundRect"))
    pitch = (CODE_BOTTOM - CODE_TOP - 2 * ROW_PAD) / len(DIFF_ROWS)
    (lbx, lbw), (rbx, rbw) = boxes
    tag_w, tag_h = 1.08, min(0.19, pitch - 0.02)
    tag_x = rbx + rbw - tag_w - 0.12
    y = CODE_TOP + ROW_PAD
    for left_code, right_code, kind, group in DIFF_ROWS:
        if left_code:
            add(tree, code_line(next_id(tree), lbx, y, lbw - 0.05, pitch, left_code, "", None, size=SIZE))
        if kind == "added":  # the whole row is new
            add(tree, simple_box(next_id(tree), "Inserted row", rbx + 0.06, y + 0.005, rbw - 0.12, pitch - 0.01,
                                 fill=NEW_FILL))
            add(tree, simple_box(next_id(tree), "Insertion mark", rbx + 0.06, y + 0.005, 0.04, pitch - 0.01,
                                 fill="0B6E4F"))
        elif kind == "changed":  # only the renamed name is new
            old, new = left_code.split(), right_code.split()
            word = next(n for o, n in zip(old, new) if o != n)
            col = right_code.index(word)
            add(tree, simple_box(next_id(tree), "Renamed name", rbx + LINE_INSET + col * CHAR_W - 0.03,
                                 y + 0.005, len(word) * CHAR_W + 0.06, pitch - 0.01, fill=NEW_FILL))
        add(tree, code_line(next_id(tree), rbx, y, tag_x - rbx - 0.05, pitch, right_code, "", None, size=SIZE))
        if group:
            color = GLOSS_COLORS[group]
            add(tree, simple_box(next_id(tree), "Evidence tag", tag_x, y + (pitch - tag_h) / 2, tag_w, tag_h,
                                 fill="FFFFFF", line=color, geom="roundRect",
                                 paras=para(run_xml(TAG_LABELS[group], color, True, 850), algn="ctr")))
        y += pitch


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    path = find_slide(work)
    tree = etree.parse(str(path))
    edit(tree)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    rels = (path.parent / "_rels" / f"{path.name}.rels").read_text(encoding="utf-8")
    notes = re.search(r'Target="\.\./notesSlides/(notesSlide\d+\.xml)"', rels)
    if notes:
        n, s, start, end, _ = next(row for row in timeline() if row[1]["title"] == "Step 4: generate")
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
