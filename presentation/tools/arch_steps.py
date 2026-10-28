#!/usr/bin/env python3
"""Architecture build-up 1/6 to 5/6, derived from the author's 6/6 slide.

The author re-aligned the full method figure on the "6 / 6" slide. This script duplicates that
slide five times, inserts the copies just before it, and on copy k removes everything that
belongs to later steps, so every box keeps exactly the author's position. The current step's box
gets the author's highlight (the 6/6 Integrate style: emerald 2.3 pt border, DDE5E1 fill; white
boxes keep their fill), and the title, subtitle, progress dashes, page number and speaker notes
follow the step.

Shape ids are those of the author's slide (RIGOR_ISWC2026_3.pptx); the script stops if any is
missing, e.g. after a re-export renumbers them.

Usage: python3 arch_steps.py <in.pptx> <out.pptx>
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import A, NS, set_text, sp_by_id  # noqa: E402
from results import set_notes  # noqa: E402
from speech import notes_for, timeline  # noqa: E402

PPTX_SCRIPTS = Path("/root/.claude/skills/synced/954bdbed-f7d8-4e99-9a67-2334e0f54e1f_"
                    "96ca6ae3-be58-43c5-a447-6fed00e21488/pptx/scripts")
DONE, UPCOMING = "0B6E4F", "D3DAD6"   # progress dashes
HIGHLIGHT_LINE, HIGHLIGHT_FILL, BOX_FILL = "0B6E4F", "DDE5E1", "F3F5F4"

# Which step each part of the figure belongs to (ids on the author's 6/6 slide).
PARTS = {
    1: {998, 999, 1030, 1031, 1032, 1033, 1034, 1035, 1036, 1037, 1038, 1039, 1040, 1041, 1076, 1077, 1078},
    2: {1000, 1075, 1001, 1002, 1080, 1003, 1042, 1043, 1044, 1045},
    3: {1004, 1008, 1011, 1079, 1005, 1006, 1007, 1009, 1010, 1012, 1013, 1014, *range(1046, 1060),
        1086, 1087, 1088},
    4: {1015, 1081, 1016, 1017, 1018, 987, 1019, 1020, 1021, 1022, 986, 1060, 1061, 1062, 1063, 1089,
        1023, 1024, 1064},
    5: {1065, 1066, 1067, 1068, 1084, 1085},
    6: {1025, 1026, 1069, 1070, 1071, 1072, 1027, 1028, 1082, 1083, 1029, 1074, 1073},
}
CURRENT = {1: (1030, 1076), 2: (1042,), 3: (1046,), 4: (1060,), 5: (1065,)}
TITLE, SUBTITLE, PAGE, DASHES = 989, 990, 997, (991, 992, 993, 994, 995, 996)
STEPS = {
    1: ("1. Follow the foreign keys", "Start where the foreign-key dependencies begin."),
    2: ("2. Anchor generation in the schema", "The schema-grounded branch feeds the generation step."),
    3: ("3. Retrieve the missing semantics", "Context is retrieved independently from three sources."),
    4: ("4. Generate a delta ontology", "The Gen-LLM receives the direct mapping and the retrieved context."),
    5: ("5. Validate before integration", "Validate and refine each delta before it joins the core."),
}


def by_id(tree):
    return {int(c.get("id")): c.getparent().getparent() for c in tree.iter(f"{{{NS['p']}}}cNvPr")}


def set_fill(sp, color):
    fill = sp.find("p:spPr/a:solidFill/a:srgbClr", NS)
    fill.set("val", color)


def highlight(sp):
    ln = sp.find("p:spPr/a:ln", NS)
    ln.set("w", "29200")
    ln.find("a:solidFill/a:srgbClr", NS).set("val", HIGHLIGHT_LINE)
    if sp.find("p:spPr/a:solidFill/a:srgbClr", NS).get("val") == BOX_FILL:
        set_fill(sp, HIGHLIGHT_FILL)


def edit(tree, k):
    shapes = by_id(tree)
    missing = set().union(*PARTS.values(), {TITLE, SUBTITLE, PAGE}, DASHES) - set(shapes)
    if missing:
        raise SystemExit(f"step-6 slide is missing shape ids {sorted(missing)}")
    for step in range(k + 1, 7):
        for sid in PARTS[step]:
            shapes[sid].getparent().remove(shapes[sid])
    for sid in CURRENT[k]:
        highlight(shapes[sid])
    title, subtitle = STEPS[k]
    set_text(tree, TITLE, title)
    set_text(tree, SUBTITLE, " " + subtitle)  # the author's subtitle starts with a space
    set_text(tree, PAGE, f"{k} / 6")
    for i, sid in enumerate(DASHES, 1):
        set_fill(shapes[sid], DONE if i <= k else UPCOMING)


def add_notes(work, slide_name, source_notes, k):
    """Give a duplicated slide its own notes slide (a copy of the 6/6 notes, rewritten)."""
    notes_dir = work / "ppt/notesSlides"
    n = max(int(re.findall(r"\d+", p.name)[0]) for p in notes_dir.glob("notesSlide*.xml")) + 1
    dest = notes_dir / f"notesSlide{n}.xml"
    shutil.copy(source_notes, dest)
    rels = (notes_dir / "_rels" / f"{source_notes.name}.rels").read_text(encoding="utf-8")
    rels = re.sub(r'Target="\.\./slides/slide\d+\.xml"', f'Target="../slides/{slide_name}"', rels)
    (notes_dir / "_rels" / f"{dest.name}.rels").write_text(rels, encoding="utf-8")
    ct = work / "[Content_Types].xml"
    ct.write_text(ct.read_text(encoding="utf-8").replace(
        "</Types>", f'<Override PartName="/ppt/notesSlides/{dest.name}" ContentType="application/'
                    f'vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'), encoding="utf-8")
    slide_rels = work / "ppt/slides/_rels" / f"{slide_name}.rels"
    text = slide_rels.read_text(encoding="utf-8")
    slide_rels.write_text(text.replace(
        "</Relationships>", '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/'
                            f'2006/relationships/notesSlide" Target="../notesSlides/{dest.name}"/></Relationships>'),
        encoding="utf-8")
    n_, s, start, end, _ = next(row for row in timeline() if row[1]["title"] == f"Architecture {k}/6")
    set_notes(dest, notes_for(n_, s, start, end))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    slides = sorted((work / "ppt/slides").glob("slide*.xml"), key=lambda p: int(re.findall(r"\d+", p.name)[0]))
    six = next(p for p in slides if "REASSEMBLING THE ARCHITECTURE" in (t := p.read_text(encoding="utf-8"))
               and "6 / 6" in t)
    source_notes = work / "ppt/notesSlides" / re.search(
        r'notesSlides/(notesSlide\d+\.xml)', (work / "ppt/slides/_rels" / f"{six.name}.rels").read_text()).group(1)
    # Insert the copies before the 6/6 slide: after the slide that precedes it.
    pres = (work / "ppt/presentation.xml").read_text(encoding="utf-8")
    rels = {m[0]: m[1] for m in re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
                                           (work / "ppt/_rels/presentation.xml.rels").read_text())}
    order = [rels[r] for r in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', pres)]
    after = order[order.index(six.name) - 1]
    for k in range(1, 6):
        out = subprocess.run([sys.executable, str(PPTX_SCRIPTS / "add_slide.py"), str(work), six.name,
                              "--after", after], capture_output=True, text=True, cwd=PPTX_SCRIPTS, check=True)
        new = re.findall(r"(slide\d+\.xml)", out.stdout)[0]
        path = work / "ppt/slides" / new
        tree = etree.parse(str(path))
        edit(tree, k)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
        add_notes(work, new, source_notes, k)
        after = new
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
