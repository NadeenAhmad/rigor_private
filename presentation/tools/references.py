#!/usr/bin/env python3
"""References on slides: author-year citations under the claims they support.

Slide 4 (research gap): the schema-first and LLM-first approaches the slide calls limited,
all cited in the RIGOR paper's Section 2.

Usage: python3 references.py <in.pptx> <out.pptx>
"""
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from code_excerpts import add, move, next_id, para, run_xml, simple_box  # noqa: E402

MUTED, SIZE = "5B6770", 1000
SLIDE4 = {
    # card x, card width, citation runs (name, reference)
    "schema": (0.88, 5.63, [("W3C Direct Mapping", "Arenas et al. 2012; Sequeda et al. 2012"),
                            ("BootOX", "Jiménez-Ruiz et al., ISWC 2015"),
                            ("Ben Mahria et al.", "J. Big Data 2021"),
                            ("RODI benchmark", "Pinkel et al., SWJ 2017")]),
    "llm": (6.79, 5.67, [("LLMs4OL", "Babaei Giglou et al., ISWC 2023"),
                         ("Mateiu & Groza", "SYNASC 2023"),
                         ("OntoGenix", "Val-Calvo et al., IP&M 2025"),
                         ("LLM4VKG", "Xiao et al., IJCAI 2025")]),
}


def citation_runs(items):
    runs = []
    for k, (name, ref) in enumerate(items):
        if k:
            runs.append(run_xml("  ·  ", MUTED, False, SIZE, font="Calibri"))
        nb = "\u00a0"  # keep each citation on one line; wrap only at the separators
        keep = {" ": nb, "-": "\u2011"}  # non-breaking space and hyphen
        runs.append(run_xml("".join(keep.get(c, c) for c in name), MUTED, True, SIZE, font="Calibri"))
        runs.append(run_xml("".join(keep.get(c, c) for c in f" [{ref}]"), MUTED, False, SIZE, font="Calibri"))
    return "".join(runs)


def edit_slide4(tree):
    shift = 0.43  # make room under the cards
    for sid, y in ((123, 5.19), (124, 5.38), (125, 5.32)):
        move(tree, sid, y=y + shift)
    for x, w, items in SLIDE4.values():
        box = simple_box(next_id(tree), "References", x + 0.04, 4.95, w - 0.08, 0.50,
                         paras=para(citation_runs(items)))
        box.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr").set("anchor", "t")
        add(tree, box)


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    path = work / "ppt/slides/slide4.xml"
    tree = etree.parse(str(path))
    edit_slide4(tree)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
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
