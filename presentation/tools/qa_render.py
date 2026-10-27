#!/usr/bin/env python3
"""Render a deck to JPEGs for visual QA.

LibreOffice has no Consolas, so it substitutes DejaVu Sans Mono, which is about 9% wider
(0.602 em vs 0.550 em). To keep overflow checks honest, this renders a QA-only copy in
which every Consolas run is scaled by 0.55/0.60. The real deck is not modified.

Usage: python3 qa_render.py deck.pptx out_dir [dpi]
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

SOFFICE = Path("/root/.claude/skills/synced/954bdbed-f7d8-4e99-9a67-2334e0f54e1f_96ca6ae3-be58-43c5-a447-6fed00e21488"
               "/pptx/scripts/office/soffice.py")
RATIO = 0.550 / 0.602


def scale_mono(xml):
    def fix(m):
        rpr = m.group(0)
        if 'typeface="Consolas"' not in rpr:
            return rpr
        return re.sub(r'sz="(\d+)"', lambda s: f'sz="{int(int(s.group(1)) * RATIO)}"', rpr, count=1)
    return re.sub(r"<a:rPr\b[^>]*>.*?</a:rPr>", fix, xml, flags=re.S)


def main(deck, out_dir, dpi="110"):
    out = Path(out_dir)
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    qa = out / "qa.pptx"
    with zipfile.ZipFile(deck) as zin, zipfile.ZipFile(qa, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r"ppt/slides/slide\d+\.xml$", item.filename):
                data = scale_mono(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    subprocess.run([sys.executable, str(SOFFICE), "--headless", "--convert-to", "pdf", "--outdir", str(out), str(qa)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(["pdftoppm", "-jpeg", "-r", dpi, str(out / "qa.pdf"), str(out / "s")], check=True)
    print("\n".join(sorted(str(p) for p in out.glob("s-*.jpg"))))


if __name__ == "__main__":
    main(*sys.argv[1:])
