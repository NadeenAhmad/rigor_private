#!/usr/bin/env python3
"""Phase 6: backup slides B2-B6 after B1 (untimed, for Q&A).

B2 expert scores per quality dimension (paper Table 11)   B3 retrieval ablation (Table 14)
B4 cost and runtime (Table 15, totals computed from rows) B5-B6 hard questions (speech.py)

Usage: python3 backups.py <in.pptx> <out.pptx>
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
from results import ADD_SLIDE, BODY, EMERALD, INK, MUTED, NS, TEAL, Ids, add, delete, run, set_runs, shape  # noqa: E402
from speech import HARD_QUESTIONS  # noqa: E402

KEEP = {946, 947, 948}  # eyebrow, title, subtitle of the B1 template

DIMS = ["Accuracy", "Completeness", "Conciseness", "Adaptability", "Clarity", "Consistency",
        "Domain enrichment", "Average"]
METHODS = ["Direct Map.", "LLM baseline", "Non-Iter.", "RIGOR"]
TABLE11 = {  # rows = DIMS, columns = METHODS
    "Liver cancer": [(3.75, 3.30, 2.35, 3.75), (2.75, 3.00, 2.25, 3.85), (4.05, 3.65, 2.95, 3.60),
                     (2.10, 1.95, 1.55, 2.10), (2.50, 2.45, 2.90, 3.90), (3.30, 2.85, 2.35, 3.35),
                     (2.20, 2.00, 2.35, 3.50), (2.95, 2.74, 2.39, 3.44)],
    "eICU-CRD": [(3.05, 3.90, 0.85, 4.00), (2.45, 3.10, 0.80, 4.10), (3.65, 3.80, 2.05, 3.80),
                 (1.90, 2.10, 1.00, 2.20), (2.55, 2.75, 1.50, 4.10), (2.60, 3.40, 1.30, 3.15),
                 (1.90, 2.00, 1.15, 3.35), (2.59, 3.01, 1.24, 3.53)],
}
TABLE14 = [  # variant, schema F1, external F1, docs F1, avg precision, avg recall, avg F1
    ("no retrieval", 1.0000, 0.0328, 0.9188, 0.7722, 0.6224, 0.6505),
    ("only schema + core", 0.9965, 0.0328, 0.9158, 0.7576, 0.6224, 0.6484),
    ("only external ontologies", 1.0000, 0.0326, 0.9188, 0.7631, 0.6224, 0.6505),
    ("only documentation", 1.0000, 0.0347, 0.9209, 0.7687, 0.6240, 0.6519),
    ("full RIGOR", 0.9929, 0.0589, 0.9139, 0.7565, 0.6299, 0.6552),
]
TABLE15 = [  # database, model, requests, tokens, cost USD, runtime min
    ("eICU-CRD", "Claude Opus 4.6", 168, 1_660_000, 21.30, 182.4),
    ("eICU-CRD", "Mistral-Small-24B", 96, 1_290_000, 0.0831, 188.4),
    ("eICU-CRD", "DeepSeek-V3", 52, 334_000, 0.1760, 20.6),
    ("Liver cancer", "Claude Opus 4.6", 70, 731_000, 10.30, 94.4),
    ("Liver cancer", "Mistral-Small-24B", 34, 246_000, 0.0154, 20.5),
    ("Liver cancer", "DeepSeek-V3", 33, 142_000, 0.0747, 23.8),
    ("Chinook", "Claude Opus 4.6", 24, 119_000, 1.28, 9.0),
    ("Chinook", "Mistral-Small-24B", 53, 344_000, 0.02, 65.0),
    ("Chinook", "DeepSeek-V3", 22, 71_000, 0.03, 12.0),
]


def table(tree, ids, x, y, widths, header, rows, row_h=0.34, size=1150, bold_cells=(), emerald_col=None,
          left_cols=1):
    """Header + rows of text cells; alternate rows tinted; bold_cells = {(r, c)} set in emerald bold."""
    add(tree, shape(ids(), "Table header rule", x, y + row_h, sum(widths), 0.01, fill="D9DEDB"))
    for r, cells in enumerate([header] + rows):
        cy = y + r * row_h + (0.02 if r else 0)
        if r and r % 2 == 0:
            add(tree, shape(ids(), "Row tint", x, cy, sum(widths), row_h, fill="F3F5F4"))
        cx = x
        for c, (text, w) in enumerate(zip(cells, widths)):
            if r == 0:
                xml = run(text, TEAL, 1000, True, "Consolas")
            else:
                strong = (r - 1, c) in bold_cells
                color = EMERALD if strong or (c == emerald_col and c) else INK
                xml = run(text, color, size, strong or (c == emerald_col and c > 0))
            add(tree, shape(ids(), f"Cell {r}-{c}", cx, cy, w, row_h, xml, align="l" if c < left_cols else "r",
                            anchor="ctr", inset=0.08))
            cx += w


def strip(tree):
    delete(tree, *[int(e.get("id")) for e in tree.iter(f"{{{NS['p']}}}cNvPr")
                   if e.get("id") and int(e.get("id")) not in KEEP and int(e.get("id")) > 1
                   and e.getparent().getparent().tag != f"{{{NS['p']}}}spTree"])


def header(tree, eyebrow, title, subtitle):
    set_runs(tree, 946, [(eyebrow, None, None)])
    set_runs(tree, 947, [(title, None, None)])
    set_runs(tree, 948, [(subtitle, None, None)])


def caption(tree, ids, text, y=6.62):
    add(tree, shape(ids(), "Backup caption", 0.93, y, 11.47, 0.26, run(text, MUTED, 1050)))


def b2(tree, ids):
    header(tree, "BACKUP  /  B2", "Experts score RIGOR highest on average in both databases.",
           "Human expert scores (0–5) per quality dimension  ·  10 experts, 2 tables per database  ·  "
           "best value per row in bold")
    for i, (db, rows) in enumerate(TABLE11.items()):
        x = 0.90 + i * 5.85
        add(tree, shape(ids(), "DB label", x, 2.05, 5.6, 0.30, run(db, EMERALD, 1200, True, "Consolas")))
        bold = {(r, c + 1) for r, vals in enumerate(rows) for c, v in enumerate(vals) if v == max(vals)}
        table(tree, ids, x, 2.42, [1.70, 0.98, 0.98, 0.98, 0.98], ["DIMENSION"] + [m.upper() for m in METHODS],
              [[d] + [f"{v:.2f}" for v in vals] for d, vals in zip(DIMS, rows)], row_h=0.42, bold_cells=bold)
    caption(tree, ids, "Where RIGOR is not best: Conciseness (Direct Mapping on liver cancer), Consistency "
                       "(LLM baseline on eICU-CRD).  Source: paper Table 11; LLM-judge scores in Table 10.")


def b3(tree, ids):
    header(tree, "BACKUP  /  B3", "Full retrieval gives the best overall alignment, by a small margin.",
           "Retrieval ablation on Chinook with Claude  ·  F1 of alignment between the ontology and each "
           "resource (cosine ≥ 0.55)")
    rows = [[v, f"{s:.3f}", f"{e:.3f}", f"{d:.3f}", f"{p:.3f}", f"{r:.3f}", f"{f:.3f}"]
            for v, s, e, d, p, r, f in TABLE14]
    best = {}
    for c in range(1, 7):
        col = [float(r[c]) for r in rows]
        best.update({(i, c): True for i, v in enumerate(col) if v == max(col)})
    table(tree, ids, 0.90, 2.20, [3.10, 1.40, 1.40, 1.40, 1.40, 1.40, 1.40],
          ["VARIANT", "SCHEMA F1", "EXTERNAL F1", "DOCS F1", "AVG P", "AVG R", "AVG F1"], rows,
          row_h=0.52, size=1300, bold_cells=set(best))
    add(tree, shape(ids(), "Ablation note", 0.90, 5.55, 11.50, 0.62,
                    run("Read with care: ", EMERALD, 1300, True)
                    + run("schema alignment is saturated because every variant keeps the direct-mapping seed, and "
                          "external-ontology F1 is low in absolute terms. Full RIGOR is best on external F1, "
                          "average recall and average F1.", BODY, 1300),
                    geom="roundRect", fill="E6F2EC", anchor="ctr", inset=0.20))
    caption(tree, ids, "Source: paper Table 14 (Chinook, Claude). One database and one model.")


def b4(tree, ids):
    header(tree, "BACKUP  /  B4", "All nine RIGOR runs together cost about 33 dollars.",
           "Generation cost and runtime per database and model, including Gen-LLM and Judge-LLM calls")
    totals = [sum(r[i] for r in TABLE15) for i in (2, 3, 4, 5)]
    rows = [[db, m, f"{req}", f"{tok:,}", f"${cost:,.2f}", f"{mins:.1f}"] for db, m, req, tok, cost, mins in TABLE15]
    rows.append(["Total", "", f"{totals[0]}", f"{totals[1]:,}", f"${totals[2]:,.2f}", f"{totals[3]:.1f}"])
    table(tree, ids, 0.90, 2.05, [2.30, 2.70, 1.50, 1.90, 1.50, 1.60],
          ["DATABASE", "MODEL", "REQUESTS", "TOKENS", "COST (USD)", "RUNTIME (MIN)"], rows,
          row_h=0.40, size=1200, bold_cells={(len(rows) - 1, c) for c in range(6)}, left_cols=2)
    caption(tree, ids, "Source: paper Table 15. Totals are summed from the nine rows; the paper's printed "
                       "totals (651 requests, $34.28) also include runs not listed in the table.")


def hard(tree, ids, part, items):
    header(tree, f"BACKUP  /  B{4 + part}", f"Hard questions ({part}/2)",
           "Short, honest answers  ·  full versions in presentation/SPEECH.md")
    for i, (q, a) in enumerate(items):
        a = re.sub(r"^VERIFY IN THE CODE BEFORE THE TALK\. Expected answer: (\w)", lambda m: m.group(1).upper(), a)
        y = 1.98 + i * 0.92
        add(tree, shape(ids(), f"Q card {i}", 0.90, y, 11.50, 0.84, geom="roundRect", fill="F3F5F4", line="D9DEDB"))
        add(tree, shape(ids(), f"Q {i}", 1.10, y + 0.07, 11.10, 0.28, run(q, INK, 1250, True)))
        add(tree, shape(ids(), f"A {i}", 1.10, y + 0.35, 11.10, 0.46, run(a, BODY, 1100)))


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    after, made = "slide24.xml", []
    for _ in range(5):
        out = subprocess.run([sys.executable, str(ADD_SLIDE), str(work), "slide24.xml", "--after", after],
                             capture_output=True, text=True, check=True).stdout
        after = re.search(r"(slide\d+\.xml)", out.split("Created")[-1]).group(1)
        made.append(after)
    builders = [b2, b3, b4, lambda t, i: hard(t, i, 1, HARD_QUESTIONS[:5]), lambda t, i: hard(t, i, 2, HARD_QUESTIONS[5:])]
    for name, build in zip(made, builders):
        path = work / "ppt/slides" / name
        tree = etree.parse(str(path))
        strip(tree)
        build(tree, Ids(tree))
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print("wrote", dst, "(backups:", ", ".join(made) + ")")


if __name__ == "__main__":
    main(*sys.argv[1:3])
