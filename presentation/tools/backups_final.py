#!/usr/bin/env python3
"""Backup slides for Q&A, appended after the closing slide of the final deck.

B0 is a clickable index; every backup links back to it. Each backup answers one question we expect,
with the evidence on the slide and the spoken answer in the notes. Facts come from the paper
(arXiv:2506.01232v2, tables named on each slide) and, where the paper is silent, from the RIGOR code
(github.com/NadeenAhmad/RIGOR: app.py, eval.py, expert_survey/, cqs/, output/).

Usage: python3 backups_final.py <deck.pptx> <out.pptx>
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from arch_steps import PPTX_SCRIPTS  # noqa: E402
from results import BODY, EMERALD, INK, MUTED, NS, P, RED, TEAL, Ids, add, set_notes  # noqa: E402
from results_ladder import HEADER, place, set_text  # noqa: E402
from slide9_diff import text as sp_text  # noqa: E402

A = NS["a"]
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400
LINE, PANEL, STRIP, GOOD, BAD = "D9DEDB", "F3F5F4", "E6F2EC", "DDF0E5", "F6DAD3"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"


# ---- drawing helpers ---------------------------------------------------------------------------

def run(t, color=INK, size=1400, bold=False, italic=False, font="Calibri", link=None):
    hl = f'<a:hlinkClick r:id="{link}" action="ppaction://hlinksldjump"/>' if link else ""
    return (f'<a:r><a:rPr lang="en-US" sz="{size}" b="{int(bold)}" i="{int(italic)}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{font}"/>'
            f'<a:cs typeface="{font}"/>{hl}</a:rPr><a:t xml:space="preserve">{escape(t)}</a:t></a:r>')


def para(runs, align="l", after=0, bullet=False):
    bu = ('<a:buFont typeface="Arial"/><a:buChar char="•"/>' if bullet else "<a:buNone/>")
    ind = ' marL="228600" indent="-228600"' if bullet else ' marL="0" indent="0"'
    return (f'<a:p><a:pPr algn="{align}"{ind}><a:spcBef><a:spcPts val="0"/></a:spcBef>'
            f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>{bu}</a:pPr>{"".join(runs)}</a:p>')


def box(ids, x, y, w, h, paras, anchor="t", fill=None, line=None, geom="rect", inset=0.0, name="Text"):
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    line_xml = (f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
                else "<a:ln><a:noFill/></a:ln>")
    adj = '<a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst>' if geom == "roundRect" else "<a:avLst/>"
    ins = int(inset * EMU)
    body = "".join(paras) if paras else "<a:p><a:endParaRPr lang=\"en-US\"/></a:p>"
    txbox = ' txBox="1"' if paras and not fill and not line else ""
    i = ids()
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvSpPr><p:cNvPr id="{i}" name="{name} {i}"/>'
        f'<p:cNvSpPr{txbox}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
        f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm><a:prstGeom prst="{geom}">{adj}</a:prstGeom>'
        f'{fill_xml}{line_xml}</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="{ins}" tIns="{ins // 2}" '
        f'rIns="{ins}" bIns="{ins // 2}" anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>{body}'
        f'</p:txBody></p:sp>')


class Slide:
    def __init__(self, tree, label):
        self.tree, self.ids, self.label = tree, Ids(tree), label

    def put(self, *a, **k):
        add(self.tree, box(self.ids, *a, **k))

    def card(self, x, y, w, h, label, fill="FFFFFF"):
        self.put(x, y, w, h, [], fill=fill, line=LINE, geom="roundRect", name="Card")
        self.put(x + 0.25, y + 0.16, w - 0.5, 0.3, [para([run(label, TEAL, 1300, True, font="Consolas")])],
                 name="Card label")

    def bullets(self, x, y, w, h, items, size=1300, after=500):
        """items: str or list of (text, bold) segments."""
        ps = []
        for it in items:
            segs = [(it, False)] if isinstance(it, str) else it
            ps.append(para([run(t, INK if b else BODY, size, b) for t, b in segs], after=after, bullet=True))
        self.put(x, y, w, h, ps, name="Bullets")

    def strip(self, y, segs, h=0.46):
        self.put(0.88, y, 11.57, h, [para([run(t, EMERALD if b else BODY, 1350, b) for t, b in segs])],
                 anchor="ctr", fill=STRIP, geom="roundRect", inset=0.22, name="Strip")

    def source(self, line):
        self.put(0.90, 6.98, 10.2, 0.26, [para([run(line, MUTED, 1000)])], name="Source")

    def grid(self, x, y, widths, rows, row_h=0.36, size=1250, head=True, fills=None, aligns=None):
        """rows: lists of str or (str, color, bold). fills: {row index: fill}."""
        for r, cells in enumerate(rows):
            yy = y + r * row_h
            fill = (fills or {}).get(r, PANEL if (r % 2 == 0 and not (head and r == 0)) else None)
            if fill:
                self.put(x, yy, sum(widths), row_h, [], fill=fill, name="Row")
            xx = x
            for c, (w, cell) in enumerate(zip(widths, cells)):
                t, color, bold = (cell, INK, False) if isinstance(cell, str) else cell
                if head and r == 0:
                    color, bold, sz, font = TEAL, True, 1100, "Consolas"
                else:
                    sz, font = size, "Calibri"
                al = (aligns or ["l"] + ["ctr"] * (len(widths) - 1))[c]
                self.put(xx + 0.08, yy, w - 0.16, row_h, [para([run(t, color, sz, bold, font=font)], align=al)],
                         anchor="ctr", name="Cell")
                xx += w
            if head and r == 0:
                self.put(x, yy + row_h - 0.01, sum(widths), 0.012, [], fill="9AA3A9", name="Rule")


def g(t):  # emerald bold cell
    return (t, EMERALD, True)


def red(t):
    return (t, RED, True)


def b(t):
    return (t, INK, True)


# ---- the backups -------------------------------------------------------------------------------

def b1(s):
    s.grid(0.88, 2.42, [2.35, 4.85, 4.37], [
        ["BIAS", "WHAT WE DID", "WHAT REMAINS"],
        [b("Self-preference"), "Judge is GPT-5.4, a different model family from all three generators",
         "Inside RIGOR, the Judge-LLM is the generator's own model (graph checks are model-free)"],
        [b("Position"), "Fragments anonymized as A–D",
         red("Order was fixed (RIGOR always D), for judge and experts: re-judge with shuffled order")],
        [b("Verbosity"), "Scored per dimension, not only ranked: the judge marks RIGOR down on conciseness "
                         "where it is long (Chinook 2.59 vs 3.32)",
         "Clarity and domain enrichment reward annotations by design"],
        [b("Leniency"), "Ten human experts, blind, same rubric: RIGOR 1st in 70%, significant agreement",
         "Judge is harsher than experts on the LLM baseline (eICU 0.28 vs 3.01)"],
    ], row_h=0.74, size=1200, aligns=["l", "l", "l"])
    s.strip(6.30, [("No LLM involved: ", False), ("OOPS! pitfalls 8 vs 52–499, highest BURR F1 on all three. ", True),
                   ("Both agree with the judges.", False)])
    s.source("Paper Sec. 4.4, 4.6, Tables 1, 7, 10–13; fragment order: Fig. 7 and expert_survey/METHOD_KEY.json.")


def b2(s):
    s.card(0.88, 2.42, 5.65, 3.70, "HOW WE USED THEM")
    s.bullets(1.13, 2.90, 5.20, 2.00, [
        [("5 per table, 290 in total", True), (", generated from the schema (Mistral Small 24B), verified by a "
                                                "domain expert", False)],
        [("Never shown to the generator", True), (", so they stay an independent test", False)],
        [("Given to judge and experts", True), (" with the schema and the four anonymized fragments: "
                                                 "can this ontology answer them?", False)],
    ], size=1250)
    s.put(1.13, 5.05, 5.15, 0.90, [para([run("Example (liver, general_aftercare)", MUTED, 1100, italic=True)]),
                                   para([run("Which patients developed bile-duct stenosis after surgery, and what "
                                             "treatment was documented?", INK, 1250, True)])],
          fill=PANEL, geom="roundRect", inset=0.12, name="Example")
    s.card(6.80, 2.42, 5.65, 3.70, "WHY NOT RUN THEM AS SPARQL (YET)")
    s.bullets(7.05, 2.90, 5.20, 3.10, [
        [("No instance data: ", True), ("the ontologies are schemas (TBox), so data questions return empty "
                                         "for every method", False)],
        [("It would test the wrong thing: ", True), ("answerability mostly checks that columns are kept "
                                                      "(Result 2), not labels, value meanings or hierarchy", False)],
        [("A second error source: ", True), ("4 × 290 hand- or LLM-written queries", False)],
    ], size=1250)
    s.strip(6.30, [("Next: ", True), ("populate a KG through RIGOR's provenance mapping (already exported for BURR) "
                                      "and run the CQs as SPARQL.", False)])
    s.source("Paper Sec. 4.2, 4.6 (Strategies 6–7), App. Table 6; cqs/ in the RIGOR repository.")


def b3(s):
    s.grid(0.88, 2.42, [2.30, 1.55, 3.70, 1.20, 1.00], [
        ["TABLE", "DATABASE", "WHY THIS TABLE", "KENDALL'S W", "P"],
        [b("general_aftercare"), "Liver", "clinical follow-up, several foreign keys", "0.59", "< 0.001"],
        [b("complication"), "Liver", "encoded severity grades, part-whole", "0.39", "< 0.01"],
        [b("hospital"), "eICU", "compact administrative entity", "0.59", "< 0.001"],
        [b("patient"), "eICU", "hub referenced by most tables", "0.72", "< 0.001"],
    ], row_h=0.44, aligns=["l", "l", "l", "ctr", "ctr"])
    s.card(10.85, 2.42, 1.60, 2.20, "RIGOR 1ST", fill=STRIP)
    s.put(10.95, 2.90, 1.40, 1.60, [para([run("14 / 20", EMERALD, 2000, True)], align="ctr", after=200),
                                    para([run("liver", MUTED, 1150)], align="ctr", after=600),
                                    para([run("14 / 20", EMERALD, 2000, True)], align="ctr", after=200),
                                    para([run("eICU", MUTED, 1150)], align="ctr")], name="Firsts")
    s.bullets(0.98, 4.85, 11.40, 1.30, [
        [("Who: ", True), ("ten experts in ontology engineering and knowledge representation (profiles in paper "
                           "Fig. 8)", False)],
        [("Task: ", True), ("per table, schema + 5 CQs + four anonymized fragments; score 7 dimensions 0–5, rank "
                            "1st–4th. Biggest RIGOR gains: clarity and domain enrichment", False)],
        [("Why 4 tables, Claude only: ", True), ("survey length (4 fragments × 7 scores + a ranking per table); "
                                                 "the LLM judge covered all 58 tables", False)],
    ], size=1250, after=300)
    s.source("Paper Sec. 4.6 (Strategy 7), Tables 1, 11, 12, Fig. 8.")


def b4(s):
    s.grid(0.88, 2.42, [3.15, 4.70, 3.72], [
        ["ROLE", "MODEL", "SAFEGUARD"],
        [b("Ontology generation"), "Claude Opus 4.6 · Mistral Small 24B · DeepSeek-V3", "same models for all three LLM methods"],
        [b("Judge-LLM inside RIGOR"), "the generator's own model", "plus deterministic graph checks"],
        [b("Documentation"), "GPT-4o", "reviewed by domain experts"],
        [b("Competency questions"), "Mistral Small 24B", "verified by a domain expert"],
        [b("Evaluation judge"), "GPT-5.4", "different family from every generator"],
        [b("External ontologies"), "ICD-10, Disease, Cell, Nutrition · Music, Performed Music, DOREMUS",
         "chosen by domain experts"],
    ], row_h=0.44, size=1200, aligns=["l", "l", "l"])
    s.strip(5.75, [("Data leakage: ", True), ("the liver-cancer registry is private, so no model has seen it, and the "
                                              "judge still ranked RIGOR first on all 16 of its tables.", False)], h=0.62)
    s.source("Paper Sec. 4.1, 4.3, 4.4, Tables 1, 4.")


def b5(s):
    s.card(0.88, 2.42, 5.65, 3.55, "DIRECT MAPPING ON CHINOOK: 78%")
    s.bullets(1.13, 2.90, 5.20, 3.00, [
        [("43 ordinary columns: ", True), ("a property with exactly the same name, so all are matched", False)],
        [("21 key columns: ", True), ("primary keys become HasKey axioms, not properties; foreign keys are "
                                      "renamed (ArtistId → hasArtist, ReportsTo → hasEmployee)", False)],
        [("The metric compares names, ", True), ("so most key columns fall below 0.55. Nothing was lost.", False)],
    ], size=1250)
    s.card(6.80, 2.42, 5.65, 3.55, "RIGOR WITH DEEPSEEK ON eICU: 77%")
    s.bullets(7.05, 2.90, 5.20, 3.00, [
        [("Real omissions: ", True), ("the final ontology has 197 data properties; the seed had 299", False)],
        [("Why: ", True), ("in the paper's pipeline only the validated delta is merged, so a delta that renames or "
                           "drops seed properties loses columns", False)],
        [("Fixed: ", True), ("the current RIGOR code merges the direct mapping into the core first, so every "
                             "column stays", False)],
    ], size=1250)
    s.strip(6.20, [("Coverage: ", True), ("a column counts as covered when some ontology class or property name has "
                                          "embedding similarity ≥ 0.55 (all-MiniLM-L6-v2).", False)])
    s.source("Paper Tables 8, 9, Algorithm 1 (line 22); output/generated_ontologies/direct_mapping_chinook.owl, app.py.")


def b6(s):
    s.grid(0.88, 2.42, [3.20, 2.80, 2.80, 2.77], [
        ["AVERAGE PER ONTOLOGY", "ALL PITFALLS", "WITHOUT P08", "CRITICAL (P19)"],
        ["Direct Mapping", "499", "250", "15"],
        ["LLM baseline", "307", "34", "17"],
        ["Non-Iterative RAG", "52", "29", "5.4"],
        [g("RIGOR"), g("8"), g("8"), g("2.8")],
    ], row_h=0.50, size=1400, fills={4: STRIP})
    s.bullets(0.98, 5.10, 11.40, 1.00, [
        [("P08 = missing annotations: ", True), ("72% of the other methods' pitfalls. Without it, RIGOR still has "
                                                  "the fewest, about 4 times fewer than the next method", False)],
        [("P19 = multiple domains or ranges, ", True), ("the only critical pitfall found: zero in 6 of RIGOR's 9 "
                                                        "ontologies", False)],
    ], size=1250, after=300)
    s.source("Paper Table 7, averaged over 3 (Direct Mapping) or 9 ontologies; global pitfalls without counts "
             "(P10, P22) excluded.")


def b7(s):
    s.grid(0.88, 2.42, [3.40, 2.05, 2.05, 2.05, 2.02], [
        ["VARIANT (CHINOOK, CLAUDE)", "SCHEMA", "EXT. ONTOLOGIES", "DOCUMENTATION", "AVERAGE"],
        ["No retrieval", "1.000", "0.033", "0.919", "0.651"],
        ["Schema + core only", "0.997", "0.033", "0.916", "0.648"],
        ["External ontologies only", "1.000", "0.033", "0.919", "0.651"],
        ["Documentation only", "1.000", "0.035", g("0.921"), "0.652"],
        [g("Full RIGOR"), "0.993", g("0.059"), "0.914", g("0.655")],
    ], row_h=0.46, size=1350, fills={5: STRIP})
    s.bullets(0.98, 5.35, 11.40, 0.80, [
        [("Full RIGOR aligns best overall ", True), ("and nearly doubles external-ontology alignment; each source "
                                                     "helps its own resource", False)],
        [("Why the margins are small: ", True), ("every variant keeps the direct-mapping seed, so schema alignment is "
                                                 "saturated, and Chinook's names are self-explanatory", False)],
    ], size=1250, after=300)
    s.source("Paper Sec. 4.8, Tables 5, 14. F1 of matches (cosine ≥ 0.55) between ontology elements and each "
             "resource.")


def b8(s):
    s.grid(0.88, 2.42, [3.05, 2.13, 2.13, 2.13, 2.13], [
        ["ADDS ONE STEP EACH", "DIRECT MAPPING", "+ LLM", "+ RETRIEVAL", "+ SEED, CORE, CHECKS"],
        ["", ("rules only", MUTED, False), ("LLM baseline", MUTED, False), ("Non-Iterative RAG", MUTED, False),
         g("RIGOR")],
        ["Columns covered", "91%", "74%", red("34%"), g("93%")],
        ["OOPS! pitfalls", "499", "307", "52", g("8")],
        ["Labels & comments", "0", "0", "389", g("777")],
        ["Provenance links", "0", "0", "0", g("196")],
        ["Judge: ranked 1st", "0 / 58", "0 / 58", "0 / 58", g("58 / 58")],
    ], row_h=0.42, size=1300)
    s.strip(5.55, [("Reading: ", True), ("retrieval adds labels but drops columns; the seed brings the columns back; "
                                         "the checks cut the errors.", False)])
    s.put(0.90, 6.15, 11.5, 0.60, [para([run("Not isolated: seed, growing core and validation were added together. A "
                                             "leave-one-component-out ablation is future work.", MUTED, 1200,
                                             italic=True)])], name="Caveat")
    s.source("Paper Sec. 4.5 (controlled spectrum, App. Table 3), Tables 1, 7, 8, 9; averages per ontology.")


def b9(s):
    levels = [("LEVEL 0", ["Artist", "Employee", "Genre", "MediaType", "Playlist"]),
              ("LEVEL 1", ["Album", "Customer"]), ("LEVEL 2", ["Track", "Invoice"]),
              ("LEVEL 3", ["InvoiceLine", "PlaylistTrack"])]
    s.card(0.88, 2.42, 5.65, 3.75, "CHINOOK, IN PROCESSING ORDER")
    for k, (lab, tables) in enumerate(levels):
        x = 1.13 + k * 1.33
        s.put(x, 2.92, 1.20, 0.26, [para([run(lab, MUTED, 1050, True, font="Consolas")], align="ctr")])
        for j, t in enumerate(tables):
            s.put(x, 3.25 + j * 0.50, 1.20, 0.40, [para([run(t, INK, 1150, True)], align="ctr")], anchor="ctr",
                  fill=STRIP if k == 0 else PANEL, line=LINE, geom="roundRect")
    s.put(1.13, 5.72, 5.20, 0.35, [para([run("Track waits for Album, Genre and MediaType.", BODY, 1200, italic=True)])])
    s.card(6.80, 2.42, 5.65, 3.75, "WHY")
    s.bullets(7.05, 2.90, 5.20, 3.20, [
        [("Reuse: ", True), ("the classes a table points to already exist, so the LLM links to their IRIs instead "
                             "of inventing new ones", False)],
        [("Context: ", True), ("the parents' labels and comments are in the growing core, ready to retrieve", False)],
        [("Reproducible: ", True), ("level by level, ties broken alphabetically", False)],
        [("Cycles: ", True), ("a self-reference (Employee.ReportsTo) does not block the order and becomes a recursive "
                              "relation; tables in a longer cycle are appended at the end", False)],
    ], size=1200, after=300)
    s.strip(6.32, [("Any parents-first order works; ", True), ("we chose breadth-first for level-by-level growth. "
                                                               "We did not compare orders.", False)])
    s.source("Paper Sec. 3.1–3.2 (FK-BFS); app.py traverse_schema_fk_order; schema_chinook.json.")


def b10(s):
    s.card(0.88, 2.42, 7.30, 3.80, "JUDGE-LLM: 14 CRITERIA")
    s.bullets(1.13, 2.90, 6.85, 3.25, [
        [("Critical: ", True), ("FK never a data property; no URI both object and data property; exactly one "
                                "domain and range; valid XSD ranges; no self-loops unless recursive; keep the "
                                "direct mapping's coverage and key restrictions", False)],
        [("Important: ", True), ("English labels and comments; meanings of coded values; junction tables as "
                                 "relations; reuse core concepts; disjointness only when safe", False)],
        [("Minor: ", True), ("external alignments; provenance on every element", False)],
    ], size=1200, after=400)
    s.card(8.40, 2.42, 4.05, 3.80, "GRAPH CHECKS (RULES)")
    s.bullets(8.65, 2.90, 3.60, 3.25, [
        "Type consistency", "One domain and one range per property", "XSD-only data ranges", "Full schema coverage",
    ], size=1250, after=500)
    s.strip(6.36, [("On failure: ", True), ("issues go back to the Gen-LLM (k = 2 retries); then the Judge-LLM keeps "
                                            "the better of its last version and the direct mapping.", False)])
    s.source("Paper Sec. 3.5, App. A.2 (Fig. 3), Algorithm 1.")


def b11(s):
    s.grid(0.88, 2.42, [2.60, 3.10, 5.87], [
        ["CHOICE", "VALUE", "WHY"],
        [b("Seed"), "W3C Direct Mapping", "every column is in the ontology before enrichment [Sequeda et al., WWW 2012]"],
        [b("Order"), "foreign keys, breadth-first", "parents before children (backup B9)"],
        [b("Embeddings"), "all-MiniLM-L6-v2", "strong on MTEB, light enough to re-index every iteration"],
        [b("Retrieved items"), "k = 3 per source", "enough context without crowding out the seed and instructions"],
        [b("Match threshold"), "cosine ≥ 0.55", "empirically tuned, as in prior work [Galli et al. 2024; Fathallah et al. 2024]"],
        [b("Retries"), "k = 2", "self-refinement gains fade after two rounds [Madaan et al. 2023; Shinn et al. 2023]"],
        [b("Syntax"), "OWL 2 Manchester", "readable for experts, generated reliably by LLMs"],
    ], row_h=0.48, size=1200, aligns=["l", "l", "l"])
    s.source("Paper Sec. 3.2–3.5, 4.6.")


def b12(s):
    s.grid(0.88, 2.42, [2.65, 2.97, 2.97, 2.98], [
        ["DATABASE", "CLAUDE OPUS 4.6", "MISTRAL SMALL 24B", "DEEPSEEK-V3"],
        [b("eICU-CRD · 31 tables"), "$21.30 · 182 min", "$0.08 · 188 min", "$0.18 · 21 min"],
        [b("Liver Cancer · 16"), "$10.30 · 94 min", "$0.02 · 21 min", "$0.07 · 24 min"],
        [b("Chinook · 11"), "$1.28 · 9 min", "$0.02 · 65 min", "$0.03 · 12 min"],
    ], row_h=0.50, size=1300)
    s.strip(4.47, [("All nine runs: ", True), ("651 requests · 5.3 M tokens · $34.28 · 11.4 hours", False)])
    s.bullets(0.98, 5.15, 11.40, 1.10, [
        [("Scales with the number of tables: ", True), ("one prompt per table, and k = 3 items per source keep each "
                                                        "prompt bounded as the core grows", False)],
        [("Privacy: ", True), ("only the schema and documentation are sent, never patient rows; Mistral Small 24B "
                               "can run on-premises and reaches 95–99% coverage with RIGOR", False)],
    ], size=1250, after=300)
    s.source("Paper Table 15 (runtime includes Gen-LLM and Judge-LLM calls), Table 9.")


def b13(s):
    s.card(0.88, 2.42, 5.10, 3.80, "SETUP")
    s.bullets(1.13, 2.90, 4.65, 3.25, [
        "ISWC scenario: 9 tables, 46 columns, 11 foreign keys",
        "RIGOR with Claude; documentation by GPT-5.4; SWRC and BIBO as external ontologies",
        [("Mappings come from provenance: ", True), ("prov:wasDerivedFrom exported as a D2RQ mapping", False)],
        "F1 against the gold mapping, per element type",
    ], size=1200, after=400)
    s.grid(6.20, 2.42, [2.55, 1.20, 1.20, 1.30], [
        ["SYSTEM", "CLASSES", "RELATIONS", "ATTRIBUTES"],
        ["W3C-Mapper / Ontop", "0.53", "0.11", "0.39"],
        ["RDB2Onto", "0.67", "0.62", "0.44"],
        ["OntoGenix", "0.58", "0.35", "0.44"],
        [g("RIGOR"), g("0.73"), g("0.67"), g("0.52")],
        [("Gemini 2.5 Pro*", MUTED, False), ("0.65", MUTED, False), ("0.68", MUTED, False), ("0.56", MUTED, False)],
    ], row_h=0.46, size=1300, fills={4: STRIP})
    s.put(6.20, 5.30, 6.25, 0.85, [para([run("* In BURR, a general LLM prompted with examples to write the mapping "
                                             "directly: a mapping generator, not an ontology-learning system. "
                                             "RIGOR also builds labels, hierarchy and provenance.", MUTED, 1150,
                                             italic=True)])], name="Footnote")
    s.source("Paper Sec. 4.6 (Strategy 8), Table 13; Gemini: BURR [Laskowski et al., PACMMOD 2025], Table 3.")


def b14(s):
    items = [
        ("COVERAGE 0.77–1.00", "A delta can drop seed properties.", "Merge the direct mapping first (RIGOR App)."),
        ("RESIDUAL P19", "Critical pitfall in 3 of 9 RIGOR runs.", "Validators inside the loop (RIGOR App)."),
        ("DISJOINTNESS, ALIGNMENTS", "13 disjointness axioms in one run; external F1 ≤ 0.06.",
         "Couple with a reasoner and entity linking."),
        ("EVALUATION DESIGN", "Fixed fragment order; experts on 4 tables, Claude only.",
         "Shuffled re-judging; more tables and models."),
        ("ABLATION", "One database, one LLM; seed, core and checks not isolated.",
         "Leave-one-component-out on all databases."),
        ("SCHEMA ONLY", "No instance data, so CQs are judged, not executed.", "Let the data speak (future work)."),
    ]
    w, h = 3.73, 1.85
    for k, (lab, issue, fix) in enumerate(items):
        x, y = 0.88 + (k % 3) * (w + 0.19), 2.42 + (k // 3) * (h + 0.22)
        s.card(x, y, w, h, lab)
        s.put(x + 0.25, y + 0.55, w - 0.5, 0.65, [para([run(issue, INK, 1250)])])
        s.put(x + 0.25, y + 1.20, w - 0.5, 0.55, [para([run("→ " + fix, EMERALD, 1250, True)])])
    s.source("Paper Sec. 4.7–5, Tables 7–9, 14.")


BACKUPS = [
    # (key, group, question title, one-line answer, builder, spoken answer)
    ("B1", "EVALUATION", "Could the LLM judge be biased?",
     "Possibly, which is why it is never our only evidence: four independent checks point the same way.", b1,
     "Yes, LLM judges have known biases, so we never rely on the judge alone. Self-preference: the judge is "
     "GPT-5.4, a different family from all three generators. Verbosity: it scores seven dimensions, and it does "
     "mark RIGOR down on conciseness where RIGOR is long, on Chinook. Position: the fragments were anonymized, but "
     "the order was fixed, with RIGOR always last, for the judge and the experts; shuffled re-judging is the "
     "obvious next check. And the evidence without any LLM, OOPS! pitfalls and the BURR benchmark, points the same "
     "way as both judges."),
    ("B2", "EVALUATION", "How were the CQs used? Why not answer them in SPARQL?",
     "As a shared yardstick for blind judging: what each table's ontology must be able to answer.", b2,
     "Five competency questions per table, 290 in total, generated from the schema and verified by a domain "
     "expert. They were never shown to the generator, so they are an independent test. The judge and the experts "
     "read them next to the four anonymized fragments and asked: can this ontology answer them? We did not execute "
     "them as SPARQL for three reasons: the ontologies have no instance data, so every method would return empty; "
     "answerability would mostly test that columns are kept, which Result 2 already shows, not labels or value "
     "meanings; and writing over a thousand queries adds its own errors. With RIGOR's provenance mapping we can "
     "populate a knowledge graph and run them, and that is the natural next step."),
    ("B3", "EVALUATION", "Who were the experts, and do they agree?",
     "Ten ontology engineers, blind, on four tables; agreement is significant on every table.", b3,
     "Ten experts in ontology engineering and knowledge representation. Four tables chosen for different modeling "
     "challenges, two per medical database. Per table they saw the schema, five CQs and four anonymized fragments, "
     "scored seven dimensions and ranked them. Agreement, Kendall's W, is between 0.39 and 0.72 and significant on "
     "every table. RIGOR was first in 14 of 20 rankings on each database. Four tables because each one is a long "
     "survey task; the LLM judge covered all 58."),
    ("B4", "EVALUATION", "Which model did what?",
     "Generation, documentation, questions and judging are separated; no model grades its own work.", b4,
     "Three generators, Claude, Mistral and DeepSeek, each used in all three LLM methods. Inside RIGOR the "
     "Judge-LLM is the generator's own model, but deterministic graph checks back it up. Documentation came from "
     "GPT-4o and was reviewed by domain experts; the CQs came from Mistral and were verified by an expert; the "
     "evaluation judge is GPT-5.4, a different family. On leakage: the liver-cancer registry is private, no model "
     "saw it, and RIGOR still ranked first on all 16 of its tables."),
    ("B5", "RESULTS", "Why is coverage below 100%?",
     "Two different reasons: a name-based metric for Direct Mapping, real omissions in one RIGOR run.", b5,
     "Two different stories. Direct Mapping on Chinook loses nothing: all ordinary columns have a property with "
     "the same name. The misses are key columns: primary keys become HasKey axioms, not properties, and foreign "
     "keys get new names like hasArtist, and the metric compares names. RIGOR with DeepSeek on eICU is a real "
     "omission: the final ontology has 197 data properties against 299 in the seed, because in the paper's "
     "pipeline only the validated delta is merged. The current RIGOR code merges the direct mapping first, so every "
     "column stays."),
    ("B6", "RESULTS", "Aren't the OOPS! counts just missing labels?",
     "No: without P08, RIGOR still has the fewest pitfalls, and the fewest critical ones.", b6,
     "Most of the other methods' pitfalls are indeed P08, missing annotations, about 72 percent. Remove P08 and "
     "RIGOR still has the fewest: 8 per ontology against 29 for the next method and 250 for direct mapping, whose "
     "rest is mostly missing domains and ranges. And for the only critical pitfall, P19, multiple domains or "
     "ranges, RIGOR has 2.8 per ontology and zero in six of nine."),
    ("B7", "RESULTS", "Did you run an ablation?",
     "Yes, for the retrieval sources: all sources together align best, but the margins are small.", b7,
     "Yes, for the three retrieval sources, on Chinook with Claude. Full RIGOR has the best average alignment and "
     "nearly doubles the alignment with external ontologies; documentation alone gives the best documentation "
     "alignment. The margins are small for two reasons: every variant keeps the direct-mapping seed, so schema "
     "alignment is saturated, and Chinook's names are self-explanatory, so retrieval matters less than on a "
     "clinical schema with names like BM_DEP."),
    ("B8", "RESULTS", "Which part of RIGOR matters most?",
     "Our four methods add one step at a time, so each step's effect is visible.", b8,
     "The four methods form a ladder, each adding one capability. Adding retrieval in one pass adds labels but "
     "loses columns: coverage drops to a third. Adding the seed, the growing core and the checks brings coverage to "
     "93 percent and cuts pitfalls from 52 to 8. Honestly, those three were added together, so their separate "
     "effects are not isolated; a leave-one-component-out ablation is future work."),
    ("B9", "METHOD", "Why breadth-first foreign-key order?",
     "Parents before children, so every table can reuse the classes it points to.", b9,
     "Because a table should be generated after the tables it points to. When we reach Track, Album, Genre and "
     "MediaType already exist, so the LLM links to their IRIs and retrieves their labels instead of inventing "
     "duplicates. Breadth-first gives that order level by level, with alphabetical ties, so runs are reproducible. "
     "A self-reference like ReportsTo does not block anything and becomes a recursive relation. Any parents-first "
     "order would satisfy the dependency; we did not compare orders."),
    ("B10", "METHOD", "What exactly does validation check?",
     "Fourteen Judge-LLM criteria in three severities, plus four deterministic graph checks.", b10,
     "Two stages. The Judge-LLM checks fourteen criteria in three severities: critical ones protect the structure, "
     "like one domain and one range per property and keeping the direct mapping's coverage; important ones cover "
     "labels, value meanings, reuse and safe disjointness; minor ones cover alignments and provenance. Then rules "
     "check type consistency, domains and ranges, XSD ranges and coverage on the parsed graph. Failures go back with "
     "the issues listed, up to two retries."),
    ("B11", "METHOD", "Why these settings?",
     "Each setting trades context against noise; values follow the literature or a small tuning run.", b11,
     "The seed guarantees coverage. MiniLM is strong on MTEB and cheap enough to re-index the core after every "
     "table. Three items per source give context without crowding out the seed and the instructions. The 0.55 "
     "threshold was tuned and matches prior work. Two retries because self-refinement gains fade after two rounds. "
     "Manchester syntax because experts can read it and LLMs generate it reliably."),
    ("B12", "PRACTICE", "What does it cost, and does it scale?",
     "All nine RIGOR runs: about $34 and 11 hours; the small open model costs cents.", b12,
     "All nine RIGOR runs together cost about 34 dollars and 11 hours, most of it Claude. Mistral Small, a 24B "
     "open model, costs cents and still reaches 95 to 99 percent coverage. Cost grows with the number of tables: "
     "one prompt per table, and retrieval keeps each prompt bounded as the core grows. Only schema and "
     "documentation leave the hospital, and the small model can run on-premises."),
    ("B13", "PRACTICE", "BURR: setup, and why not compare with Gemini?",
     "ISWC scenario, mappings recovered from RIGOR's provenance; Gemini is a mapping generator.", b13,
     "We used BURR's ISWC scenario. RIGOR's provenance annotations are exported as a D2RQ mapping and scored "
     "against the gold mapping. Gemini 2.5 Pro in BURR is a general LLM prompted to write the mapping directly: a "
     "mapping generator, not an ontology-learning system. Against it, RIGOR's classes are higher, relations on par, "
     "attributes lower, and RIGOR also builds labels, hierarchy and provenance."),
    ("B14", "PRACTICE", "What are the limitations?",
     "What we know is missing, and what we would do about it.", b14,
     "Coverage can drop below the seed, which merging the direct mapping first fixes. A critical pitfall remains "
     "in three of nine runs, which validators in the loop address. Disjointness and external alignments are "
     "sparse. The evaluation used a fixed fragment order and four expert tables. The ablation covers one database "
     "and one LLM. And we work on schemas only, which is where instance data comes in."),
]


# ---- slide plumbing ----------------------------------------------------------------------------

def order(work):
    pres = (work / "ppt/presentation.xml").read_text(encoding="utf-8")
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"',
                           (work / "ppt/_rels/presentation.xml.rels").read_text()))
    return [rels[x] for x in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', pres)]


def copy_slide(work, src, after):
    out = subprocess.run([sys.executable, str(PPTX_SCRIPTS / "add_slide.py"), str(work), src, "--after", after],
                         capture_output=True, text=True, cwd=PPTX_SCRIPTS, check=True)
    return re.findall(r"(slide\d+\.xml)", out.stdout)[0]


def link(work, slide, target):
    """Add a slide-jump relationship from `slide` to `target`; return its id."""
    path = work / "ppt/slides/_rels" / f"{slide}.rels"
    xml = path.read_text(encoding="utf-8")
    n = 1 + max(int(x) for x in re.findall(r'Id="rId(\d+)"', xml))
    path.write_text(xml.replace("</Relationships>", f'<Relationship Id="rId{n}" Type="{SLIDE_REL}" '
                                                    f'Target="{target}"/></Relationships>'), encoding="utf-8")
    return f"rId{n}"


def reset(tree, eyebrow, title, subtitle):
    root = tree.find(".//p:spTree", NS)
    shapes = [el for el in root if etree.QName(el).localname in ("sp", "cxnSp", "pic", "grpSp", "graphicFrame")]
    eb = next(sp for sp in shapes if sp_text(sp).startswith("RESULT"))
    tt = shapes[shapes.index(eb) + 1]
    for sp in shapes:
        if sp is not eb and sp is not tt:
            root.remove(sp)
    set_text(eb, eyebrow)
    place(eb, *HEADER["eyebrow"])
    set_text(tt, title)
    place(tt, *HEADER["title"])
    for rp in tt.iter(f"{{{A}}}rPr"):
        rp.set("sz", "2600")


def add_notes(work, slide, paragraphs):
    rels_path = work / "ppt/slides/_rels" / f"{slide}.rels"
    notes_dir = work / "ppt/notesSlides"
    template = sorted(notes_dir.glob("notesSlide*.xml"))[0]
    n = max(int(re.findall(r"\d+", p.name)[0]) for p in notes_dir.glob("notesSlide*.xml")) + 1
    dest = notes_dir / f"notesSlide{n}.xml"
    shutil.copy(template, dest)
    rels = (notes_dir / "_rels" / f"{template.name}.rels").read_text(encoding="utf-8")
    (notes_dir / "_rels" / f"{dest.name}.rels").write_text(
        re.sub(r'Target="\.\./slides/slide\d+\.xml"', f'Target="../slides/{slide}"', rels), encoding="utf-8")
    ct = work / "[Content_Types].xml"
    ct.write_text(ct.read_text(encoding="utf-8").replace(
        "</Types>", f'<Override PartName="/ppt/notesSlides/{dest.name}" ContentType="application/vnd.'
                    'openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'), encoding="utf-8")
    rels_path.write_text(rels_path.read_text(encoding="utf-8").replace(
        "</Relationships>", '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/'
                            f'2006/relationships/notesSlide" Target="../notesSlides/{dest.name}"/></Relationships>'),
        encoding="utf-8")
    set_notes(dest, paragraphs)


def main(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    template = next(n for n in order(work) if "RESULT 1" in (work / "ppt/slides" / n).read_text(encoding="utf-8"))
    after = order(work)[-1]
    slides = []
    for _ in range(len(BACKUPS) + 1):
        after = copy_slide(work, template, after)
        slides.append(after)
    index, backups = slides[0], slides[1:]

    # B1..B14
    for name, (key, group, q, answer, build, spoken) in zip(backups, BACKUPS):
        path = work / "ppt/slides" / name
        tree = etree.parse(str(path))
        reset(tree, f"BACKUP {key}  /  {group}", q, answer)
        s = Slide(tree, key)
        s.put(*HEADER["subtitle"], [para([run(answer, BODY, 1600)])], name="Subtitle")
        build(s)
        back = link(work, name, index)
        s.put(10.45, 7.06, 1.95, 0.28, [para([run("↩ All questions", TEAL, 1100, font="Consolas", link=back)],
                                             align="r")], name="Back link")
        s.put(12.45, 7.06, 0.50, 0.28, [para([run(key, "6B747B", 1100, font="Consolas")], align="r")],
              name="Slide number")
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
        add_notes(work, name, [f"BACKUP {key}: {q}", "", f"SPOKEN ANSWER: {spoken}", "",
                               f"ON THE SLIDE: {answer}"])

    # B0 index
    path = work / "ppt/slides" / index
    tree = etree.parse(str(path))
    reset(tree, "BACKUP  /  QUESTIONS", "Questions we expect", "")
    s = Slide(tree, "B0")
    s.put(*HEADER["subtitle"], [para([run("Click a question to jump to its answer; every answer links back here.",
                                          BODY, 1600)])], name="Subtitle")
    groups = ["EVALUATION", "RESULTS", "METHOD", "PRACTICE"]
    w, h = 5.69, 2.05
    for k, grp in enumerate(groups):
        x, y = 0.88 + (k % 2) * (w + 0.19), 2.42 + (k // 2) * (h + 0.20)
        s.card(x, y, w, h, grp)
        rows = [(key, q, name) for name, (key, g_, q, *_r) in zip(backups, BACKUPS) if g_ == grp]
        ps = []
        for key, q, name in rows:
            rid = link(work, index, name)
            ps.append(para([run(f"{key}  ", TEAL, 1200, True, font="Consolas", link=rid),
                            run(q, INK, 1300, link=rid)], after=400))
        s.put(x + 0.25, y + 0.55, w - 0.45, h - 0.65, ps, name="Links")
    s.put(12.45, 7.06, 0.50, 0.28, [para([run("B0", "6B747B", 1100, font="Consolas")], align="r")],
          name="Slide number")
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    add_notes(work, index, ["BACKUP INDEX (not part of the timed talk).", "",
                            "Each question links to its backup slide; '↩ All questions' on every backup returns "
                            "here. In slideshow mode the links work with a click."])

    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print(f"wrote {dst}: index + {len(BACKUPS)} backups")


if __name__ == "__main__":
    main(*sys.argv[1:3])
