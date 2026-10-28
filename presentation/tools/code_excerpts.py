#!/usr/bin/env python3
"""Phase 3: replace pseudo-notation ontology excerpts with real OWL 2 Manchester Syntax.

Every excerpt is verified with the OWL API before it goes on a slide: the text shown on
the slide is exactly the text parsed (plus hidden prefix and declaration lines).

Usage:
  python3 code_excerpts.py verify <owlapi_classpath>    # parse every excerpt, write presentation/excerpts/
  python3 code_excerpts.py apply <in.pptx> <out.pptx>   # put the excerpts on slides 3, 7, 9 and 10

The OWL API checker (OwlCheck.java) lives next to this script; compile it with
  javac -cp "<owlapi_classpath>" OwlCheck.java
"""
import copy
import re
import subprocess
import sys
import zipfile
import tempfile
import shutil
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
EMU = 914400
HERE = Path(__file__).resolve().parent
EXCERPT_DIR = HERE.parent / "excerpts"

HEADER = """Prefix: : <http://example.org/rigor#>
Prefix: rdfs: <http://www.w3.org/2000/01/rdf-schema#>
Prefix: xsd: <http://www.w3.org/2001/XMLSchema#>
Prefix: prov: <http://www.w3.org/ns/prov#>
Prefix: src: <http://example.org/provenance/admission/>
Ontology: <http://example.org/rigor>
AnnotationProperty: prov:wasDerivedFrom
Class: PatientData
Individual: src:ecog_stage
"""

# Text shown on the slides. "\t" separates code from its plain-English gloss.
SEED = """Class: Admission
  SubClassOf:
    patient_id some PatientData,\t← each has a patient
    BM_DEP only xsd:integer,\t← if present, an integer
    ecog_stage only xsd:integer
ObjectProperty: patient_id\t← the FK becomes a link
DataProperty: BM_DEP
DataProperty: ecog_stage"""

SEED_COMPACT = """Class: Admission
  SubClassOf:
    patient_id some PatientData,
    BM_DEP only xsd:integer,
    ecog_stage only xsd:integer"""

# Slide 3 target: one line per entry (code, gloss, knowledge source). The reveal adds one
# source per click: schema, documentation, standard vocabulary, LLM knowledge.
TARGET_LINES = [
    ("Class: Admission", "", "schema"),
    ("  SubClassOf: hasPatientData some PatientData", "schema (foreign key)", "schema"),
    ("  DisjointWith: PatientData", "LLM knowledge", "llm"),
    ("DataProperty: BM_DEP", "", "schema"),
    ("  Annotations:", "", "schema"),
    ('    rdfs:label "Bone Marrow Depression"@en,', "documentation", "docs"),
    ("    skos:relatedMatch obo:DOID_4961,", "Disease Ontology", "vocab"),
    ("    prov:wasDerivedFrom admission:BM_DEP", "source column", "schema"),
    ("DataProperty: ecog_stage", "", "schema"),
    ("  Annotations:", "", "schema"),
    ('    rdfs:label "ECOG Performance Status"@en,', "documentation", "docs"),
    ('    rdfs:comment "0 = fully active … 5 = dead"@en,', "LLM knowledge", "llm"),
    ("    prov:wasDerivedFrom admission:ecog_stage", "source column", "schema"),
]
TARGET_GROUP_BREAKS = (2, 7)  # a little extra space after these line indices
REVEAL_ORDER = ("schema", "docs", "vocab", "llm")
TARGET = "\n".join(code for code, _, _ in TARGET_LINES)
TARGET_HEADER = (HEADER.replace("Ontology:", "Prefix: skos: <http://www.w3.org/2004/02/skos/core#>\n"
                                "Prefix: obo: <http://purl.obolibrary.org/obo/>\n"
                                "Prefix: admission: <http://example.org/provenance/admission/>\nOntology:")
                 + "AnnotationProperty: skos:relatedMatch\nObjectProperty: hasPatientData\n"
                   "Individual: admission:BM_DEP\nIndividual: admission:ecog_stage\nIndividual: obo:DOID_4961\n")


def target_state(upto):
    """The fragment as shown after the first `upto` reveal steps. Commas are fixed so that every
    intermediate state is itself valid Manchester Syntax."""
    shown = [code for code, _, g in TARGET_LINES if g in REVEAL_ORDER[:upto]]
    out = []
    for k, line in enumerate(shown):
        nxt = shown[k + 1] if k + 1 < len(shown) else ""
        continues = nxt.startswith("    ") and not nxt.lstrip().startswith(("Class:", "DataProperty:"))
        out.append(line.rstrip(",") + ("," if line.startswith("    ") and continues else ""))
    return "\n".join(out)

DELTA = """ObjectProperty: hasPatientData
  Domain: Admission  Range: PatientData
DataProperty: BM_DEP
  Annotations: rdfs:label "Bone Marrow Depression"@en
DataProperty: ecog_stage
  Annotations:
    rdfs:comment "ECOG 0 = fully active,
      1 = restricted, 2 = ambulatory,
      3 = limited self-care,
      4 = completely disabled, 5 = dead"@en,
    prov:wasDerivedFrom src:ecog_stage"""

# Future work (slide 22): instance pattern + drug label -> a reviewable candidate, no causal axiom.
FUTURE = """Individual: cisplatin_bmdep
  Types: CandidateAssociation
  Facts:
    hasDrug Cisplatin,
    hasCondition BoneMarrowDepression,
    status "needs expert review"
  Annotations:
    prov:wasDerivedFrom src:BM_DEP,
    prov:wasDerivedFrom doc:cisplatin_label"""

FUTURE_SUPPORT = """Class: CandidateAssociation
ObjectProperty: hasDrug
ObjectProperty: hasCondition
DataProperty: status
Individual: Cisplatin
Individual: BoneMarrowDepression
Individual: src:BM_DEP
Individual: doc:cisplatin_label
"""

# Slide 10: errors a capable LLM still makes, which the paper's checks target.
# Semantic: a duplicate concept (Judge criterion 9, core consistency); valid OWL on its own.
DUP_WRONG = "ObjectProperty: hasPatient  Range: Patient"
DUP_RIGHT = "ObjectProperty: hasPatientData  Range: PatientData"
# Structural: a reused property gets a second domain (OOPS! P19; graph check: one domain per property).
CORE_LINK = "ObjectProperty: hasPatientData\n  Domain: Admission  Range: PatientData"
DOMAIN_WRONG = "ObjectProperty: hasPatientData  Domain: LabResult"
DOMAIN_RIGHT = "ObjectProperty: labResultOf  Domain: LabResult"


def code_only(text):
    return "\n".join(line.split("\t")[0] for line in text.splitlines())


def annotated(subject, ann):
    return f"DataProperty: {subject}\n  Annotations: {ann}"


# name -> (file body, expect OWL 2 DL)
CHECKS = {
    "slide03_target": (TARGET_HEADER + TARGET, True),
    **{f"slide03_reveal_step{k}": (TARGET_HEADER + target_state(k), True) for k in (1, 2, 3)},
    "slide07_seed": (HEADER + "ObjectProperty: patient_id\nDataProperty: BM_DEP\nDataProperty: ecog_stage\n"
                     + code_only(SEED), True),
    "slide09_before_seed": (HEADER + "ObjectProperty: patient_id\nDataProperty: BM_DEP\nDataProperty: ecog_stage\n"
                            + SEED_COMPACT, True),
    "slide09_after_delta": (HEADER + "Class: Admission\n" + DELTA, True),
    "slide10_duplicate_concept": (HEADER + "Class: Patient\n" + DUP_WRONG, True),
    "slide10_second_domain_merged": (HEADER + "Class: Admission\nClass: LabResult\n" + CORE_LINK + "\n"
                                     + DOMAIN_WRONG, True, {"hasPatientData": 2}),
    "slide10_corrected_merged": (HEADER + "Class: Admission\nClass: LabResult\n" + CORE_LINK + "\n"
                                 + DOMAIN_RIGHT + "  Range: PatientData", True, {"hasPatientData": 1, "labResultOf": 1}),
    "slide22_future_candidate": (HEADER.replace("Ontology:", "Prefix: doc: <http://example.org/docs/>\nOntology:")
                                 + FUTURE_SUPPORT + FUTURE, True),
}


def verify(classpath):
    EXCERPT_DIR.mkdir(exist_ok=True)
    ok = True
    for old in EXCERPT_DIR.glob("slide10_*.omn"):
        old.unlink()
    for name, (body, expect_dl, *domains) in CHECKS.items():
        path = EXCERPT_DIR / f"{name}.omn"
        path.write_text(body + "\n", encoding="utf-8")
        out = subprocess.run(["java", "-cp", f"{classpath}:{HERE}", "OwlCheck", str(path), "axioms"],
                             capture_output=True, text=True).stdout
        parsed = "PARSED OK" in out
        is_dl = "OWL 2 DL: yes" in out
        good = parsed and is_dl == expect_dl
        for prop, n in (domains[0] if domains else {}).items():  # expected number of domain axioms
            found = len(re.findall(rf"ObjectPropertyDomain\(<http://example.org/rigor#{prop}>", out))
            good &= found == n
            print(f"      {prop}: {found} domain axiom(s), expected {n}")
        ok &= good
        print(f"{'PASS' if good else 'FAIL'}  {name}: parsed={parsed} dl={is_dl} (expected dl={expect_dl})")
    sys.exit(0 if ok else 1)


# ---- Syntax colouring -------------------------------------------------------------------

KW, NAME, LIT, XSD, ANN, IRI, GLOSS = "0B6E4F", "1F2328", "A3401E", "6F42C1", "5B6770", "2F6B78", "5B6770"
KEYWORDS = {"Class:", "ObjectProperty:", "DataProperty:", "SubClassOf:", "Domain:", "Range:",
            "Annotations:", "some", "only", "Individual:", "Types:", "Facts:", "DisjointWith:"}
TOKEN = re.compile(r'"[^"]*"(?:@\w+)?|"[^"]*$|<[^>]*>|\s+|[\w:]+|.')


def tokens(line, in_string):
    """Yield (text, colour, bold) for one code line; track literals spanning lines."""
    out = []
    if in_string:
        end = line.find('"')
        if end < 0:
            return [(line, LIT, False)], True
        m = re.match(r'[^"]*"(?:@\w+)?', line)
        out.append((m.group(0), LIT, False))
        line, in_string = line[m.end():], False
    for tok in TOKEN.findall(line):
        if tok.startswith('"'):
            out.append((tok, LIT, False))
            if not re.fullmatch(r'"[^"]*"(?:@\w+)?', tok):
                in_string = True
        elif tok.startswith("<"):
            out.append((tok, IRI, False))
        elif tok in KEYWORDS:
            out.append((tok, KW, True))
        elif tok.startswith("xsd:"):
            out.append((tok, XSD, False))
        elif tok.startswith(("rdfs:", "prov:", "skos:")):
            out.append((tok, ANN, False))
        elif tok.startswith(("src:", "doc:", "obo:", "admission:")):
            out.append((tok, IRI, False))
        else:
            out.append((tok, NAME, False))
    return out, in_string


def run_xml(text, color, bold, size, font="Consolas", italic=False):
    return (f'<a:r><a:rPr lang="en-US" sz="{size}" b="{int(bold)}" i="{int(italic)}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:rPr>'
            f'<a:t xml:space="preserve">{escape(text)}</a:t></a:r>')


def code_paragraphs(text, size, line_pts, tab_in=None, gloss_size=1200, color_override=None):
    paras, in_string = [], False
    tab = f'<a:tabLst><a:tab pos="{int(tab_in * EMU)}" algn="l"/></a:tabLst>' if tab_in else ""
    for line in text.splitlines():
        code, _, gloss = line.partition("\t")
        toks, in_string = tokens(code, in_string)
        runs = "".join(run_xml(t, color_override or c, b, size) for t, c, b in toks)
        if gloss:
            runs += run_xml("\t" + gloss, GLOSS, False, gloss_size, font="Calibri", italic=True)
        paras.append(f'<a:p><a:pPr><a:lnSpc><a:spcPts val="{line_pts}"/></a:lnSpc>'
                     f'<a:spcBef><a:spcPts val="0"/></a:spcBef>{tab}</a:pPr>{runs}'
                     f'<a:endParaRPr lang="en-US" sz="{size}"/></a:p>')
    return "".join(paras)


def code_block(shape_id, x, y, w, h, text, size, line_pts, tab_in=None, inset=0.12):
    ins = int(inset * EMU)
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{shape_id}" name="Code {shape_id}" '
        f'descr="OWL Manchester Syntax excerpt"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/><a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 5000"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="F6F8FA"/></a:solidFill>'
        f'<a:ln w="9525"><a:solidFill><a:srgbClr val="D0D7DE"/></a:solidFill></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="none" lIns="{ins}" tIns="{int(0.08 * EMU)}" rIns="{ins}" bIns="{int(0.06 * EMU)}" '
        f'anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
        f'{code_paragraphs(text, size, line_pts, tab_in)}</p:txBody></p:sp>')


# ---- Slide edits ---------------------------------------------------------------------------

def sp_by_id(tree, sid):
    return next(sp for sp in tree.iter(f"{{{P}}}sp", f"{{{P}}}cxnSp")
                if sp.find(".//p:cNvPr", NS).get("id") == str(sid))


def delete(tree, *ids):
    for sid in ids:
        sp = sp_by_id(tree, sid)
        sp.getparent().remove(sp)


def move(tree, sid, x=None, y=None, w=None, h=None):
    sp = sp_by_id(tree, sid)
    off, ext = sp.find(".//a:off", NS), sp.find(".//a:ext", NS)
    for el, attr, val in ((off, "x", x), (off, "y", y), (ext, "cx", w), (ext, "cy", h)):
        if val is not None:
            el.set(attr, str(int(val * EMU)))


def set_text(tree, sid, text):
    sp = sp_by_id(tree, sid)
    ts = list(sp.iter(f"{{{A}}}t"))
    ts[0].text = text
    for t in ts[1:]:
        t.getparent().getparent().remove(t.getparent())


def replace_runs(tree, sid, runs_xml):
    """Replace all runs of a single-paragraph text shape, keeping its paragraph properties."""
    p = sp_by_id(tree, sid).find(".//a:p", NS)
    for child in list(p):
        if etree.QName(child).localname in ("r", "br"):
            p.remove(child)
    frag = etree.fromstring(f'<x xmlns:a="{A}">{runs_xml}</x>')
    anchor = p.find("a:endParaRPr", NS)
    for r in frag:
        if anchor is not None:
            anchor.addprevious(r)
        else:
            p.append(r)


def add(tree, el):
    tree.find(".//p:spTree", NS).append(el)


def next_id(tree):
    return max(int(e.get("id")) for e in tree.iter(f"{{{P}}}cNvPr")) + 1


def code_runs(text, size, color=None):
    toks, _ = tokens(text, False)
    return "".join(run_xml(t, color or c, b, size) for t, c, b in toks)


GLOSS_COLORS = {"schema": "5B6770", "docs": "2F6B78", "vocab": "0B6E4F", "llm": "6F42C1"}


def code_line(sid, x, y, w, h, code, gloss, group, size=1150, gloss_size=1100, tab_in=4.62):
    toks, _ = tokens(code, False)
    runs = "".join(run_xml(t, c, b, size) for t, c, b in toks)
    if gloss:
        runs += run_xml("\t\u2190 " + gloss, GLOSS_COLORS[group], False, gloss_size, font="Calibri", italic=True)
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{sid}" name="Code line {sid}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
        f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="{int(0.16 * EMU)}" tIns="0" rIns="0" bIns="0" '
        f'anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:tabLst><a:tab pos="{int(tab_in * EMU)}" '
        f'algn="l"/></a:tabLst></a:pPr>{runs}<a:endParaRPr lang="en-US" sz="{size}"/></a:p></p:txBody></p:sp>')


def simple_box(sid, name, x, y, w, h, fill=None, line=None, paras="", geom="rect", shadow=False, inset=0.0):
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    line_xml = (f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
                else "<a:ln><a:noFill/></a:ln>")
    shadow_xml = ('<a:effectLst><a:outerShdw blurRad="76200" dist="25400" dir="5400000" algn="t" rotWithShape="0">'
                  '<a:srgbClr val="000000"><a:alpha val="16000"/></a:srgbClr></a:outerShdw></a:effectLst>'
                  if shadow else "")
    adj = '<a:avLst><a:gd name="adj" fmla="val 9000"/></a:avLst>' if geom == "roundRect" else "<a:avLst/>"
    ins = int(inset * EMU)
    body = paras or '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/>'
        f'</p:nvSpPr><p:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/><a:ext cx="{int(w * EMU)}" '
        f'cy="{int(h * EMU)}"/></a:xfrm><a:prstGeom prst="{geom}">{adj}</a:prstGeom>{fill_xml}{line_xml}{shadow_xml}'
        f'</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="{ins}" tIns="0" rIns="{ins}" bIns="0" anchor="ctr">'
        f'<a:noAutofit/></a:bodyPr><a:lstStyle/>{body}</p:txBody></p:sp>')


def para(runs, algn="l", space_after=0):
    return (f'<a:p><a:pPr algn="{algn}"><a:spcAft><a:spcPts val="{space_after}"/></a:spcAft></a:pPr>{runs}'
            f'</a:p>')


def reveal_timing(steps, text_ids):
    """PowerPoint build: each step is a list of shape ids that fade in together on one click."""
    n = [2]

    def nid():
        n[0] += 1
        return n[0]

    clicks = []
    for step in steps:
        effects = []
        for k, spid in enumerate(step):
            node = "clickEffect" if k == 0 else "withEffect"
            effects.append(
                f'<p:par><p:cTn id="{nid()}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" '
                f'grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                f'</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst>'
                f'<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/>'
                f'</p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="250"/>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
        outer, inner = nid(), nid()
        clicks.append(f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
                      f'<p:childTnLst><p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                      f'</p:stCondLst><p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par></p:childTnLst>'
                      f'</p:cTn></p:par>')
    blds = "".join(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>' for step in steps for spid in step
                   if spid in text_ids)
    return etree.fromstring(
        f'<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst>{"".join(clicks)}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0">'
        f'<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0">'
        f'<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'<p:bldLst>{blds}</p:bldLst></p:timing>')


def edit_slide3(tree):
    """Full target example with a reveal: input -> schema -> documentation -> vocabulary -> LLM -> uses."""
    sid = [next_id(tree) - 1]

    def new_id():
        sid[0] += 1
        return sid[0]

    # Left column: input card (labelled INPUT, table as a card with green column names).
    move(tree, 77, y=2.25, h=2.77)
    set_text(tree, 78, "INPUT")
    move(tree, 78, y=2.42)
    add(tree, simple_box(new_id(), "Table card", 1.06, 2.86, 2.89, 1.96, fill="E9EEEB", geom="roundRect",
                         shadow=True))
    for mid in (79, 80):  # bring the table name and its rule in front of the new card
        el = sp_by_id(tree, mid)
        el.getparent().append(el)
    move(tree, 79, x=1.24, y=2.98, w=2.55, h=0.40)
    move(tree, 80, x=1.24, y=3.44, w=2.53)
    delete(tree, 81, 82, 83, 84)
    for k, (name, typ) in enumerate((("patient_id", "FK"), ("BM_DEP", "INT"), ("ecog_stage", "INT"))):
        y = 3.58 + k * 0.38
        add(tree, simple_box(new_id(), "Column", 1.24, y, 1.75, 0.32,
                             paras=para(run_xml(name, "0B6E4F", True, 1500, font="Calibri"))))
        add(tree, simple_box(new_id(), "Type", 2.95, y, 0.82, 0.32,
                             paras=para(run_xml(typ, "5B6770", False, 1400, font="Calibri"), algn="r")))
    move(tree, 85, y=3.60)

    # Target panel: taller, compact class diagram.
    move(tree, 86, y=2.25, h=4.77)
    set_text(tree, 87, "TARGET")
    move(tree, 87, y=2.42)
    for box, text in ((88, 89), (90, 91)):
        move(tree, box, y=2.88, h=0.52)
        move(tree, text, y=3.00)
    move(tree, 92, y=3.14)
    move(tree, 93, y=2.72, h=0.30)
    move(tree, 94, y=2.75)

    # Code panel and one text box per line, so the reveal can add them by knowledge source.
    delete(tree, 95, 96, 97, 98, 99, 100, 101)
    px, py, pw, ph = 5.45, 3.62, 6.80, 3.24
    add(tree, simple_box(new_id(), "Code panel", px, py, pw, ph, fill="F6F8FA", line="D0D7DE", geom="roundRect"))
    gap, pad = 0.10, 0.13
    pitch = (ph - 2 * pad - gap * len(TARGET_GROUP_BREAKS)) / len(TARGET_LINES)
    y, groups = py + pad, {g: [] for g in REVEAL_ORDER}
    for k, (code, gloss, group) in enumerate(TARGET_LINES):
        lid = new_id()
        add(tree, code_line(lid, px, y, pw, pitch, code, gloss, group))
        groups[group].append(lid)
        y += pitch + (gap if k in TARGET_GROUP_BREAKS else 0)

    # What this enables: its own card under the input.
    delete(tree, 103)
    move(tree, 102, x=0.88, y=5.26, w=3.25, h=1.76)
    enables = new_id()
    items = "".join(para(run_xml(t, "1F2328", False, 1400, font="Calibri"), space_after=300)
                    for t in ("Concept-based access", "Knowledge graph population", "Cross-system interoperability"))
    add(tree, simple_box(enables, "What this enables", 1.10, 5.38, 2.85, 1.38,
                         paras=para(run_xml("WHAT THIS ENABLES", "0B6E4F", True, 1150), space_after=500) + items))
    sp_by_id(tree, enables).find(".//a:bodyPr", NS).set("anchor", "t")

    # Reveal: diagram + schema lines, then documentation, vocabulary, LLM knowledge, then uses.
    steps = [[88, 89, 90, 91, 92, 93, 94] + groups["schema"], groups["docs"], groups["vocab"], groups["llm"],
             [102, enables]]
    text_ids = {int(e.get("id")) for sp in tree.iter(f"{{{P}}}sp") if sp.find("p:txBody", NS) is not None
                for e in [sp.find(".//p:cNvPr", NS)]}
    sld = tree.getroot()
    old = sld.find("p:timing", NS)
    if old is not None:
        sld.remove(old)
    sld.find("p:clrMapOvr", NS).addnext(reveal_timing(steps, text_ids))


def edit_slide7(tree):
    delete(tree, 230, 231, 232, 233, 234, 235, 236, 237)
    add(tree, code_block(next_id(tree), 6.27, 2.95, 5.76, 2.14, SEED, 1400, 1600, tab_in=3.55))


def edit_slide9(tree):
    # Left: the seed (same text as slide 7, compact).
    delete(tree, 336, 337, 338, 339)
    add(tree, code_block(next_id(tree), 1.04, 3.30, 2.92, 1.20, SEED_COMPACT, 1100, 1350, inset=0.08))
    move(tree, 340, y=4.72)
    # Right: the enriched delta as code; evidence tags re-aligned to the lines they support.
    delete(tree, 365, 366, 367, 368, 369, 370, 371, 373, 374, 375, 377, 378, 379, 381, 382, 383)
    top, line_in, pad = 2.98, 1450 / 100 / 72, 0.08
    add(tree, code_block(next_id(tree), 5.40, top, 6.75, 11 * line_in + 0.22, DELTA, 1300, 1450))
    for tag, line in ((372, 0), (376, 3), (380, 6), (384, 10)):
        centre = top + pad + (line + 0.5) * line_in
        move(tree, tag, x=11.05, y=centre - 0.11, w=1.0, h=0.22)
        sp = sp_by_id(tree, tag)  # bring the tag in front of the code block
        sp.getparent().append(sp)


def edit_slide10(tree):
    set_text(tree, 340, "Why validate? Plausible fragments can still break the ontology.")
    set_text(tree, 341, "A fluent LLM rarely gets syntax wrong. The risk is fragments that clash with the "
                        "ontology built so far.")
    set_text(tree, 379, "Generated link for admission:")
    move(tree, 381, x=1.25, w=4.85)
    replace_runs(tree, 381, code_runs(DUP_WRONG, 1200, color="B23A1E"))
    set_text(tree, 382, "Valid OWL, but the core already has PatientData: now two classes mean the same "
                        "thing. The Judge-LLM sees the core and asks to reuse it.")
    set_text(tree, 384, "Later table lab_result reuses a core property:")
    replace_runs(tree, 386, code_runs(DOMAIN_WRONG, 1200, color="B23A1E"))
    set_text(tree, 387, "Merged with Domain: Admission, OWL infers that every lab result is an admission "
                        "(pitfall P19). Graph checks allow one domain per property.")
    replace_runs(tree, 389, run_xml("CORRECTION   ", "0B6E4F", True, 1300, font="Calibri")
                 + code_runs(DUP_RIGHT, 1200)
                 + run_xml("    |    ", "5B6770", False, 1300, font="Calibri")
                 + code_runs(DOMAIN_RIGHT, 1200))


EDITS = {3: edit_slide3, 7: edit_slide7, 9: edit_slide9, 10: edit_slide10}

# Speaker-note sentences that described the old pseudo-notation panels.
NOTE_EDITS = {
    9: [("The panel now uses the same format in every row: ontology term, generated statement, "
         "and the evidence informing it.",
         "Read the code top to bottom: each block is one ontology element, and the tag on the right "
         "shows the evidence that informed it."),
        ("Each row is an ontology contribution;", "Each code block is an ontology contribution;")],
}


def apply(src, dst):
    work = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
    for n, fn in EDITS.items():
        path = work / f"ppt/slides/slide{n}.xml"
        tree = etree.parse(str(path))
        fn(tree)
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    for n, pairs in NOTE_EDITS.items():
        path = work / f"ppt/notesSlides/notesSlide{n}.xml"
        xml = path.read_text(encoding="utf-8")
        for old, new in pairs:
            assert old in xml, old
            xml = xml.replace(old, escape(new))
        path.write_text(xml, encoding="utf-8")
    Path(dst).unlink(missing_ok=True)
    files = sorted(f for f in work.rglob("*") if f.is_file())
    files.sort(key=lambda f: f.name != "[Content_Types].xml")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.relative_to(work).as_posix())
    shutil.rmtree(work)
    print("wrote", dst)


if __name__ == "__main__":
    {"verify": lambda: verify(sys.argv[2]), "apply": lambda: apply(sys.argv[2], sys.argv[3])}[sys.argv[1]]()
