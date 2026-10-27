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

TARGET = """DataProperty: BM_DEP
  Annotations: rdfs:label
    "Bone Marrow Depression"@en\t← human-readable name
DataProperty: ecog_stage
  Annotations: rdfs:comment
    "ECOG 0 = fully active … 5 = dead"@en,\t← what the codes mean
    prov:wasDerivedFrom src:ecog_stage\t← where it came from"""

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
    "slide03_target": (HEADER + code_only(TARGET), True),
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
            "Annotations:", "some", "only", "Individual:", "Types:", "Facts:"}
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
        elif tok.startswith(("rdfs:", "prov:")):
            out.append((tok, ANN, False))
        elif tok.startswith(("src:", "doc:")):
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
    return next(sp for sp in tree.iter(f"{{{P}}}sp") if sp.find(".//p:cNvPr", NS).get("id") == str(sid))


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


def edit_slide3(tree):
    # The link Admission -hasPatientData-> PatientData stays as the diagram above the code.
    delete(tree, 95, 96, 97, 98, 99, 100, 101)
    add(tree, code_block(next_id(tree), 5.57, 3.86, 6.55, 7 * 1500 / 100 / 72 + 0.18, TARGET, 1300, 1500, tab_in=4.35))


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
