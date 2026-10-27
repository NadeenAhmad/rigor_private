#!/usr/bin/env python3
"""QA: WCAG contrast of every text run against the fill it sits on, plus small-text inventory.

The background of a run is its own shape's fill, else the last earlier filled shape that
contains the shape's centre, else the slide background.

Usage: python3 qa_contrast.py deck.pptx [min_ratio]
"""
import collections
import re
import sys
import zipfile

from lxml import etree

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def luminance(hex_color):
    c = [int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def ratio(a, b):
    hi, lo = sorted((luminance(a), luminance(b)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


def box(sp):
    off, ext = sp.find(".//" + A + "off"), sp.find(".//" + A + "ext")
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")), int(off.get("y"))
    return x, y, x + int(ext.get("cx")), y + int(ext.get("cy"))


def main(deck, min_ratio="4.5"):
    min_ratio = float(min_ratio)
    worst = collections.defaultdict(lambda: (99.0, None))
    sizes = collections.Counter()
    with zipfile.ZipFile(deck) as z:
        names = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                       key=lambda n: int(re.findall(r"\d+", n)[0]))
        for name in names:
            n = int(re.findall(r"\d+", name)[0])
            tree = etree.fromstring(z.read(name))
            bg = tree.find(".//" + P + "bg//" + A + "srgbClr")
            bg = bg.get("val") if bg is not None else "FFFFFF"
            shapes = list(tree.iter(P + "sp"))
            filled = [(box(s), s.find(P + "spPr/" + A + "solidFill/" + A + "srgbClr").get("val")) for s in shapes
                      if s.find(P + "spPr/" + A + "solidFill/" + A + "srgbClr") is not None and box(s)]
            for sp in shapes:
                b = box(sp)
                if b is None:
                    continue
                cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
                under = bg
                for fb, fc in filled:
                    if fb != b and fb[0] <= cx <= fb[2] and fb[1] <= cy <= fb[3]:
                        under = fc
                own = sp.find(P + "spPr/" + A + "solidFill/" + A + "srgbClr")
                if own is not None:
                    under = own.get("val")
                for r in sp.iter(A + "r"):
                    text = (r.findtext(A + "t") or "").strip()
                    color = r.find(A + "rPr/" + A + "solidFill/" + A + "srgbClr")
                    if not text or color is None:
                        continue
                    size = int(r.find(A + "rPr").get("sz", "1800")) / 100
                    sizes[(n, size)] += 1
                    key = (color.get("val"), under)
                    value = ratio(*key)
                    if value < worst[key][0]:
                        worst[key] = (value, (n, size, text[:40]))
    failures = [(v, k, ex) for k, (v, ex) in worst.items() if v < min_ratio]
    for v, (c, u), ex in sorted(failures):
        print(f"FAIL {v:4.2f}  text {c} on {u}  e.g. slide {ex[0]}, {ex[1]:.0f} pt: \"{ex[2]}\"")
    print(f"contrast: {len(worst)} color pairs, {len(failures)} below {min_ratio}:1")
    small = collections.Counter(n for (n, s), c in sizes.items() for _ in range(c) if s < 12)
    print("runs under 12 pt per slide:", dict(sorted(small.items())))


if __name__ == "__main__":
    main(*sys.argv[1:])
