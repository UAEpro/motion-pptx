#!/usr/bin/env python3
"""Check an animated .pptx for the rules PowerPoint enforces beyond the XSD.

    python check_deck.py deck.pptx [--xsd path/to/pml.xsd]

Exit 1 on any problem. Checks, per slide:
  - every image relationship resolves to a file in the package
  - cNvPr ids are unique; every animation spid exists
  - cTn ids are 1..N in document order
  - bldP: none for pictures, no animBg on text boxes
  - repeatCount only on effect-level cTn; motion paths have rAng
  - at most one <a:pPr> per paragraph
  - (optional) slide XML validates against PresentationML XSD
"""
import argparse
import collections
import posixpath
import re
import sys
import zipfile


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--xsd", help="pml.xsd (e.g. the pptx skill's scripts/office/schemas/ISO-IEC29500-4_2016/pml.xsd)")
    args = ap.parse_args()

    schema = None
    if args.xsd:
        from lxml import etree
        schema = etree.XMLSchema(etree.parse(args.xsd))

    z = zipfile.ZipFile(args.pptx)
    names = set(z.namelist())
    slides = sorted((n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)), key=lambda n: int(re.findall(r"\d+", n)[0]))
    problems = []
    for sl in slides:
        x = z.read(sl).decode("utf-8")
        tag = sl.rsplit("/", 1)[1]
        rels_name = sl.replace("slides/", "slides/_rels/") + ".rels"
        rels = z.read(rels_name).decode() if rels_name in names else ""
        for target in re.findall(r'Target="([^"]+)"', rels):
            if target.startswith("http"):
                continue
            if posixpath.normpath(posixpath.join("ppt/slides", target)) not in names:
                problems.append(f"{tag}: missing part {target}")
        shapes = re.findall(r'<p:(sp|pic|cxnSp)>\s*<p:nv(?:Sp|Pic|CxnSp)Pr>\s*<p:cNvPr id="(\d+)"', x)
        ids = [i for _, i in shapes]
        dup = [k for k, v in collections.Counter(ids).items() if v > 1]
        if dup:
            problems.append(f"{tag}: duplicate shape ids {dup}")
        kind = {i: k for k, i in shapes}
        txbox = set(re.findall(r'<p:cNvPr id="(\d+)"[^>]*>(?:</p:cNvPr>)?<p:cNvSpPr txBox="1"', x))
        for spid in set(re.findall(r'spid="(\d+)"', x)):
            if spid not in kind:
                problems.append(f"{tag}: animation targets unknown shape {spid}")
        ctn = [int(v) for v in re.findall(r'<p:cTn id="(\d+)"', x)]
        if ctn and ctn != list(range(1, len(ctn) + 1)):
            problems.append(f"{tag}: cTn ids not sequential")
        for spid, rest in re.findall(r'<p:bldP spid="(\d+)"([^/]*)/>', x):
            if kind.get(spid) == "pic":
                problems.append(f"{tag}: bldP on picture {spid}")
            if spid in txbox and "animBg" in rest:
                problems.append(f"{tag}: animBg on text box {spid}")
        for m in re.finditer(r'<p:cBhvr[^>]*>\s*<p:cTn [^>]*repeatCount', x):
            problems.append(f"{tag}: repeatCount on a behaviour cTn (put it on the effect)")
        if re.search(r"<p:animMotion(?![^>]*rAng)", x):
            problems.append(f"{tag}: animMotion without rAng")
        for p in re.findall(r"<a:p>(.*?)</a:p>", x, flags=re.S):
            if p.count("<a:pPr") > 1:
                problems.append(f"{tag}: paragraph with more than one a:pPr")
                break
        if schema is not None:
            from lxml import etree
            if not schema.validate(etree.fromstring(x.encode())):
                problems.append(f"{tag}: XSD: {schema.error_log.last_error}")
        n_fx = x.count("presetClass=")
        print(f"{tag}: {len(ids)} shapes, {n_fx} effects")
    if problems:
        print("\nPROBLEMS:")
        for p in problems:
            print(" -", p)
        return 1
    print("\nOK: no PowerPoint animation-rule problems found")
    return 0


if __name__ == "__main__":
    sys.exit(main())
