#!/usr/bin/env python3
"""Build the English EPUB of «The Appearance» (Part One) with pandoc, in TOC order."""
import re, subprocess, sys, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
here = pathlib.Path(__file__).resolve().parent
toc = next(root.glob("«ПОЯВЛЕНИЕ» — ОГЛАВЛЕНИЕ*.md")).read_text(encoding="utf-8").splitlines()
def norm(s):
    s = re.sub(r"[.,?«»\"]", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()
files = {norm(p.stem): p for p in root.glob("*.md") if p.name.startswith(("Глава","Явь","Видение"))}
RN = {"Глава":"Chapter","Явь":"Reality","Видение":"Vision"}
def english(p):
    m = re.match(r"(Глава|Явь|Видение)\s+(\S+)", p.stem)
    kind, num = RN[m.group(1)], m.group(2).rstrip(".").rstrip(".")
    if num.lower().startswith("пролог"): num = "Prologue"
    hits = list((root/"English").glob(f"{kind} {num}.*.md"))
    assert len(hits) == 1, (p.name, hits)
    return hits[0]
order = []
for line in toc[98:170]:
    if not line.startswith("|") or "✅" not in line: continue
    title = line.split("|")[2].strip()
    title = re.split(r" — ", title)[0]
    title = re.sub(r"\s*\(.*?\)", "", title).strip()
    key = norm(title)
    if key not in files: sys.exit(f"no file for: {title!r} ({key})")
    order.append(english(files[key]))
print(len(order), "files")
front = (here/"front_en.md").read_text(encoding="utf-8")
front, about_author = front.split("# About the Author")
about_author = "# About the Author" + about_author
desc = re.search(r"# About the Book\n(.*?)(?=\n# |\Z)", front, re.S).group(1)
desc = re.sub(r"\s+", " ", desc).strip().replace('"', "'")
parts = [front, "\n\n# Part One. The Appearance\n"]
for p in order:
    lines = p.read_text(encoding="utf-8").splitlines()
    label, ttl = lines[0].strip(), lines[1].strip()
    head = f"{label}. {ttl}" if not label.endswith(".") and not ttl.startswith("Prologue") else f"{label} {ttl}"
    body = "\n".join(lines[2:]).strip()
    body = re.sub(r"^\s*\*\s\*\s\*\s*$", "\n---\n", body, flags=re.M)
    parts.append(f"\n\n## {head}\n\n{body}\n")
parts.append("\n\n" + about_author)
(here/"book_en.md").write_text("".join(parts), encoding="utf-8")
meta = here/"meta_en.yaml"
meta.write_text("""---
title: The Appearance
subtitle: A Mystical Adventure Novel
author: Alexander Lunin
lang: en
rights: Copyright © 2020 - 2026 Alexander Lunin. All rights reserved.
publisher: Independently published
identifier:
  - scheme: ISBN
    text: 979-8176203813
belongs-to-collection: Terra Incognita
collection-type: series
group-position: 1
description: "%s"
---
""" % desc, encoding="utf-8")
cmd = ["pandoc", str(meta), str(here/"book_en.md"), "-f", "markdown+smart-implicit_figures", "-t", "epub3",
       "--toc", "--toc-depth=2", "--split-level=2", "--css", str(here/"style.css"), "-o", str(here/"the-appearance-en.epub")]
cover = next((c for c in here.glob("cover.*") if c.suffix.lower() in (".jpg",".jpeg",".png")), None)
if cover: cmd += ["--epub-cover-image", str(cover)]
subprocess.run(cmd, check=True)
print("built", here/"the-appearance-en.epub", "cover:", cover)
