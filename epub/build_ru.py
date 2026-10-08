#!/usr/bin/env python3
"""Build the Russian EPUB of «Появление» (Part One) with pandoc, in TOC order."""
import re, subprocess, sys, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
here = pathlib.Path(__file__).resolve().parent
toc = next(root.glob("«ПОЯВЛЕНИЕ» — ОГЛАВЛЕНИЕ*.md")).read_text(encoding="utf-8").splitlines()
def norm(s):
    s = re.sub(r"[.,?«»\"]", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()
files = {norm(p.stem): p for p in root.glob("*.md") if p.name.startswith(("Глава","Явь","Видение"))}
order = []
for line in toc[98:170]:
    if not line.startswith("|") or "✅" not in line: continue
    title = line.split("|")[2].strip()
    title = re.split(r" — ", title)[0]
    title = re.sub(r"\s*\(.*?\)", "", title).strip()
    key = norm(title)
    if key not in files: sys.exit(f"no file for: {title!r} ({key})")
    order.append(files[key])
print(len(order), "files")
parts = [(here/"front_ru.md").read_text(encoding="utf-8"), "\n\n# Часть первая. Появление\n"]
for p in order:
    lines = p.read_text(encoding="utf-8").splitlines()
    label, ttl = lines[0].strip(), lines[1].strip()
    head = f"{label}. {ttl}" if not label.endswith(".") else f"{label} {ttl}"
    body = "\n".join(lines[2:]).strip()
    body = re.sub(r"^\s*\*\s\*\s\*\s*$", "\n---\n", body, flags=re.M)
    parts.append(f"\n\n## {head}\n\n{body}\n")
(here/"book_ru.md").write_text("".join(parts), encoding="utf-8")
meta = here/"meta_ru.yaml"
meta.write_text("""---
title: Появление
subtitle: Мистический приключенческий роман
author: Александр Лунин
lang: ru
rights: Copyright © 2020 - 2026 Alexander Lunin. Все права защищены.
publisher: Независимое издание
identifier:
  - scheme: ISBN
    text: 979-8176203813
belongs-to-collection: Terra Incognita
collection-type: series
group-position: 1
description: Человек приходит в себя на залитой солнцем улице восточного портового города, не помня, кто он. Мистический приключенческий роман, первая книга серии «Terra Incognita».
---
""", encoding="utf-8")
cmd = ["pandoc", str(meta), str(here/"book_ru.md"), "-f", "markdown+smart-implicit_figures", "-t", "epub3",
       "--toc", "--toc-depth=2", "--split-level=2", "--css", str(here/"style.css"), "-o", str(here/"poyavlenie-ru.epub")]
cover = next((c for c in here.glob("cover.*") if c.suffix.lower() in (".jpg",".jpeg",".png")), None)
if cover: cmd += ["--epub-cover-image", str(cover)]
subprocess.run(cmd, check=True)
print("built", here/"poyavlenie-ru.epub", "cover:", cover)
