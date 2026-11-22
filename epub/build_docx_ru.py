#!/usr/bin/env python3
"""Build the Russian DOCX manuscript of «Появление» for Amazon KDP (from book_ru.md)."""
import re, subprocess, pathlib
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
here = pathlib.Path(__file__).resolve().parent
out = here / "Появление.docx"
src = (here / "book_ru.md").read_text(encoding="utf-8")
# drop the HTML-oriented title block and rebuild front matter with Word styles
body = src[src.index("# Авторские права"):]
body = re.sub(r"^---\s*$", '\n::: {custom-style="Scene Break"}\n\\\\* \\\\* \\\\*\n:::\n', body, flags=re.M)
body = body.replace("# Авторские права", "# Авторские права")
# stable ids for chapter headings + a linked table of contents
heads = []
def tag(m):
    heads.append(m.group(1)); return f"## {m.group(1)} {{#ch{len(heads):02d}}}"
body = re.sub(r"^## (.+)$", tag, body, flags=re.M)
toc = '# Содержание\n\n::: {custom-style="TOC Entry"}\n' + "\n\n".join(
    f"[{h}](#ch{i:02d})" for i, h in enumerate(heads, 1)) + "\n:::\n\n"
body = body.replace("# Часть первая. Появление", toc + "# Часть первая. Появление", 1)
md = ('::: {custom-style="Book Title"}\nПоявление\n:::\n\n'
      '::: {custom-style="Book Subtitle"}\nМистический приключенческий роман\n:::\n\n'
      '::: {custom-style="Book Author"}\nАлександр Лунин\n:::\n\n'
      '::: {custom-style="Book Subtitle"}\nСерия «Terra Incognita». Книга первая\n:::\n\n') + body
tmp = here / "_docx_en.md"
tmp.write_text(md, encoding="utf-8")
subprocess.run(["pandoc", str(tmp), "-f", "markdown+smart", "-t", "docx", "-M", "lang=ru-RU", "-o", str(out)], check=True)
tmp.unlink()

d = Document(out)
sec = d.sections[0]
sec.page_width, sec.page_height = Inches(6), Inches(9)
sec.left_margin = sec.right_margin = Inches(0.8)
sec.top_margin = sec.bottom_margin = Inches(0.8)
def style(name, **kw):
    try: s = d.styles[name]
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        s = d.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    s.font.name = "Times New Roman"
    s.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    s.font.color.rgb = None
    f = s.paragraph_format
    s.font.size = Pt(kw.get("size", 12)); s.font.bold = kw.get("bold", False); s.font.italic = False
    f.alignment = kw.get("align", WD_ALIGN_PARAGRAPH.JUSTIFY)
    f.first_line_indent = kw.get("indent"); f.left_indent = None
    f.space_before = Pt(kw.get("before", 0)); f.space_after = Pt(kw.get("after", 0))
    f.line_spacing = kw.get("ls", 1.15)
    f.page_break_before = kw.get("pb", False); f.keep_with_next = kw.get("kwn", False)
    return s
C = WD_ALIGN_PARAGRAPH.CENTER; L = WD_ALIGN_PARAGRAPH.LEFT
for n in ("Normal", "Body Text", "First Paragraph", "Compact"):
    style(n, indent=Inches(0.3))
style("Heading 1", size=20, bold=True, align=C, before=72, after=24, pb=True, kwn=True, indent=None)
style("Heading 2", size=16, bold=True, align=C, before=72, after=24, pb=True, kwn=True, indent=None)
style("Book Title", size=32, bold=True, align=C, before=150, after=18, indent=None)
style("Book Subtitle", size=14, align=C, before=6, after=6, indent=None)
style("Book Author", size=18, align=C, before=36, after=36, indent=None)
style("Scene Break", align=C, before=12, after=12, indent=None)
style("TOC Entry", align=L, after=3, indent=None, ls=1.0)
for p in d.paragraphs:  # front matter lines: no indent
    pass
# copyright / about blocks read better unindented before Part One
inside = False
for p in d.paragraphs:
    if p.style.name == "Heading 1":
        inside = p.text in ("Авторские права",)
    elif inside:
        p.paragraph_format.first_line_indent = Inches(0); p.alignment = L; p.paragraph_format.space_after = Pt(6)
# hyperlink look in TOC: plain black, no underline
from docx.shared import RGBColor
for sname in ("Hyperlink",):
    try:
        hs = d.styles[sname]; hs.font.color.rgb = RGBColor(0,0,0); hs.font.underline = False
    except KeyError: pass
d.core_properties.title = "Появление"; d.core_properties.author = "Александр Лунин"
d.core_properties.language = "ru-RU"
d.save(out)
print("built", out)
