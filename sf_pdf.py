#!/usr/bin/env python3
"""Turn the executive summary Markdown into a clean, client-ready PDF.
Supports: # / ## / ### headings, paragraphs, - bullets, 1. numbered lists,
pipe tables, **bold**, *italic*, `code`, [links](url). A line right under the
title that starts with "_" is shown as the subtitle (e.g. _Crawl date: ..._).
Usage: python3 -I sf_pdf.py <summary.md> <output.pdf>
"""
import sys, re, html
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                ListFlowable, ListItem)

NAVY, GREY, LIGHT = colors.HexColor("#1F3864"), colors.HexColor("#555555"), colors.HexColor("#EEF2F8")
PRIORITY_FILL = {"critical": "#F8CBAD", "important": "#FFE699", "minor": "#E2EFDA"}

ss = getSampleStyleSheet()
S = {
    "h1": ParagraphStyle("h1", parent=ss["Title"], fontSize=20, leading=24, textColor=NAVY, alignment=0, spaceAfter=4),
    "sub": ParagraphStyle("sub", parent=ss["Normal"], fontSize=9.5, textColor=GREY, spaceAfter=10),
    "h2": ParagraphStyle("h2", parent=ss["Heading2"], fontSize=14, leading=18, textColor=NAVY, spaceBefore=12, spaceAfter=6),
    "h3": ParagraphStyle("h3", parent=ss["Heading3"], fontSize=11.5, leading=15, textColor=NAVY, spaceBefore=8, spaceAfter=4),
    "p": ParagraphStyle("p", parent=ss["BodyText"], fontSize=10, leading=14, spaceAfter=6),
    "cell": ParagraphStyle("cell", parent=ss["BodyText"], fontSize=8.5, leading=11),
    "head": ParagraphStyle("head", parent=ss["BodyText"], fontSize=8.5, leading=11, textColor=colors.white, fontName="Helvetica-Bold"),
}

def inline(t):
    t = html.escape(t, quote=False)
    t = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<link href="\2" color="#0563C1"><u>\1</u></link>', t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<i>\1</i>", t)
    t = re.sub(r"(?<!\w)_(?!\s)(.+?)(?<!\s)_(?!\w)", r"<i>\1</i>", t)
    t = re.sub(r"`([^`]+)`", r'<font face="Courier">\1</font>', t)
    return t

def table(rows, width):
    cells = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
    cells = [r for r in cells if not all(re.fullmatch(r":?-{2,}:?", c) for c in r)]
    n = max(len(r) for r in cells)
    cells = [r + [""] * (n - len(r)) for r in cells]
    data = [[Paragraph(inline(c), S["head"]) for c in cells[0]]] + \
           [[Paragraph(inline(c), S["cell"]) for c in r] for r in cells[1:]]
    # every column fits its longest word; spare width goes to text-heavy columns
    def word_w(i, row):
        font = "Helvetica-Bold" if row == 0 else "Helvetica"
        return max([stringWidth(w, font, 8.5) for w in re.split(r"\s+", re.sub(r"[*`_]", "", cells[row][i])) if w] or [0])
    minw = [max(word_w(i, r) for r in range(len(cells))) + 16 for i in range(n)]
    if sum(minw) > width:
        minw = [width * m / sum(minw) for m in minw]
    lens = [max(len(r[i]) for r in cells) for i in range(n)]
    extra = [max(min(l, 80) - 12, 0) for l in lens]
    spare = width - sum(minw)
    cw = [m + (spare * e / sum(extra) if sum(extra) else spare / n) for m, e in zip(minw, extra)]
    st = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C9D3E3")),
          ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    for i, r in enumerate(cells[1:], 1):
        if i % 2 == 0:
            st.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
        fill = PRIORITY_FILL.get(r[0].strip("* ").lower())
        if fill:
            st.append(("BACKGROUND", (0, i), (0, i), colors.HexColor(fill)))
    t = Table(data, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle(st))
    return t

def build(md_path, pdf_path):
    lines = open(md_path, encoding="utf-8").read().splitlines()
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm,
                            topMargin=16*mm, bottomMargin=16*mm)
    W = doc.width
    out, para, i, title = [], [], 0, ""

    def flush():
        if para:
            out.append(Paragraph(inline(" ".join(para)), S["p"])); para.clear()

    while i < len(lines):
        ln = lines[i].rstrip()
        if not ln.strip():
            flush(); i += 1; continue
        if ln.startswith("# "):
            flush(); title = ln[2:].strip(); out.append(Paragraph(inline(title), S["h1"]))
            if i + 2 < len(lines) and lines[i+1].strip() == "" and lines[i+2].strip().startswith("_"):
                out.append(Paragraph(inline(lines[i+2].strip().strip("_")), S["sub"])); i += 3; continue
            if i + 1 < len(lines) and lines[i+1].strip().startswith("_"):
                out.append(Paragraph(inline(lines[i+1].strip().strip("_")), S["sub"])); i += 2; continue
            i += 1; continue
        if ln.startswith("## "):
            flush(); out.append(Paragraph(inline(ln[3:]), S["h2"])); i += 1; continue
        if ln.startswith("### "):
            flush(); out.append(Paragraph(inline(ln[4:]), S["h3"])); i += 1; continue
        if ln.lstrip().startswith("|"):
            flush(); rows = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                rows.append(lines[i]); i += 1
            out += [table(rows, W), Spacer(1, 8)]; continue
        m = re.match(r"\s*(-|\*|\d+\.)\s+(.*)", ln)
        if m:
            flush(); ordered = m.group(1)[0].isdigit(); items = []
            while i < len(lines):
                m = re.match(r"\s*(-|\*|\d+\.)\s+(.*)", lines[i])
                if m:
                    items.append(m.group(2)); i += 1
                elif lines[i].startswith("  ") and lines[i].strip() and items:
                    items[-1] += " " + lines[i].strip(); i += 1
                else:
                    break
            out.append(ListFlowable([ListItem(Paragraph(inline(t), S["p"]), leftIndent=12) for t in items],
                                    bulletType="1" if ordered else "bullet", start="1" if ordered else None,
                                    leftIndent=14, bulletFontSize=9))
            continue
        para.append(ln.strip()); i += 1
    flush()

    def footer(c, d):
        c.saveState(); c.setFont("Helvetica", 8); c.setFillColor(GREY)
        c.drawString(18*mm, 9*mm, title[:110]); c.drawRightString(A4[0] - 18*mm, 9*mm, f"Page {d.page}")
        c.restoreState()
    doc.title = title
    doc.build(out, onFirstPage=footer, onLaterPages=footer)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    build(sys.argv[1], sys.argv[2])
    print("PDF:", sys.argv[2])
