from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, PageBreak, Paragraph, Spacer,
    Table, TableStyle, Preformatted, KeepTogether
)


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "Netflix_Content_Analysis_Complete_Project_Handbook.docx"
OUTPUT = ROOT / "Netflix_Content_Analysis_Complete_Project_Handbook.pdf"

RED = colors.HexColor("#E50914")
INK = colors.HexColor("#17171B")
GRAY = colors.HexColor("#64646E")
LIGHT = colors.HexColor("#F4F5F7")
DARK = colors.HexColor("#17171B")


def iter_blocks(parent):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield DocxParagraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield DocxTable(child, parent)


def run_markup(paragraph):
    if not paragraph.runs:
        return escape(paragraph.text)
    parts = []
    for run in paragraph.runs:
        text = escape(run.text).replace("\n", "<br/>")
        if not text:
            continue
        if run.bold:
            text = f"<b>{text}</b>"
        if run.italic:
            text = f"<i>{text}</i>"
        parts.append(text)
    return "".join(parts)


styles = getSampleStyleSheet()
body = ParagraphStyle(
    "Body", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.2,
    leading=11.4, textColor=INK, spaceAfter=5.5
)
h1 = ParagraphStyle(
    "H1", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=16,
    leading=19, textColor=INK, spaceBefore=12, spaceAfter=8, keepWithNext=True
)
h2 = ParagraphStyle(
    "H2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12,
    leading=14.5, textColor=RED, spaceBefore=10, spaceAfter=6, keepWithNext=True
)
h3 = ParagraphStyle(
    "H3", parent=styles["Heading3"], fontName="Helvetica-Bold", fontSize=9.8,
    leading=12, textColor=INK, spaceBefore=7, spaceAfter=3, keepWithNext=True
)
bullet = ParagraphStyle(
    "Bullet", parent=body, leftIndent=15, firstLineIndent=-8, bulletIndent=7,
    spaceAfter=3.5
)
small = ParagraphStyle(
    "Small", parent=body, fontSize=7.5, leading=9.3, textColor=GRAY
)
code_style = ParagraphStyle(
    "Code", fontName="Courier", fontSize=6.7, leading=8.3, textColor=INK,
    leftIndent=6, rightIndent=6, spaceAfter=0
)
cover_title = ParagraphStyle(
    "CoverTitle", fontName="Helvetica-Bold", fontSize=27, leading=31,
    textColor=INK, alignment=TA_CENTER, spaceBefore=65, spaceAfter=8
)
cover_sub = ParagraphStyle(
    "CoverSub", fontName="Helvetica", fontSize=13, leading=16,
    textColor=GRAY, alignment=TA_CENTER, spaceAfter=20
)


def footer(canvas, doc):
    canvas.saveState()
    width, height = letter
    canvas.setStrokeColor(colors.HexColor("#E2E3E6"))
    canvas.setLineWidth(0.4)
    canvas.line(0.75 * inch, 0.48 * inch, width - 0.75 * inch, 0.48 * inch)
    canvas.setFont("Helvetica", 7.2)
    canvas.setFillColor(GRAY)
    canvas.drawString(0.75 * inch, 0.28 * inch, "NETFLIX CONTENT ANALYSIS | COMPLETE PROJECT HANDBOOK")
    page = f"Page {doc.page}"
    canvas.drawRightString(width - 0.75 * inch, 0.28 * inch, page)
    canvas.restoreState()


docx = Document(SOURCE)
story = []
cover_seen = False
list_number = 0

for block in iter_blocks(docx):
    if isinstance(block, DocxParagraph):
        text = block.text.strip()
        xml = block._p.xml
        if "w:type=\"page\"" in xml:
            story.append(PageBreak())
            list_number = 0
            continue
        if not text:
            story.append(Spacer(1, 3))
            continue
        style_name = block.style.name if block.style else ""
        markup = run_markup(block)

        if text == "Netflix Content Analysis" and not cover_seen:
            story.append(Paragraph(markup, cover_title))
            cover_seen = True
        elif text.startswith("Beginner-to-Advanced"):
            story.append(Paragraph(markup, cover_sub))
        elif style_name == "Heading 1":
            story.append(Paragraph(markup, h1))
        elif style_name == "Heading 2":
            story.append(Paragraph(markup, h2))
        elif style_name == "Heading 3":
            story.append(Paragraph(markup, h3))
        elif style_name.startswith("List Bullet"):
            story.append(Paragraph(markup, bullet, bulletText="•"))
        elif style_name.startswith("List Number"):
            list_number += 1
            story.append(Paragraph(markup, bullet, bulletText=f"{list_number}."))
        else:
            alignment = block.alignment
            if alignment == 1:
                local = ParagraphStyle("CenterBody", parent=body, alignment=TA_CENTER)
                story.append(Paragraph(markup, local))
            else:
                story.append(Paragraph(markup, body))

    else:
        rows = len(block.rows)
        cols = len(block.columns)
        if rows == 1 and cols == 1:
            text = block.cell(0, 0).text
            if "\n" in text or any(token in text for token in ("const ", "for row", "credits =", "|--")):
                story.append(Table(
                    [[Preformatted(text, code_style)]],
                    colWidths=[6.35 * inch],
                    style=TableStyle([
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F2F4")),
                        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9DBDF")),
                        ("LEFTPADDING", (0, 0), (-1, -1), 8),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                        ("TOPPADDING", (0, 0), (-1, -1), 7),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                    ])
                ))
            else:
                story.append(Table(
                    [[Paragraph(escape(text).replace("\n", "<br/>"), body)]],
                    colWidths=[6.35 * inch],
                    style=TableStyle([
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9E9EA")),
                        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#EDC7CA")),
                        ("LEFTPADDING", (0, 0), (-1, -1), 12),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                        ("TOPPADDING", (0, 0), (-1, -1), 9),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                    ])
                ))
            story.append(Spacer(1, 5))
            continue

        data = []
        for r, row in enumerate(block.rows):
            current = []
            for cell in row.cells:
                text = escape(cell.text).replace("\n", "<br/>")
                current.append(Paragraph(text, small if cols >= 4 else body))
            data.append(current)

        available = 6.35 * inch
        if cols == 2:
            widths = [available * 0.27, available * 0.73]
        elif cols == 4:
            widths = [available * 0.18, available * 0.29, available * 0.23, available * 0.30]
        elif cols == 5:
            widths = [available * 0.15, available * 0.27, available * 0.19, available * 0.20, available * 0.19]
        else:
            widths = [available / cols] * cols
        table = Table(data, colWidths=widths, repeatRows=1, splitByRow=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), DARK),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#DADCE0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(table)
        story.append(Spacer(1, 6))

pdf = BaseDocTemplate(
    str(OUTPUT), pagesize=letter,
    leftMargin=0.75 * inch, rightMargin=0.75 * inch,
    topMargin=0.65 * inch, bottomMargin=0.62 * inch,
    title="Netflix Content Analysis - Complete Project Handbook",
    author="Kshitij",
)
frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id="normal")
pdf.addPageTemplates([PageTemplate(id="handbook", frames=frame, onPage=footer)])
pdf.build(story)
print(OUTPUT)
