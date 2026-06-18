from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
PDF = ROOT / "Netflix_Content_Analysis_Complete_Project_Handbook.pdf"
OUT = ROOT / "handbook_pdf_review"
OUT.mkdir(exist_ok=True)

pdf = pdfium.PdfDocument(PDF)
thumbs = []

for index in range(len(pdf)):
    page = pdf[index]
    image = page.render(scale=1.35).to_pil().convert("RGB")
    path = OUT / f"page-{index + 1:03d}.png"
    image.save(path)
    thumb = image.copy()
    thumb.thumbnail((390, 505))
    thumbs.append((index + 1, thumb))

for sheet_index in range(0, len(thumbs), 8):
    batch = thumbs[sheet_index:sheet_index + 8]
    sheet = Image.new("RGB", (820, 2140), "white")
    draw = ImageDraw.Draw(sheet)
    for local, (page_number, thumb) in enumerate(batch):
        col = local % 2
        row = local // 2
        x = 15 + col * 405
        y = 30 + row * 525
        sheet.paste(thumb, (x, y))
        draw.text((x, y - 18), f"Page {page_number}", fill="black")
    sheet.save(OUT / f"contact-{sheet_index // 8 + 1}.png")

print(f"Rendered {len(pdf)} pages to {OUT}")
