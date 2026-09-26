#!/usr/bin/env python3
"""Builds each post's hero image: a real LibreOffice render of the post's
template (or a before/after of the real print output), placed on a 1600x900
(16:9) canvas. The hero is the first image in the post, so Blogger uses it
as og:image, and Google Discover wants >=1200px-wide 16:9 images that are
"relevant and representative" (see tasks/briefs.md 2026-09-26).

Run from the repo root: python3 scripts/make-heroes.py
Needs libreoffice-calc, openpyxl, pymupdf, pillow.
"""
import io
import os

import pymupdf
from openpyxl import load_workbook
from PIL import Image, ImageFilter

from importlib.util import spec_from_file_location, module_from_spec

_spec = spec_from_file_location("rp", os.path.join(os.path.dirname(__file__), "render-previews.py"))
rp = module_from_spec(_spec)
_spec.loader.exec_module(rp)

W, H, PAD = 1600, 900, 50
BG = (244, 246, 243)


def render(src, sheet, hide_cols="", name="hero", zoom=3.0):
    """PDF-render one sheet (optionally hiding columns) and return a cropped PIL image."""
    wb = load_workbook(src)
    wb.move_sheet(sheet, -wb.index(wb[sheet]))
    for ws in wb.worksheets:
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
        ws.print_options.gridLines = True
    for col in hide_cols:
        wb[sheet].column_dimensions[col].hidden = True
    path = f"{rp.TMP}/{name}.xlsx"
    wb.save(path)
    page = pymupdf.open(rp.to_pdf(path))[0]
    rects = [b[:4] for b in page.get_text("blocks")] + [d["rect"] for d in page.get_drawings()]
    clip = pymupdf.Rect(rects[0])
    for r in rects[1:]:
        clip |= pymupdf.Rect(r)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip + (-6, -6, 6, 6))
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def card(img, max_w, max_h):
    scale = min(max_w / img.width, max_h / img.height)
    img = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
    shadow = Image.new("RGBA", (img.width + 24, img.height + 24), (0, 0, 0, 0))
    shadow.paste((0, 0, 0, 60), (12, 16, img.width + 12, img.height + 16))
    shadow = shadow.filter(ImageFilter.GaussianBlur(8))
    shadow.paste(img, (12, 12))
    return shadow


def save(out, *images):
    canvas = Image.new("RGB", (W, H), BG)
    n = len(images)
    slot_w = (W - PAD * (n + 1)) // n
    for i, img in enumerate(images):
        c = card(img, slot_w, H - 2 * PAD)
        x = PAD + i * (slot_w + PAD) + (slot_w - c.width) // 2
        canvas.paste(c, (x, (H - c.height) // 2), c)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    canvas.save(out, optimize=True)
    print("wrote", out, os.path.getsize(out) // 1024, "KB")


def main():
    t = "templates/"
    save("images/2026-09-25-a-split-city-state-zip-excel-google-sheets/split-city-state-zip-template.png",
         render(t + "tidy-tabs-split-city-state-zip.xlsx", "Split", name="split"))
    save("images/2026-09-28-a-keep-leading-zeros-zip-codes-excel/zip-codes-leading-zeros-fixed.png",
         render(t + "tidy-tabs-zip-codes-leading-zeros.xlsx", "Broken (numbers)", name="zip"))
    save("images/2026-09-30-a-freelance-project-tracker-google-sheets/freelance-project-tracker-template.png",
         render(t + "tidy-tabs-freelance-project-tracker.xlsx", "Projects", hide_cols="CFGHIKLM", name="proj"))
    save("images/2026-10-01-a-craft-inventory-tracker-google-sheets/craft-inventory-tracker-template.png",
         render(t + "tidy-tabs-craft-inventory-tracker.xlsx", "Items", hide_cols="CDEFGHLNO", name="items"))
    # Print post: the real 6-page default printout next to the real 1-page result.
    d = "images/2026-09-29-a-print-excel-one-page-us-letter/"
    save(d + "print-excel-one-page-before-after.png",
         Image.open(d + "before-six-pages-libreoffice.png").convert("RGB"),
         Image.open(d + "after-one-page-libreoffice.png").convert("RGB"))


if __name__ == "__main__":
    main()
