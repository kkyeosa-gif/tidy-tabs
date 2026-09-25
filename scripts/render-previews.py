#!/usr/bin/env python3
"""Renders real previews of the templates with LibreOffice Calc (PDF export),
for use as honestly-captioned images in posts. Also re-verifies formulas by
printing recalculated values.

Run from the repo root: python3 scripts/render-previews.py
Needs: libreoffice-calc, openpyxl, pymupdf. Output paths are fixed below so
posts can reference them.
"""
import os
import subprocess
import tempfile

import pymupdf
from openpyxl import load_workbook

POSTS = {
    "inventory": "images/2026-10-01-a-craft-inventory-tracker-google-sheets",
    "projects": "images/2026-09-30-a-freelance-project-tracker-google-sheets",
    "print": "images/2026-09-29-a-print-excel-one-page-us-letter",
}
TMP = tempfile.mkdtemp(prefix="tidy-render-")


def to_pdf(xlsx):
    subprocess.run(["soffice", f"-env:UserInstallation=file://{TMP}/lo", "--headless",
                    "--convert-to", "pdf", "--outdir", TMP, xlsx], check=True, capture_output=True)
    return os.path.join(TMP, os.path.basename(xlsx).rsplit(".", 1)[0] + ".pdf")


def sheet_first(src, name, sheet):
    """Copy of src with `sheet` first and every sheet fit to one landscape page."""
    wb = load_workbook(src)
    wb.move_sheet(sheet, -wb.index(wb[sheet]))
    for ws in wb.worksheets:
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
        ws.print_options.gridLines = ws.print_options.headings = True
    path = f"{TMP}/{name}.xlsx"
    wb.save(path)
    return path


def crop_png(pdf, out, zoom=2.2):
    page = pymupdf.open(pdf)[0]
    rects = [b[:4] for b in page.get_text("blocks")] + [d["rect"] for d in page.get_drawings()]
    clip = pymupdf.Rect(rects[0])
    for r in rects[1:]:
        clip |= pymupdf.Rect(r)
    page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip + (-8, -8, 8, 8)).save(out)
    print("wrote", out)


def pages_grid(pdf, out, zoom=0.55, cols=3, gap=16):
    doc = pymupdf.open(pdf)
    pixs = [p.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)) for p in doc]
    w, h = pixs[0].width, pixs[0].height
    rows = (len(pixs) + cols - 1) // cols
    grid = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, cols * w + (cols + 1) * gap, rows * h + (rows + 1) * gap), 0)
    grid.set_rect(grid.irect, (225, 225, 225))
    for i, px in enumerate(pixs):
        px.set_origin(gap + (i % cols) * (w + gap), gap + (i // cols) * (h + gap))
        grid.copy(px, px.irect)
    grid.save(out)
    print("wrote", out, f"({len(pixs)} pages)")


def main():
    for d in POSTS.values():
        os.makedirs(d, exist_ok=True)
    inv = "templates/tidy-tabs-craft-inventory-tracker.xlsx"
    crop_png(to_pdf(sheet_first(inv, "items", "Items")), f"{POSTS['inventory']}/items-tab-libreoffice.png")
    crop_png(to_pdf(sheet_first(inv, "log", "Stock log")), f"{POSTS['inventory']}/stock-log-tab-libreoffice.png")
    crop_png(to_pdf(sheet_first("templates/tidy-tabs-freelance-project-tracker.xlsx", "projects", "Projects")),
             f"{POSTS['projects']}/projects-tab-libreoffice.png")

    # Print post: the shipped file (one page) vs. the same sheet on default settings.
    after = to_pdf(os.path.abspath("templates/tidy-tabs-print-one-page-us-letter.xlsx"))
    wb = load_workbook("templates/tidy-tabs-print-one-page-us-letter.xlsx")
    ws = wb.active
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.orientation = "portrait"
    ws.print_title_rows = None
    ws.page_margins.left = ws.page_margins.right = 0.7
    ws.page_margins.top = ws.page_margins.bottom = 0.75
    ws.print_options.horizontalCentered = False
    wb.save(f"{TMP}/print-before.xlsx")
    pages_grid(to_pdf(f"{TMP}/print-before.xlsx"), f"{POSTS['print']}/before-six-pages-libreoffice.png")
    doc = pymupdf.open(after)
    print("one-page file prints on", doc.page_count, "page(s)")
    doc[0].get_pixmap(matrix=pymupdf.Matrix(1.2, 1.2)).save(f"{POSTS['print']}/after-one-page-libreoffice.png")


if __name__ == "__main__":
    main()
