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


def heroes_2026_10_06():
    t = "templates/"
    d = "images/2026-10-06-team-0%d-%s/"
    save(d % (1, "loan-payment-calculator-pmt-excel-google-sheets") + "loan-payment-calculator-template.png",
         render(t + "tidy-tabs-loan-payment-calculator.xlsx", "Loans", name="loans"))
    save(d % (2, "percent-change-month-over-month-excel-google-sheets") + "percent-change-template.png",
         render(t + "tidy-tabs-percent-change.xlsx", "Sales", name="pctchange"))
    save(d % (3, "rank-top-customers-excel-google-sheets") + "rank-top-customers-template.png",
         render(t + "tidy-tabs-rank-top-customers.xlsx", "Customers", name="rank"))
    save(d % (4, "invoice-number-generator-text-excel-google-sheets") + "invoice-number-generator-template.png",
         render(t + "tidy-tabs-invoice-number-generator.xlsx", "Invoices", name="invnum"))
    save(d % (5, "combine-address-columns-textjoin-excel-google-sheets") + "combine-address-columns-textjoin-template.png",
         render(t + "tidy-tabs-combine-address-columns.xlsx", "Addresses", hide_cols="H", name="addr"))


def heroes_2026_10_07():
    t = "templates/"
    d = "images/2026-10-07-team-0%d-%s/"
    save(d % (1, "sale-price-discount-calculator-excel-google-sheets") + "sale-price-discount-calculator-template.png",
         render(t + "tidy-tabs-sale-price-discount-calculator.xlsx", "Sale prices", hide_cols="F", name="saleprice"))
    save(d % (2, "shipping-cost-weight-tier-lookup-excel-google-sheets") + "shipping-weight-tier-lookup-template.png",
         render(t + "tidy-tabs-shipping-weight-tier-lookup.xlsx", "Orders", hide_cols="DE", name="shiptier"))
    save(d % (3, "reorder-point-calculator-excel-google-sheets") + "reorder-point-calculator-template.png",
         render(t + "tidy-tabs-reorder-point-calculator.xlsx", "Reorder", hide_cols="I", name="reorder"))
    save(d % (4, "freelance-hourly-rate-calculator-excel-google-sheets") + "freelance-hourly-rate-calculator-template.png",
         render(t + "tidy-tabs-freelance-hourly-rate-calculator.xlsx", "Hourly rate", name="hourly"))
    save(d % (5, "convert-text-to-numbers-dollar-signs-excel-google-sheets") + "convert-text-to-numbers-template.png",
         render(t + "tidy-tabs-convert-text-to-numbers.xlsx", "Pasted amounts", hide_cols="EF", name="text2num"))


def main():
    heroes_2026_10_07()
    return
    # older heroes (rebuilding rewrites their PNGs; the early return above skips them)
    t = "templates/"
    save("images/2026-09-25-a-split-city-state-zip-excel-google-sheets/split-city-state-zip-template.png",
         render(t + "tidy-tabs-split-city-state-zip.xlsx", "Split", name="split"))
    save("images/2026-09-28-a-keep-leading-zeros-zip-codes-excel/zip-codes-leading-zeros-fixed.png",
         render(t + "tidy-tabs-zip-codes-leading-zeros.xlsx", "Broken (numbers)", name="zip"))
    save("images/2026-09-30-a-freelance-project-tracker-google-sheets/freelance-project-tracker-template.png",
         render(t + "tidy-tabs-freelance-project-tracker.xlsx", "Projects", hide_cols="CFGHIKLM", name="proj"))
    save("images/2026-10-01-a-craft-inventory-tracker-google-sheets/craft-inventory-tracker-template.png",
         render(t + "tidy-tabs-craft-inventory-tracker.xlsx", "Items", hide_cols="CDEFGHLNO", name="items"))
    save("images/2026-09-29-team-01-dependent-dropdown-list-excel-google-sheets/dependent-dropdown-lists-tab.png",
         render(t + "tidy-tabs-dependent-dropdown-list.xlsx", "Lists", name="dropdown"))
    save("images/2026-09-29-team-02-client-contact-list-template-google-sheets/client-contact-list-template.png",
         render(t + "tidy-tabs-client-contact-list.xlsx", "Clients", hide_cols="G", name="clients"))
    save("images/2026-09-29-team-03-packing-slip-template-google-sheets/packing-slip-template.png",
         render(t + "tidy-tabs-packing-slip-template.xlsx", "Packing slip", name="packing"))
    save("images/2026-09-29-team-04-running-balance-column-excel-google-sheets/running-balance-cash-log-template.png",
         render(t + "tidy-tabs-running-balance-cash-log.xlsx", "Cash log", name="cashlog"))
    save("images/2026-09-29-team-05-highlight-duplicate-values-excel-google-sheets/highlight-duplicate-emails-template.png",
         render(t + "tidy-tabs-highlight-duplicates-sample.xlsx", "Orders", name="dupes"))
    # 2026-09-30 team posts
    d = "images/2026-09-30-team-0%d-%s/"
    save(d % (1, "receivables-aging-report-excel-google-sheets") + "receivables-aging-report-template.png",
         render(t + "tidy-tabs-receivables-aging-report.xlsx", "Aging", name="aging"))
    save(d % (2, "time-to-decimal-hours-excel-google-sheets") + "time-to-decimal-hours-template.png",
         render(t + "tidy-tabs-time-to-decimal-hours.xlsx", "Shifts", hide_cols="AGI", name="shifts"))
    save(d % (3, "markup-vs-margin-calculator-excel-google-sheets") + "markup-vs-margin-calculator-template.png",
         render(t + "tidy-tabs-markup-vs-margin-calculator.xlsx", "Pricing", hide_cols="EI", name="pricing"))
    save(d % (4, "business-days-ship-by-date-excel-google-sheets") + "business-days-ship-by-date-template.png",
         render(t + "tidy-tabs-business-days-ship-by-date.xlsx", "Orders", hide_cols="BF", name="orders"))
    save(d % (5, "vlookup-price-list-excel-google-sheets") + "vlookup-price-list-template.png",
         render(t + "tidy-tabs-vlookup-price-list.xlsx", "Order lines", hide_cols="FG", name="orderlines"))
    # 2026-10-02 team posts
    d = "images/2026-10-02-team-0%d-%s/"
    save(d % (1, "break-even-calculator-excel-google-sheets") + "break-even-calculator-template.png",
         render(t + "tidy-tabs-break-even-calculator.xlsx", "Break-even", name="breakeven"))
    save(d % (2, "subscription-renewal-tracker-excel-google-sheets") + "subscription-renewal-tracker-template.png",
         render(t + "tidy-tabs-subscription-renewal-tracker.xlsx", "Subscriptions", hide_cols="CDHIJK", name="subs"))
    save(d % (3, "budget-vs-actual-variance-excel-google-sheets") + "budget-vs-actual-variance-template.png",
         render(t + "tidy-tabs-budget-vs-actual.xlsx", "Budget vs actual", hide_cols="GHI", name="budget"))
    save(d % (4, "farmers-market-sales-log-excel-google-sheets") + "farmers-market-sales-log-template.png",
         render(t + "tidy-tabs-farmers-market-sales-log.xlsx", "Sales", name="market"))
    save(d % (5, "clean-customer-list-trim-proper-excel-google-sheets") + "clean-customer-list-trim-proper-template.png",
         render(t + "tidy-tabs-clean-customer-list.xlsx", "Clean names", hide_cols="DE", name="clean"))
    # Print post: the real 6-page default printout next to the real 1-page result.
    d = "images/2026-09-29-a-print-excel-one-page-us-letter/"
    save(d + "print-excel-one-page-before-after.png",
         Image.open(d + "before-six-pages-libreoffice.png").convert("RGB"),
         Image.open(d + "after-one-page-libreoffice.png").convert("RGB"))


if __name__ == "__main__":
    main()
