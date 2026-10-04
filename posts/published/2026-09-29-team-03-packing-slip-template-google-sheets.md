---
threads_url: https://www.threads.com/@tin_ylab/post/DeEkdXCnD3E
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-build-packing-slip-template-in.html
title: How to Build a Packing Slip Template in Google Sheets
labels: page-setup, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Google Sheets print steps from Google Docs Editors Help (not tested here).
image_prompts: A small craft business packing table with an open shipping box, a roll of packing tape, and a printed shipping label being applied to the box, label text blurred and unreadable, no visible logos.
image_alt: Shipping box on a packing table with tape and a blurred address label
search_description: Build a one-page packing slip template with an Order #, Ship to address, and a self-totaling items table sized to print on US Letter paper.
threads: 9 items packed, one formula: =SUM(D11:D14) totals the Qty packed column on a one-page packing slip.\nSet the page to Letter, fit to one page, and Excel or Sheets keeps it there.\nOrdered 4 birthday cards but only packed 3? The total still adds up, it just counts what shipped.
---
Build a one-page packing slip with an Order #, a Ship to address, and an items table that totals itself with `=SUM(D11:D14)` under Qty packed, then set the page to US Letter, portrait, fit to one page.

> Works in: LibreOffice Calc 24.2.7 (Linux), where this template was built and the Total packed formula and one-page print were verified. Excel and Google Sheets print steps below come from Microsoft Support and Google Docs Editors Help. Neither was hands-on tested here.

![Packing slip with order details, a shipping address, and an itemized quantity table](images/2026-09-29-team-03-packing-slip-template-google-sheets/packing-slip-template.png) *LibreOffice Calc 24.2 PDF export of the Packing slip tab with sample order data.*

## Steps

1. In row 1, type your business name, then add a **PACKING SLIP** label a row or two below it.
2. Add labeled fields for **Order #**, **Order date**, and **Ship to**, with the customer's name and address typed under Ship to across a few rows.
3. A few rows down, add column headers for the items table: **Item**, **SKU**, **Qty ordered**, **Qty packed**.
4. List one item per row under those headers, one row per SKU you're shipping.
5. Below the last item row, add a **Total packed** label and sum the Qty packed column, for example `=SUM(D11:D14)` if your items run from row 11 to row 14.
6. Add a short thank-you line or contact email at the bottom for the customer.
7. Compare each row's Qty packed to Qty ordered before you print. If they don't match, that item shipped short.

### In Excel

1. Confirm the packing slip tab is the sheet showing on screen before you print (a workbook can hold other tabs, like notes, that you don't want to print).
2. Go to **Page Layout > Size** and choose **Letter**.
3. Go to **Page Layout > Scale to Fit** and set both Width and Height to **1 page**.
4. Open **File > Print** and check the preview shows 1 of 1 pages before you print.

### In Google Sheets

1. If you're starting from the downloaded file, go to **File > Import > Upload** to bring it into Sheets.
2. Click the **Packing slip** tab so it's the active sheet.
3. Go to **File > Print**.
4. In the print settings panel, set Paper size to **Letter** and Print to **Current sheet**, so only the active tab is included.
5. Set scale to **Fit to page** and confirm the preview reads 1 of 1 pages, then print.

## Example

Sample data below is fictional, for a small craft business shipping an online order.

| Field | Value |
|---|---|
| Order # | 1042 |
| Order date | 09/26/2026 |
| Ship to | Priya Nair, 48 Willow St, Burlington, VT 05401 |

| Item | SKU | Qty ordered | Qty packed |
|---|---|---|---|
| Lavender soy candle, 8 oz | CND-LAV8 | 2 | 2 |
| Charcoal soap bar | SOP-CHR | 3 | 3 |
| Brass hoop earrings | EAR-HOOP | 1 | 1 |
| Letterpress birthday card | CRD-BDAY | 4 | 3 |

Total packed: **9**, from `=SUM(D11:D14)`, which is 2 + 3 + 1 + 3.

Notice the birthday card row: 4 were ordered but only 3 packed. The formula still adds up correctly; it just totals what you actually packed, not what was ordered. Add a note in a blank row below the table explaining the short shipment before you print.

## Troubleshooting

### The printed packing slip is 2 pages instead of 1
This happens when the workbook prints every tab, not just the packing slip. If your file has a second tab (notes, instructions, or anything else), select only the **Packing slip** tab, or set the print range to the current sheet, before you print.

### Total packed doesn't update when you add a row
The `SUM()` range doesn't cover the new row. If you add items above row 14, extend the formula, for example `=SUM(D11:D15)`.

### The ZIP code loses its leading zero
A ZIP like 05401 turns into 5401 when the cell is formatted as a number. Format the Ship to cells as **Plain text** before typing or pasting the address.

### Order date shows a number like 46652 instead of a date
The cell is formatted as General or Number instead of Date. Select the cell and apply a date format.

### Qty packed is less than Qty ordered and nothing flags it
The template doesn't color or warn on a shortfall by default; it only totals whatever numbers are in the Qty packed column. Scan the two columns side by side before you print, and add a note in a blank row for any item that shipped short.

## Template

[Download the .xlsx](templates/tidy-tabs-packing-slip-template.xlsx)

The file has two tabs: **Packing slip**, the printable one-page slip with the Order #, Ship to block, items table, and Total packed formula, and **How to use**, a notes tab with fill-in instructions that isn't meant to print. To open it in Google Sheets, go to **File > Import > Upload** and select the file.

For the page setup and print-range steps above, see Google's help on [printing from Google Sheets](https://support.google.com/docs/answer/7663148) and Microsoft's [print help search results](https://support.microsoft.com/en-us/search?query=fit%20worksheet%20to%20one%20page) for Excel's scale-to-fit options.
