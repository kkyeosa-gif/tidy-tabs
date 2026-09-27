---
title: How to Build a Simple Invoice Template in Google Sheets and Excel (US Letter, Net 30)
labels: invoicing, freelance
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A stack of printed invoices fanned out next to a stamped envelope and a desk calculator, no visible text.
image_alt: Printed invoices and envelope stacked on a desk beside a calculator
search_description: Build a simple Net 30 invoice template in Google Sheets or Excel with automatic due dates, subtotal, and tax, sized to print on US Letter paper.
threads: Freelancers: stop typing your due date by hand. Set it as invoice date + 30 with a formula like =B2+30 and it updates itself every time you change the invoice date. Works the same way in Excel or Google Sheets.
---
Build an invoice with a header block, a line-item table, and a formula for the due date: type `=B2+30` in the due date cell (where B2 is the invoice date) to get a Net 30 date automatically. Add `=SUM()` for the subtotal and format the sheet for US Letter so it prints on one page.

> Works in: Excel (Microsoft 365 and Excel 2019+) and Google Sheets. Menu paths below come from Microsoft and Google's own help pages, not hands-on testing.

![Spreadsheet example with columns Field, Value](images/2026-09-27-zauto-02-how-to-build-a-simple-invoice-template-i/how-to-build-a-simple-invoice-template-in-google-s-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

1. In row 1, type your business name, address, and phone number, one per cell down column A.
2. A few rows down, add labels for **Invoice #**, **Invoice Date**, and **Due Date** in one column, with the values next to them.
3. Below that, add a **Bill To** section with the client's name and address.
4. Add a line-item table with headers: **Description**, **Qty**, **Rate**, **Amount**.
5. In the Amount column, multiply quantity by rate, for example `=C10*D10`.
6. Below the last line item, add **Subtotal**, **Tax**, and **Total** rows.
7. For Subtotal, use `=SUM(E10:E14)` (adjust the range to your rows).
8. For Tax, multiply the subtotal by your rate, for example `=E15*0.06`.
9. For Total, add Subtotal and Tax: `=E15+E16`.
10. In the Due Date cell, type `=` and click the Invoice Date cell, then type `+30` so it reads `=B2+30`.

### In Excel

1. Select the Invoice Date and Due Date cells.
2. Go to **Home > Number Format** and choose **Short Date**.
3. Go to **File > Page Setup**, set paper size to **Letter**, and choose **Fit Sheet on One Page**.

### In Google Sheets

1. Select the Invoice Date and Due Date cells.
2. Go to **Format > Number > Date** to format them as dates.
3. Go to **File > Print**, set paper size to **Letter**, and set scale to **Fit to width**.

## Example

Sample data below is fictional, for a freelance design business.

| Field | Value |
|---|---|
| Invoice # | 1042 |
| Invoice Date | 09/28/2026 |
| Due Date | 10/28/2026 |
| Bill To | Whitman Family Farm |

| Description | Qty | Rate | Amount |
|---|---|---|---|
| Logo design | 1 | $450.00 | $450.00 |
| Business card layout | 2 | $75.00 | $150.00 |

| | |
|---|---|
| Subtotal | $600.00 |
| Tax (6%) | $36.00 |
| Total | $636.00 |

The Due Date column reads 10/28/2026 because it's calculated as the Invoice Date plus 30 days, not typed in by hand.

## Troubleshooting

### Due date shows a number like 46683 instead of a date
The cell is formatted as General or Number instead of Date. Select the cell and apply a date format from **Home > Number Format** in Excel or **Format > Number > Date** in Google Sheets.

### Total doesn't change when you edit a line item
The Subtotal formula's range doesn't include the row you edited. Check that `SUM()` covers every line-item row, and extend the range if you add new rows in the middle of the table.

### Invoice prints across two pages
The columns are too wide for Letter paper. Set the print scale to **Fit to width** (Google Sheets) or **Fit Sheet on One Page** (Excel) before printing.

### Tax amount looks way too high
The formula multiplies by a whole number instead of a decimal, for example `*6` instead of `*0.06`. Fix the tax rate cell to read as a percentage or decimal.

### Dollar amounts show extra decimal places or no dollar sign
The cells are still in General format. Select the amount, subtotal, tax, and total cells and apply **Currency** formatting so they display as $450.00 instead of 450.

## Copy-paste setup

Type these headers and formulas directly into a blank sheet:

**Header block (column A, rows 1-6):**
```
Your Business Name
Your Address
Invoice #: [type number]
Invoice Date: [type date]
Due Date: =B4+30
Bill To: [client name]
```

**Line-item table headers (row 9):**
```
Description | Qty | Rate | Amount
```

**Amount formula (first line item row):**
```
=B10*C10
```

**Totals block:**
```
Subtotal: =SUM(D10:D14)
Tax (enter your rate): =D15*0.06
Total: =D15+D16
```

For official background on the menu paths used above, see Microsoft's guide on [applying number formats](https://support.microsoft.com/en-us/office/format-numbers-as-currency-dates-or-percentages-in-excel-8f9b9e78-2ff7-4b7f-87c7-ade54227a30e) and page setup, and Google's guide on [setting print options in Google Sheets](https://support.google.com/docs/answer/91062).
