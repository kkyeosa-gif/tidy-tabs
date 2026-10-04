---
threads_url: https://www.threads.com/@tin_ylab/post/DeEmi_2m4B7
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-add-running-balance-column-in.html
title: How to Add a Running Balance Column in Excel and Google Sheets
labels: excel-formulas, cash-flow, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel/Google Sheets steps from Microsoft Support / Google Docs Editors Help (not tested here).
image_prompts: A farmers market vendor's metal cash box open on a folding table with cash and coins inside, next to a small paper receipt notepad with blurred handwritten entries, no visible logos.
image_alt: Open cash box with bills and coins beside a handwritten receipt notepad
search_description: Add a running balance column to a cash log in Excel or Google Sheets with one formula, plus a troubleshooting guide and free template.
threads: Redoing cash box math by hand every time money moves gets old fast.\nOne formula carries the balance down for you: =E2+C3-D3, adds money in, subtracts money out.\nStart at $250, run four transactions, land on $445.70 without touching a calculator.
---
Keep a running balance by adding each row's money in and subtracting its money out from the balance directly above it. Starting from an opening balance you type into the first row, the formula for every row after that is `=E2+C3-D3`, copied straight down the column.

> Works in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps below come from Microsoft Support and Google Docs Editors Help, not tested here.

![Cash log listing market sales and expenses next to a running balance column](images/2026-09-29-team-04-running-balance-column-excel-google-sheets/running-balance-cash-log-template.png) *LibreOffice Calc 24.2 PDF export of the Cash log tab with sample entries.*

## Steps

1. Set up five columns: Date, Description, Money in, Money out, Balance.
2. In the first data row, type your opening balance directly into the Balance column. There's no formula here, just the number you're starting with.
3. In the second data row's Balance cell, enter `=E2+C3-D3` (adjust the row numbers to match your sheet). This takes the balance above, adds this row's money in, and subtracts this row's money out.
4. Copy that formula down for every new row. Each one automatically points to the balance in the row just above it.
5. Add a new row at the top of your cash box or bank log each time money moves, and drag the formula down one more row.

### In Excel

- Select the cell with your first Balance formula, then drag the small square at its bottom-right corner (the fill handle) down through the rows you need. Excel adjusts the relative references (E2, C3, D3) automatically as it fills.
- To fill a longer range at once, select the formula cell plus all the empty cells below it, then press **Ctrl+D** to fill down.

### In Google Sheets

- Select the formula cell, grab the fill handle at its bottom-right corner, and drag down. Sheets updates the relative references the same way Excel does.
- You can also select the range and use **Edit > Fill down**, or the keyboard shortcut **Ctrl+D**.

Both apps treat `E2`, `C3`, and `D3` as relative references by default, which is what makes the formula shift down correctly as you copy it. If you're not sure why a reference changes (or doesn't) when you copy a formula, Microsoft's and Google's help centers cover the basics of relative references and filling formulas down a column.

## Example

Sample data below is fictional, made up for this post: a cash box for a craft fair and farmers market seller.

| Date | Description | Money in | Money out | Balance |
|---|---|---|---|---|
| 09/01/2026 | Opening balance | | | $250.00 |
| 09/06/2026 | Farmers market cash sales | $180.00 | | $430.00 |
| 09/09/2026 | Bought jars and labels | | $45.50 | $384.50 |
| 09/15/2026 | Etsy payout deposited | $96.20 | | $480.70 |
| 09/20/2026 | Paid booth fee, Oct market | | $35.00 | $445.70 |

Each Balance cell after the opening row uses the running formula, for example row 4: `=E3+C4-D4`, which reads $384.50 (previous balance) plus $96.20 (money in) minus $0.00 (no money out) for $480.70.

## Troubleshooting

### The first row's formula shows an error or a blank

The first row has no balance above it to reference, so it can't use the running formula. Type your opening balance as a plain number in that row instead of a formula, then start the formula from the second row down.

### Dragging the formula seems to create a circular chain

This is normal, not a mistake. Each row's Balance formula depends on the row above it, so the whole column forms one running chain back to your opening balance. If you see an actual circular reference warning, you likely pointed a formula at its own row (for example `=E3+C3-D3` in row 3) instead of the row above it.

### The balance goes negative

A negative balance usually means real activity, not a broken formula: you recorded more money out than you had on hand. Check that you didn't skip a deposit row, enter a payout twice, or put an amount in the wrong column (Money in vs. Money out).

### Balances don't match your actual cash count

Compare row by row against receipts or bank entries. A single transposed digit, a missing row, or an amount entered in the wrong column throws off every balance below it, since each one builds on the last.

### Copying the formula overwrites the wrong cell

Make sure you're dragging the fill handle down the Balance column only, not across into Money in or Money out. If you copy and paste instead of dragging, paste with **Paste Special > Formulas** so you don't overwrite number formatting in the column.

## Template

[Download the .xlsx](templates/tidy-tabs-running-balance-cash-log.xlsx)

The download has one tab, **Cash log**, set up with the Date, Description, Money in, Money out, and Balance columns from the example above, formulas already in place. To use it in Google Sheets, go to **File > Import > Upload**, choose the file, and select **Insert new sheet(s)**.

Reference pages for filling formulas down and relative references:
- [Microsoft Office Support](https://support.microsoft.com/en-us/office)
- [Google Docs Editors Help](https://support.google.com/docs)
