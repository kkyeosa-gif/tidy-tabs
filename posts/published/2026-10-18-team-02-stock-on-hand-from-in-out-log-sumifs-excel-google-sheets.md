---
threads_url: https://www.threads.com/@tin_ylab/post/Depfzm2gTAB
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-track-stock-on-hand-from-log-of.html
title: How to Track Stock on Hand From a Log of Items In and Out in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 5 sample SKUs were checked against a Python per-SKU tally, plus edge cases such as an unused SKU, more sold than received, a lower case SKU, a SKU with a trailing space, a quantity typed as text, a blank SKU and a Type outside the list. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Track stock on hand from an in and out log with SUMIFS in Excel and Google Sheets. Subtract Sold from Received per SKU. Free stock log template.
image_prompts: Wooden storage shelf with stacked cut soap bars wrapped in plain paper, two lidded candle jars and a small cardboard shipping box beside them, soft daylight from a side window, no screens, no readable text, no labels, no logos
image_alt: Stacked soap bars and lidded candle jars on a shelf beside a shipping box
threads: Stock on hand is received minus sold, per SKU, from one log.\n=SUMIFS(Log!$D$2:$D$500,Log!$B$2:$B$500,A2,Log!$C$2:$C$500,"Received") does the first half.\nOne trailing space in a SKU and that row silently drops out, so the sheet flags it.
---
Keep one log of every receipt and sale, then total each SKU with `=SUMIFS(Log!$D$2:$D$500,Log!$B$2:$B$500,A2,Log!$C$2:$C$500,"Received")` and the same formula with `"Sold"`, and subtract. Add a third SUMIFS for signed Adjustment rows so counts and breakage are included: on hand is `=B2-C2+D2`.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Stock tab with received, sold, adjustments, on hand and status for five SKUs](images/2026-10-18-team-02-stock-on-hand-from-in-out-log-sumifs-excel-google-sheets/stock-on-hand-log-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

You need two tabs. **Log** holds one row per stock movement: date in **A**, SKU in **B**, Type in **C**, Qty in **D**. **Stock** holds one row per SKU: SKU in **A**, then Received, Sold, Adjustments, On hand, Minimum and Status in **B:G**.

1. On a tab named **Log**, type the headers Date, SKU, Type, Qty and Note in **A1:E1**.
2. Add one row each time stock moves. Type Received and Sold quantities as positive numbers. Type an Adjustment with its sign: -2 for two broken bars, 3 for three found in a recount.
3. On a tab named **Stock**, type the headers SKU, Received, Sold, Adjustments, On hand, Minimum and Status in **A1:G1**. List each SKU once in **A2:A6** and type the count you want to reorder at in **F**.
4. In **B2**, type `=IF(A2="","",SUMIFS(Log!$D$2:$D$500,Log!$B$2:$B$500,A2,Log!$C$2:$C$500,"Received"))`. SUMIFS adds the Qty column for rows where the SKU and the Type both match.
5. In **C2**, type the same formula with `"Sold"`. In **D2**, use `"Adjustment"`.
6. In **E2**, type `=IF(A2="","",B2-C2+D2)`.
7. In **G2**, type `=IF(A2="","",IF(E2<0,"Check",IF(NOT(ISNUMBER(F2)),"OK",IF(E2<=F2,"Reorder","OK"))))`. A SKU at or below its minimum shows Reorder. A negative count shows Check.
8. Select **B2:G2** and drag the fill handle down. The ranges stop at row 500 of the Log and the template fills formulas to row 20 of the Stock tab.

The `$` signs lock the Log ranges so they stay put while `A2` moves to `A3`, `A4` and so on. The ranges stop at row 500 instead of using whole columns, so rows added below 500 are not counted.

Microsoft documents the function on its [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) page, and Google on its [SUMIFS](https://support.google.com/docs/answer/3238496) page.

### In Excel

The formulas work as typed. The template adds dropdowns for SKU and Type on the Log, which are optional and not needed for the formulas.

### In Google Sheets

The same formulas should work as typed, and you can fill down by dragging the small square at the corner of the selected cell. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small soap, candle and lip balm shop with 30 log rows from 09/01/2026 to 10/05/2026.

| SKU | Received | Sold | Adjustments | On hand | Minimum | Status |
|---|---|---|---|---|---|---|
| SOAP-OAT | 72 | 66 | 0 | 6 | 10 | Reorder |
| SOAP-LAV | 36 | 31 | -2 | 3 | 10 | Reorder |
| CNDL-CED | 36 | 24 | -1 | 11 | 6 | OK |
| CNDL-VAN | 36 | 21 | 0 | 15 | 6 | OK |
| BALM-LIP | 60 | 54 | 3 | 9 | 8 | OK |

SOAP-LAV is 36 - 31 - 2 = 3, so it needs reordering. The Stock tab also totals the whole log beside the table: 240 received, 196 sold and 0 net adjustments (-1, -2 and +3), so 44 on hand. The sum of the five On hand cells is also 44, and the check cell shows OK.

## Troubleshooting

### One SKU is too high and the check cell says CHECK
A SKU in the Log has a trailing space, such as `SOAP-OAT ` with a space after it. It does not match the Stock tab, so that sale is left out of every SKU. In the recalculated copy, a 2 unit sale typed this way left SOAP-OAT at 6 instead of 4, the log total at 42 against 44 on the Stock tab, and the check at CHECK. The Log's Row check column shows Check for that row. Retype the SKU, or pick it from the dropdown.

### A quantity is ignored
A Qty typed as text, such as `5` stored as text, is skipped by SUMIFS without an error. The totals do not change, the Row check column says Check and the check cell shows CHECK. Retype the number.

### On hand is negative
More was sold than received, or a Received row is missing. In the test, adding a 30 unit sale of CNDL-VAN gave -15 and a Status of Check. Add the missing receipt or fix the typo.

### A SKU with no log rows says Reorder
It has 0 on hand, which is at or below its minimum. In the test, a new SKU with a minimum of 5 and no rows showed 0 and Reorder. Add its first Received row.

### A row is ignored because of its Type
Received, Sold and Adjustment must be spelled exactly. A Type such as Returned is not counted, the row shows Check and the check cell shows CHECK. SUMIFS ignores upper and lower case, so `soap-oat` still counts.

## Limits

This counts units. It does not value inventory or work out cost of goods sold, which are tax and accounting questions for your accountant. On hand covers the whole log, with no date filter. SKUs containing `*` or `?` were not tested, because SUMIFS criteria can treat those characters as wildcards. Stock rows 2 to 20 and Log rows 2 to 500 have formulas.

## Template

[Download the .xlsx](templates/tidy-tabs-stock-on-hand-log.xlsx)

The file has three tabs. **Log** has 30 sample rows in yellow input cells, a Type dropdown, a SKU dropdown fed by the Stock tab and a Row check column. **Stock** has the five SKUs and minimums in yellow, formulas for Received, Sold, Adjustments, On hand and Status, and a totals and check block in **I:J**. **How to use** has short notes. The Log prints the first 40 rows, and every tab prints on US Letter.

The shop and numbers are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
