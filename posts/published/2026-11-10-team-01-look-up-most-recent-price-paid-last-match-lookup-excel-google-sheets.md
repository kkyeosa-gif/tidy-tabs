---
threads_url: https://www.threads.com/@tin_ylab/post/DfkSp8FG70d
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-look-up-most-recent-price-paid.html
title: How to Look Up the Most Recent Price Paid for an Item in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 25 workbooks and 960 comparisons (792 lookup cells and 168 summary cells) with 0 mismatches left. Checked: the sample, an empty log, the first, middle and last match, one match, no match, capitals and extra spaces in the log and in the item typed, a blank or text cost on the last row (skipped), a zero cost (counts), a text date on the last row and mid-log, blank dates, the same date twice, an older date on the last row, blank item rows, a full 200-row log, 20 items, and 10 random logs of 5 to 200 rows. Excel and Google Sheets were not used. Not run: more rows than the template holds, a non-US language setting.
search_description: Return the price from the last row that matches an item with LOOKUP(2,1/...), and flag when the last row is not the newest date. Free Excel template.
image_prompts: Paper receipts under a binder clip beside a pocket calculator and cotton wicks on a candle maker's workbench, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Paper receipts under a binder clip beside a calculator and cotton wicks on a workbench
threads: Pulled the wrong price from a purchase log?\nLOOKUP(2,1/(item match),cost) returns the last matching row, not the newest date.\nIn the sample, tissue paper returns $18.90 and gets flagged, because a row higher up is dated 10/27/2026 (made-up data).
---
Return the price from the last row that matches an item with `=LOOKUP(2,1/((Purchases!$D$4:$D$203=TRIM(A4))*ISNUMBER(Purchases!$C$4:$C$203)),Purchases!$C$4:$C$203)`. In the sample log, soy wax was bought three times and the formula returns $61.00, the cost on its last row (made-up prices).

The last row is not always the newest date, so the template also shows the newest date for the item and flags a mismatch.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (25 workbooks, 960 comparisons). The formulas use only IF, TRIM, SUMPRODUCT, LOOKUP, ISNUMBER, INDEX, ROW, MAX and N, which also exist in Excel and Google Sheets, but neither was opened for this post. How the array part of `LOOKUP(2,1/...)` and `SUMPRODUCT(MAX(...))` behaves there is not checked.

![Lookup table of six items with last price, purchase row, dates and a newer-date warning](images/2026-11-10-team-01-look-up-most-recent-price-paid-last-match-lookup-excel-google-sheets/last-price-paid-lookup-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Purchases** tab for the log and a **Lookup** tab where you type the items you want priced.

1. On **Purchases**, enter one purchase per row from row 4: the date in **A**, the item in **B**, the unit cost in **C**. The template holds rows 4 to 203.
2. In **D4**, type `=IF(TRIM(B4)="","",TRIM(B4))`. This is the item with stray outer spaces removed. Fill it down to row 203.
3. In **E4**, type `=N(A4)`. This turns each date into a number, and a blank or text date into 0. Fill it down to row 203.
4. On **Lookup**, type the items you want priced in **A4** down. Capitals do not matter.
5. In **B4**, count the usable purchases with `=IF(TRIM(A4)="","",SUMPRODUCT((Purchases!$D$4:$D$203=TRIM(A4))*ISNUMBER(Purchases!$C$4:$C$203)))`. A row only counts if the item matches and the cost is a number.
6. In **C4**, type `=IF(B4="","",IF(B4=0,"No purchase",LOOKUP(2,1/((Purchases!$D$4:$D$203=TRIM(A4))*ISNUMBER(Purchases!$C$4:$C$203)),Purchases!$C$4:$C$203)))`.
7. In **D4**, return the row number with the same formula, using `ROW(Purchases!$C$4:$C$203)` as the last argument. Wrap it in `IF(B4>0,...,"")` so unmatched items stay blank.
8. In **F4**, find the newest date among the matching rows with `=IF(ISNUMBER(D4),SUMPRODUCT(MAX((Purchases!$D$4:$D$203=TRIM(A4))*ISNUMBER(Purchases!$C$4:$C$203)*Purchases!$E$4:$E$203)),"")`.
9. In **E4**, show the date on the row the price came from: `=IF(ISNUMBER(D4),IF(ISNUMBER(INDEX(Purchases!$A$4:$A$203,D4-3)),INDEX(Purchases!$A$4:$A$203,D4-3),"No date"),"")`. The 3 is the number of rows above the first data row.
10. In **G4**, set the status: "No purchase" when nothing matched, "No date on that row" when the price row has no real date, "Newer date higher up" when **F4** is later than **E4**, otherwise "OK". Format **E** and **F** as dates and **C** as currency, then fill rows 4 to 23 down.

How it works: `1/(...)` turns matching rows into 1 and every other row into a #DIV/0! error. `LOOKUP(2,...)` looks for a 2, which is bigger than anything in the list, so it settles on the last number it can use and returns the cost from that row.

Help page: Microsoft's [LOOKUP function](https://support.microsoft.com/en-us/office/lookup-function-446d94af-663b-451d-8251-369d5e3864cb). It describes the function, not our test. We did not find a Google help page to link.

## Example

Sample data is fictional: 13 purchases and 6 items typed on the Lookup tab, with made-up prices.

| Item typed | Found | Last price | Row | Date on row | Newest date | Status |
|---|---|---|---|---|---|---|
| Soy wax, 10 lb | 3 | $61.00 | 12 | 11/03/2026 | 11/03/2026 | OK |
| Kraft mailer 6x6x4 | 3 | $0.71 | 14 | 11/09/2026 | 11/09/2026 | OK |
| "  tissue paper, 500 sheets " | 3 | $18.90 | 15 | 10/15/2026 | 10/27/2026 | Newer date higher up |
| Lavender oil, 4 oz | 2 | $15.10 | 13 | 11/05/2026 | 11/05/2026 | OK |
| 8 oz tin | 2 | $1.20 | 16 | 11/07/2026 | 11/07/2026 | OK |
| Wax melt molds | 0 | No purchase | | | | No purchase |

The tissue paper item was typed in lower case with extra spaces and still matched. Its last row is dated 10/15/2026, but an earlier row is dated 10/27/2026, so the status flags it. The price shown is the last row's $18.90, not necessarily the latest purchase.

The summary block counts 6 items typed, 5 with a price, 1 with no purchase and 1 with a newer date higher up. The CHECK cell says OK.

## Troubleshooting

### The price is from the last row, not the newest date
That is how `LOOKUP(2,1/...)` works: it returns the last matching row of the list. If your log is not in date order, read the Status column. "Newer date higher up" means another matching row has a later date. Sort the log by date, or read the price from that row.

### An item shows No purchase
Names must match after outer spaces are removed and capitals are ignored. "Kraft mailer" does not match "Kraft mailer 6x6x4". Copy the item name from the log instead of retyping it.

### A purchase with no cost is skipped
Rows with a blank or text cost are not counted, so the formula returns the previous matching row. A cost of 0 counts as a price. Fix text costs first, or see how to convert text to numbers.

### Status says No date on that row, or the CHECK cell says Check
Some date cells in the log are text or blank. The template treats a text date as no date in the newest-date comparison and counts these rows in the CHECK block. In LibreOffice, a date-like text such as "11/05/2026" was read as a date inside arithmetic. Do not expect the same in Excel or Google Sheets; retype the date as a real date.

### The formula errors in Excel or Google Sheets
We did not run either. Microsoft documents LOOKUP, but how `LOOKUP(2,1/...)` and `SUMPRODUCT(MAX(...))` handle arrays there is not confirmed here. If a cell errors, check the ranges first.

## Template

[Download the .xlsx](templates/tidy-tabs-last-price-paid.xlsx)

The file has three tabs. **Purchases** has the log in **A:C** and the helper columns **D** and **E**, rows 4 to 203. **Lookup** holds 20 item slots in **A4:A23**, the result columns **B:G**, and a summary with a CHECK cell in **I3:J9**. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
