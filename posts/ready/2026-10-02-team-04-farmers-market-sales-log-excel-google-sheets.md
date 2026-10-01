---
title: How to Track Farmers Market Sales in Excel and Google Sheets
labels: excel-formulas, farmers-market, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Track farmers market sales in a spreadsheet: SUMPRODUCT for qty x price, plus one market day's revenue and a SUMIFS check. Free .xlsx template.
---
Multiply quantity by price and add it all up in one step with `=SUMPRODUCT(C2:C20,D2:D20)`. To get one market day's revenue, wrap the dates in the formula, as in `=SUMPRODUCT((A2:A20=F2)*C2:C20*D2:D20)`.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every total below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

## Steps

Set up two tabs: **Sales** (one row per item sold per market day) and **Daily totals** (one row per market day). The template uses rows 2 to 500 so you can keep adding rows. The short ranges in the first paragraph are the same idea on a 20-row list.

1. On the Sales tab, use these headers in row 1: Market date, Item, Qty sold, Price, Line total.
2. Type one row per item per market day. Enter dates as MM/DD/YYYY, such as 09/12/2026. Qty and Price must be plain numbers.
3. In **E2** (Line total), type `=C2*D2` and fill it down. This column is optional for the totals. It gives you a check on the SUMPRODUCT result.
4. On the Daily totals tab, type each market date in **A2**, **A3**, **A4**, and so on.
5. In **B2** (Revenue), type `=SUMPRODUCT((Sales!$A$2:$A$500=A2)*Sales!$C$2:$C$500*Sales!$D$2:$D$500)` and fill it down. The part in parentheses is TRUE for rows on that date. SUMPRODUCT multiplies it by quantity and price, then adds the results.
6. In **C2**, type `=SUMIFS(Sales!$E$2:$E$500,Sales!$A$2:$A$500,A2)`. This adds the Line total column for the same date. It should match column B.
7. In **D2**, type `=SUMIFS(Sales!$C$2:$C$500,Sales!$A$2:$A$500,A2)` for the number of items sold that day.
8. In **G1**, type `=SUMPRODUCT(Sales!C2:C500,Sales!D2:D500)` for the total across every day. In **G2**, type `=SUM(Sales!E2:E500)` as the check.
9. Format B, C, and G as currency and the date columns as dates.

### In Excel

The formulas work as typed. Microsoft documents the function on its [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) page.

### In Google Sheets

The same formulas work as typed. Google documents the function on its [SUMPRODUCT](https://support.google.com/docs/answer/3094294) page.

## Example

Sample data below is fictional. It is a small candle, soap, and jewelry booth over three Saturdays.

| Market date | Item | Qty sold | Price | Line total |
|---|---|---|---|---|
| 09/12/2026 | Lavender soy candle, 8 oz | 9 | $18.00 | $162.00 |
| 09/12/2026 | Oatmeal goat milk soap bar | 14 | $8.00 | $112.00 |
| 09/12/2026 | Brass hoop earrings | 3 | $24.00 | $72.00 |
| 09/12/2026 | Letterpress birthday card | 12 | $6.00 | $72.00 |
| 09/19/2026 | Lavender soy candle, 8 oz | 11 | $18.00 | $198.00 |
| 09/19/2026 | Oatmeal goat milk soap bar | 10 | $8.00 | $80.00 |
| 09/19/2026 | Cedar soy candle, 8 oz | 6 | $18.00 | $108.00 |
| 09/19/2026 | Letterpress birthday card | 8 | $6.00 | $48.00 |
| 09/19/2026 | Brass hoop earrings | 2 | $24.00 | $48.00 |
| 09/26/2026 | Lavender soy candle, 8 oz | 7 | $18.00 | $126.00 |
| 09/26/2026 | Oatmeal goat milk soap bar | 16 | $8.00 | $128.00 |
| 09/26/2026 | Brass hoop earrings | 4 | $24.00 | $96.00 |
| 09/26/2026 | Letterpress birthday card | 10 | $6.00 | $60.00 |

The Daily totals tab then shows:

| Market date | Revenue (SUMPRODUCT) | Revenue (SUMIFS check) | Items sold |
|---|---|---|---|
| 09/12/2026 | $418.00 | $418.00 | 38 |
| 09/19/2026 | $482.00 | $482.00 | 37 |
| 09/26/2026 | $410.00 | $410.00 | 37 |

Both grand totals, G1 and G2, show $1,310.00. For 09/12/2026 the hand math is 9 x $18.00 + 14 x $8.00 + 3 x $24.00 + 12 x $6.00 = $418.00.

SUMIFS in the check column is the same function used for monthly totals in other Tidy Tabs posts. Here it only confirms the SUMPRODUCT column.

## Troubleshooting

### The revenue cell shows #VALUE!
SUMPRODUCT with the multiplication form fails when a Qty or Price cell holds text, such as `$18.00` typed with a stray space or a pasted `12 pcs`. The template notes warn about this. I did not reproduce the error in LibreOffice for this post, so check the cause by looking for left-aligned numbers in columns C and D. Retype them as plain numbers.

### A market day shows $0.00
The date in Daily totals does not match the date in Sales. Dates stored as text, or dates with a time attached, will not match a plain date. Retype the date in both places as MM/DD/YYYY and format the cells as Date.

### The SUMPRODUCT and SUMIFS columns disagree
A Line total cell is blank or has an old typed number. Fill **E2** down to every row on the Sales tab, then compare again. The two columns should match to the cent.

### New rows are missing from the totals
The ranges stop at row 500. If you log past row 500, extend them, for example `$A$2:$A$2000`, in every formula.

### The new market day does not appear
Daily totals needs one row per date. Type the new date in the next free cell in column A and fill the formulas down from the row above.

## Template

[Download the .xlsx](templates/tidy-tabs-farmers-market-sales-log.xlsx)

The file has three tabs. **Sales** holds the 13 sample rows with the Line total formula. **Daily totals** has the SUMPRODUCT revenue, the SUMIFS check, items sold for each of the three market dates, and the two grand totals in G1 and G2. **How to use** has short fill-in notes.

The vendor, items, and sales are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
