---
title: How to Total Sales by Quarter in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against hand or Python math, including the quarter boundary days and a date typed as text. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Total sales by quarter with =YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0) and =SUMIFS($B$2:$B$200,$C$2:$C$200,"2026 Q2"). Free .xlsx template.
image_prompts: Home office desk with a paper ledger open to handwritten rows and a small stack of receipts under a brass paperweight, soft window light, no screens, no readable text, no logos
image_alt: Open paper ledger beside a receipt stack held down by a brass paperweight
threads: Need sales totals by quarter?\nTurn each date into a label with =YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0), then =SUMIFS($B$2:$B$200,$C$2:$C$200,"2026 Q2").\nThe year in the label keeps 2025 Q4 and 2026 Q4 apart. A second date-range formula checks the total.
---
Turn each date into a quarter label with `=YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0)`, then total a quarter with `=SUMIFS($B$2:$B$200,$C$2:$C$200,"2026 Q2")`. The year in the label keeps 2025 Q4 and 2026 Q4 apart.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to their help pages. Neither Excel nor Google Sheets was opened for this post.

![By quarter table totaling sales per quarter two ways with matching results](images/2026-10-11-team-01-quarterly-sales-totals-sumifs-excel-google-sheets/quarterly-sales-totals-by-quarter-template.png) *Render of the By quarter tab of the template, a PDF export from LibreOffice Calc 24.2 using fictional sample data*

## Steps

These are calendar quarters: January to March is Q1, April to June is Q2, July to September is Q3 and October to December is Q4. Fiscal or tax-filing quarters are not covered here.

1. In row 1 of a tab named Sales, type these headers in A to C: Sale date, Amount, Quarter.
2. Type each sale date in **A2** and down as a real date, such as 05/09/2026. Type each amount in column **B**.
3. In **C2**, type `=YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0)`.
4. Select **C2**, then drag the fill handle down to the last sale. A sale on 05/09/2026 now shows `2026 Q2`.
5. On a second tab, type the quarter labels you want in column A, for example `2026 Q2`.
6. Next to a label, type `=SUMIFS(Sales!$B$2:$B$200,Sales!$C$2:$C$200,A2)` and fill it down. It adds every amount whose quarter label matches.
7. To count orders as well, type `=COUNTIFS(Sales!$C$2:$C$200,A2)` in the next column.

`MONTH(A2)/3` gives 0.33 to 4. `ROUNDUP(...,0)` rounds that up, so months 1 to 3 become 1, months 4 to 6 become 2, and so on. The `&` joins the year, the text ` Q` and the number into one label.

The `$` signs lock the ranges at rows 2 to 200, so they do not shift when you fill down. Rows past 200 are not counted. Extend the ranges if you have more sales.

### Check the total without a helper column

The template also has a second method that skips column C. It adds amounts whose date falls between the first day of the quarter and the first day of the next one:

`=SUMIFS(Sales!$B$2:$B$200,Sales!$A$2:$A$200,">="&B2,Sales!$A$2:$A$200,"<"&EDATE(B2,3))`

Here B2 holds the first day of the quarter, such as 04/01/2026. `EDATE(B2,3)` is three months later, so the "less than" test stops at the last day of the quarter. A last column compares the two totals with `=IF(ROUND(C2-E2,2)=0,"Yes","CHECK")`.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) page.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [SUMIFS](https://support.google.com/docs/answer/3238496) help page. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is 16 sales from 12/20/2025 to 12/31/2026, totaled by quarter.

| Quarter | Orders | Total |
|---|---|---|
| 2025 Q4 | 1 | $90.00 |
| 2026 Q1 | 3 | $405.50 |
| 2026 Q2 | 4 | $630.25 |
| 2026 Q3 | 4 | $600.50 |
| 2026 Q4 | 4 | $779.99 |
| All quarters | 16 | $2,506.24 |

Check Q2 by hand: $150.00 plus $95.25 plus $310.00 plus $75.00 is $630.25. That covers sales from 04/01/2026 to 06/30/2026.

Boundary days land where they should. A sale on 03/31/2026 is in 2026 Q1, one on 04/01/2026 is in 2026 Q2, and 12/31/2026 is in 2026 Q4.

In the recalculated file, both methods gave the same total in all five quarters, and the sum of the quarters equaled the sum of the whole Sales tab at $2,506.24.

## Troubleshooting

### A quarter total is $0.00
The label you typed must match column C exactly. `2026 Q2` has a space before the Q, and `2026Q2` or `Q2 2026` will not match. Copy a label from column C instead of typing it.

### The two methods disagree and the check says CHECK
A date may be stored as text. In a test copy, a date typed as text in A8 still got a quarter label in LibreOffice, so the helper total stayed at $630.25. But the date-range total dropped to $320.25, because the date test needs a real date. Retype the date so it right-aligns, or convert the column to dates. How Excel and Google Sheets treat text dates was not tested.

### The quarters add up to less than the Sales tab
A sale may fall in a quarter you did not list, or its date may be blank or text. Compare the "All quarters" row with the "All sales on Sales tab" row. A gap points to a missing label.

### A sale from last year lands in this year's total
That happens if the label has no year, such as `Q4`. Keep `YEAR(A2)` in the label so 2025 Q4 and 2026 Q4 stay separate.

### New sales are not counted
The ranges stop at row 200. Change `200` to a larger row number in every formula, and fill column C down to the new rows.

## Template

[Download the .xlsx](templates/tidy-tabs-quarterly-sales-totals.xlsx)

The file has three tabs. **Sales** holds 16 sample sales in columns A and B and the quarter formula in column C. **By quarter** lists five quarters with the helper-column total, an order count, the date-range total and a Yes or CHECK flag, plus a row that adds all quarters and a row that adds the whole Sales tab. **How to use** has short fill-in notes. Each tab prints on one US Letter page wide.

The dates and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.

This is arithmetic on the numbers you type. It does not give tax advice, and tax-filing quarters can differ from calendar quarters.
