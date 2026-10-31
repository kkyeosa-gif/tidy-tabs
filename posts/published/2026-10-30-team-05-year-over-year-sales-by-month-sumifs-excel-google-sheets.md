---
threads_url: https://www.threads.com/@tin_ylab/post/DfI0zaVDj2J
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-compare-this-years-sales-to-last.html
title: How to Compare This Year's Sales to Last Year Month by Month in Excel and Google Sheets
labels: excel-formulas, small-business, sales-reports
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model (23 workbooks, 1,776 cell comparisons). Checked: years 2024, 2025 and 2027, month boundaries with times, leap day 02/29/2024, a refund as a negative amount, a month with no sales last year, a text date and an invalid year. Excel and Google Sheets were not used.
search_description: Compare this year's sales to last year by month in Excel and Google Sheets with SUMIFS and EDATE. Free template with dollar and percent change.
image_prompts: Two paper ledgers side by side with a pencil across them and a wall calendar blurred behind on a craft shop desk, soft daylight, no screens, no readable text, no logos
image_alt: Two paper ledgers side by side with a pencil across them in a craft shop
threads: Want this year next to last year, month by month?\nSUMIFS with a start date and EDATE adds up each month without a pivot table.\nFor the percent, =IF(D4=0,"n/a",C4/D4-1) shows n/a instead of #DIV/0! when last year had no sales that month.
---
Add up one month with `=SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&B4,Orders!$A:$A,"<"&EDATE(B4,1))`, where B4 is the first day of the month. Show the change from the same month last year with `=IF(D4=0,"n/a",C4/D4-1)`, where C4 is this year and D4 is last year.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same formulas as typed. Neither Excel nor Google Sheets was opened for this post.

![Twelve month table of this year's sales beside last year's with change columns](images/2026-10-30-team-05-year-over-year-sales-by-month-sumifs-excel-google-sheets/year-over-year-sales-template.png) *caption: PDF render of the template with fictional sample data, made from LibreOffice Calc 24.2.7.2*

## Steps

You need a list of orders with a real date in one column and an amount in the next. The template keeps them in **Orders!A2:B500**, with dates in column A and amounts in column B. The comparison table is on a second tab.

### In Excel

1. On the comparison tab, type the year you want to look at in **B1**, for example `2026`. In **B2**, type `=B1-1` for last year.
2. In **B4**, type `=DATE($B$1,ROWS($A$4:A4),1)`. Fill it down to **B15**. `ROWS` counts 1, 2, 3 as you go down, so each row is the first day of the next month.
3. In **A4**, type `=TEXT(B4,"mmmm")` for the month name. Fill down.
4. In **C4**, type `=SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&B4,Orders!$A:$A,"<"&EDATE(B4,1))`. This adds every amount dated on or after the first of the month and before the first of the next. Fill down.
5. In **D4**, type `=SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&EDATE(B4,-12),Orders!$A:$A,"<"&EDATE(B4,-11))`. It is the same month one year earlier. Fill down.
6. In **E4**, type `=ROUND(C4-D4,2)` for the change in dollars. Fill down.
7. In **F4**, type `=IF(D4=0,"n/a",C4/D4-1)` and format the column as **Percentage**. Fill down.
8. In **C16**, type `=SUM(C4:C15)`, and copy it across to **D16**.

Microsoft's pages: [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [EDATE function](https://support.microsoft.com/en-us/excel/functions/edate-function). They describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 8 above, typed exactly as written.

Google's help page: [SUMIFS function](https://support.google.com/docs/answer/3238496). It describes the function, not our test. We did not run the whole-column references (`Orders!$B:$B`) in Sheets. If a range misbehaves there, replace it with a fixed one such as `Orders!$B$2:$B$500`.

No pivot table is needed, and none of these formulas is an array formula.

## Example

Sample data is fictional: 120 made-up orders from an online shop, from 01/01/2025 to 09/29/2026. The year in B1 is 2026.

| Month | This year | Last year | Change ($) | Change (%) |
|---|---|---|---|---|
| January | $353.02 | $171.32 | $181.70 | 106.1% |
| February | $222.48 | $76.39 | $146.09 | 191.2% |
| March | $480.83 | $0.00 | $480.83 | n/a |
| April | $294.09 | $312.40 | -$18.31 | -5.9% |
| May | $391.20 | $276.42 | $114.78 | 41.5% |
| June | $275.80 | $289.21 | -$13.41 | -4.6% |
| July | $376.10 | $140.21 | $235.89 | 168.2% |
| August | $421.72 | $234.86 | $186.86 | 79.6% |
| September | $354.26 | $286.95 | $67.31 | 23.5% |

March shows n/a because the sample has no orders in March 2025. Dividing by $0.00 would give #DIV/0!, so the `IF` shows n/a instead.

October to December are blank for this year. Blank means the month starts after the latest order (09/29/2026), not that sales were $0.00. The latest month shown can also be partial.

The totals row reads $3,169.50 for this year against $2,785.87 for last year, up $383.63 or 13.8%. That compares nine months with twelve. The template's row 17 adds last year's same nine months only: $1,787.76, so this year is up $1,381.74, or 77.3%. Read that row for a fair comparison.

The sample includes one refund of -$24.00 on 05/31/2026, typed as a negative amount. It reduces May like any other order.

## Troubleshooting

### A month shows n/a in the percent column
Last year had $0.00 for that month, so there is no base to divide by. This is intended. The dollar change still shows.

### Late months are blank instead of $0.00
In the template, a month that starts after the latest order date stays blank, so an unfinished year does not look like a bad year. Add orders for those months and they fill in.

### A percent looks extreme
A small base makes a large percent. In the test, a May with a net of -$20 against $20 last year showed -200%. The sheet shows the number and does not judge it.

### The check cell E2 says CHECK
Two things trigger it in the template: an order with a date typed as text (it is left out of every sum and counted in **C21**), or a year in **B1** that is not a whole number. Retype the date as a real date, such as 03/15/2026.

### Orders on the last day of a month land in the wrong month
They should not. In the test, 01/31/2026 at 11:59:59 PM stayed in January and 02/01/2026 at 12:00 AM went to February, because the formula stops before the first of the next month.

## Template

[Download the .xlsx](templates/tidy-tabs-year-over-year-sales.xlsx)

The file has three tabs. **Orders** holds yellow date and amount cells in **A2:B500**, with 120 fictional rows filled in. **Year over year** has the yellow year cell **B1** and gray formulas for the 12 months, the totals row, the same-months row, the best and weakest month (**C19**, **C20**), and the check in **E2**. **How to use** has short notes.

The template does not say what a change means for your business. It was recalculated only in LibreOffice Calc 24.2.7.2. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
