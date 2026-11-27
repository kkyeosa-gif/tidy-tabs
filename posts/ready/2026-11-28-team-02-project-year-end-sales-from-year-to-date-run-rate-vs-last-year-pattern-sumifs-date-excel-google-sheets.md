---
title: How to Project Year-End Sales From Year-to-Date Sales in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 227 workbooks and 75,750 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Project year-end sales from year to date: divide sales so far by last year's share made by the same date, next to the straight-line run rate.
image_prompts: A small shop's back office desk in late November with two bound ledger books open side by side, one older with worn corners and one newer, a desktop adding machine trailing a long curl of paper tape, and a sharpened pencil resting in the gutter of the newer book, warm lamp light, realistic photo, no screens, no readable text, no logos
image_alt: Worn ledger book open beside a newer one, an adding machine and curled paper tape
threads: $39,299.04 sold through 11/25, and the straight-line run rate says the year ends at $43,599.24.\nLast year only 79.79% of sales were in by 11/25, so =ROUND(B5*B8/B7,2) projects $49,252.37 instead.\nSame sales, two very different Decembers.
---
Divide this year's sales so far by the share of last year's sales that came in by the same date: `=ROUND(B5*B8/B7,2)`, where B5 is year to date, B7 last year to the same date and B8 last year's full total. In the made-up gift shop, $39,299.04 through 11/25/2026 projects to $49,252.37, because last year only 79.79% of sales had come in by 11/25.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (227 workbooks, 75,750 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Two year-end projections side by side: last year's pattern versus a straight-line run rate](images/2026-11-28-team-02-project-year-end-sales-from-year-to-date-run-rate-vs-last-year-pattern-sumifs-date-excel-google-sheets/year-end-sales-projection-template.png) *Projection tab of a made-up gift shop as of 11/25/2026: $39,299.04 sold so far against $35,696.16 by the same date last year, a last-year-pattern projection of $49,252.37 ($1,252.37 over the $48,000.00 goal, green) next to a straight-line run rate of $43,599.24 (-$4,400.76, orange), and $8,700.96 still needed over the last 36 days. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The quick method is the straight-line run rate: sales so far, divided by the days so far, times the days in the year. It treats December like an average month, so a shop whose busiest month is still ahead gets a number that is too low. The template shows both.

The steps are the same in Excel and Google Sheets.

1. On the **Sales log** tab, type one row per sale or per day from row 5: Date in **A**, Amount in **B**, an optional Note in **C**, from 01/01 of last year onward.
2. Check column **D** on Sales log: OK, a Check message, or a "not counted" note for rows before last year or after B1.
3. On the **Projection** tab, type the last day of sales to count in **B1**. The sheet does not use TODAY.
4. Type your annual goal in **B2**, or leave it blank to skip the goal lines. Leave **B3** blank to compare with the same month and day last year.
5. **B4** finds that date with `DATE(YEAR(B1)-1,MONTH(B1),DAY(B1))`, or uses B3 if you typed one. See the DATE help from [Microsoft](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) or [Google](https://support.google.com/docs/answer/3092969).
6. **B5** totals this year to date with SUMIFS on the OK rows from 01/01 through B1: `=IF($B$27<>"OK","",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$D$5:$D$5004,"OK",'Sales log'!$A$5:$A$5004,">="&B6,'Sales log'!$A$5:$A$5004,"<"&B1+1))`. The `"<"&B1+1` part means a sale with a time of day on B1 still counts. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
7. **B7** uses the same SUMIFS for 01/01 of last year through B4, and **B8** for all of last year. **B9** shows the share, `=IF(OR(B8="",N(B8)<=0),"",ROUND(B7/B8,4))`, for reading only.
8. Read last year's pattern in **B15**: `=ROUND(B5*B8/B7,2)`, wrapped in checks that say "No sales last year" or "No sales last year by this date" instead of an error.
9. Read the straight-line run rate in **C15**: `=IF(B6="","",ROUND(B5/B10*B11,2))`, where **B10** is the days so far (`=B1-B6+1`) and **B11** the days in the year (365, or 366 in a leap year).
10. Rows **16 to 18** compare each projection with the goal. Rows **21 to 24** show what is still needed, per remaining day, beside what the same stretch of last year sold.

**CHECK** in **B32** is OK only when the settings are valid, the log covers last year from its first week, the Sales log CHECK is OK and the pattern projection is a number.

The pattern projection assumes this year keeps last year's seasonal shape. A moving holiday shifts the share by a few days: Black Friday was 11/28/2025 and 11/27/2026. To line up those days, see [How to Compare Black Friday Sales to Last Year](https://tidytabs.blogspot.com/2026/11/how-to-compare-black-friday-sales-to.html). Both numbers are estimates from your own past sales, not a promise, and the sheet gives no budget or tax advice.

## Example

The sales below are made up: a gift shop's log of 595 rows from 01/01/2025 to 11/25/2026, closed Mondays, with December the busiest month. As of 11/25/2026 with a $48,000.00 goal:

| Line | Last year's pattern | Straight-line run rate |
|---|---|---|
| Projected total for 2026 | $49,252.37 | $43,599.24 |
| Ahead (+) or short (-) of $48,000.00 | $1,252.37 | -$4,400.76 |
| Expected for the rest of the year | $9,953.33 | $4,300.20 |
| Expected per remaining day | $276.48 | $119.45 |

Sales this year to date are $39,299.04. Last year brought in $35,696.16 by 11/25/2025 out of $44,736.98 (79.79%). The year is 329 of 365 days in (90.14%).

The run rate assumes 90.14% of the year means 90.14% of the sales and calls a $4,400.76 miss. Last year only 79.79% had come in by then, so the pattern expects a $1,252.37 beat.

To reach the goal, the shop needs $8,700.96 in the last 36 days, or $241.69 per day. The same stretch of 2025 sold $9,040.82.

For the year-to-date total on its own, see [How to Calculate Year-to-Date Sales](https://tidytabs.blogspot.com/2026/10/how-to-calculate-year-to-date-sales-in.html). To see which months carry the year, see [How to Find Your Busiest and Slowest Months With a Seasonal Index](https://tidytabs.blogspot.com/2026/10/how-to-find-your-busiest-and-slowest.html).

## Troubleshooting

These cases were reproduced in LibreOffice Calc 24.2.7.2 with the template.

### The projection says "No sales last year by this date"
The log has nothing before the same date last year. With a log that started 06/01/2025 and an as-of date of 03/01/2026, **B15** showed this message, the run rate still showed, and CHECK said Check.

### CHECK is red and B29 says "Check: log starts after 01/07 of last year"
The log does not reach back to the start of last year. A log that starts later makes last year's share too high and the projection too low. A log starting 01/07/2025 passed; 01/08/2025 and 05/01/2025 did not.

### Your own version is a little off from the template
You may have divided by the rounded share in **B9**. In the sample, dividing by 79.79% instead of B7/B8 gave $49,253.09, $0.72 off from $49,252.37. Use B7 and B8 in the formula and keep B9 for reading.

### B30 says "Moved to 03/01 (no 02/29 last year)"
The as-of date is 02/29. In LibreOffice, `DATE(YEAR(B1)-1,MONTH(B1),DAY(B1))` for 02/29/2028 gave 03/01/2027. To compare with 02/28 instead, type 02/28/2027 in **B3**, and B30 says "Typed in B3".

## Template

[Download the .xlsx](templates/tidy-tabs-year-end-sales-projection.xlsx)

**Projection** holds the inputs in **B1:B3**, the two projections side by side in rows 14 to 18, the still-needed block in rows 20 to 24 and the checks in rows 26 to 32. **Sales log** takes 5,000 rows (rows 5 to 5004) with a summary and its own CHECK in **F4:G9**, which also catches rows typed below row 5004. **How to use** explains each formula. Every tab prints on US Letter.

Yellow cells are inputs and gray cells are formulas. The pattern projection is light green, rows before last year or after the as-of date are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
