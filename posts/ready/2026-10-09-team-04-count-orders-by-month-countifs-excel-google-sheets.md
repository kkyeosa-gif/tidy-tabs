---
title: How to Count Orders by Month With COUNTIFS in Excel and Google Sheets
labels: excel-formulas, order-tracking, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. Excel and Google Sheets use the same COUNTIFS, SUMIFS and EDATE functions; the help links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Count orders per month with COUNTIFS and EDATE, or total revenue with SUMIFS. Month-boundary example, text-date fixes and a free .xlsx template.
image_prompts: Home craft shop packing shelf with a row of four small kraft shipping boxes each tied with a different colored twine, a roll of washi tape, and a tray of folded paper order slips, soft window light, all paper text blurred and unreadable, no screens, no logos
image_alt: Four kraft shipping boxes tied with different colored twine beside a tray of order slips
threads: How many orders did you get in September?\nPut the first of the month in A2, then =COUNTIFS(Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1)).\nSwap in SUMIFS and the same two tests give the month's revenue instead.
---
Count the orders in a month with two date tests: `=COUNTIFS(Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`, where **A2** is the first day of the month. Swap `COUNTIFS` for `SUMIFS` and add the amount column to total the revenue.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against a hand count. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Monthly table of order counts and revenue for July to October with a totals row](images/2026-10-09-team-04-count-orders-by-month-countifs-excel-google-sheets/count-orders-by-month-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

You need an order list with real dates in one column and amounts in another, plus a small table with one row per month.

1. On a sheet named **Orders**, put the order date in column A, the customer in column B and the amount in column C. Enter dates as real dates, such as 08/31/2026.
2. On a second sheet named **Monthly**, type the first day of each month in column A: 07/01/2026, 08/01/2026, 09/01/2026, 10/01/2026.
3. In **B2**, type `=COUNTIFS(Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`.
4. In **C2**, type `=SUMIFS(Orders!$C$2:$C$200,Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`.
5. Select **B2:C2** and fill down for each month.
6. Format column A as a date, column C as currency with two decimals.
7. Under the last month, type `=SUM(B2:B5)` in **B6** for the total of the monthly counts.
8. As a check, type `=COUNT(Orders!A2:A200)` in **B7** and `=SUM(Orders!C2:C200)` in **C7**. These count every dated order without looking at months, so B6 and B7 should match.

The range stops at row 200 so you can keep adding orders without editing the formula. Extend it if your list will grow past that.

### How the two tests work

The first test, `">="&A2`, keeps orders on or after the first of the month. The second, `"<"&EDATE(A2,1)`, keeps orders before the first of the next month. `EDATE(A2,1)` moves a date forward by one month, so 08/01/2026 becomes 09/01/2026.

Using "before the first of next month" means you never have to know whether a month has 28, 30 or 31 days. It should also catch dates that carry a time of day, since 08/31/2026 6:00 PM is still before 09/01/2026. That time-of-day case was not tested here.

The `&` joins the comparison sign to the date in A2. Skip it and the formula compares against the text ">=A2" instead of the date.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) pages.

### In Google Sheets

The same formulas should work as typed. Google documents the arguments on its [SUMIFS function](https://support.google.com/docs/answer/3238496) page, and COUNTIFS takes its criteria the same way. Google Sheets was not tested for this post, so compare one month against a hand count before you rely on it.

## Example

The order list below is fictional. It has 10 orders, including one on 08/31/2026 and one on 09/01/2026 to check the month boundary.

| Month start | Orders | Revenue |
|---|---|---|
| 07/01/2026 | 2 | $365.50 |
| 08/01/2026 | 3 | $549.25 |
| 09/01/2026 | 3 | $693.75 |
| 10/01/2026 | 2 | $395.50 |
| Total | 10 | $2,004.00 |

The 08/31/2026 order is counted in August and the 09/01/2026 order in September. Neither is counted twice or dropped.

The monthly counts add up to 10 and the revenue to $2,004.00. The check cells on the **Monthly** tab, which skip the month tests, return the same 10 and $2,004.00. If those two ever disagree, some orders fall outside your month list or have text dates.

## Troubleshooting

### The count is 0 but I can see orders in that month
The dates are probably stored as text, which the date criteria skip without an error. Click a date cell. If it is left-aligned or `=ISNUMBER(A2)` returns FALSE, convert the column to real dates.

### The total of the months is lower than the order count
Some orders fall outside the months you listed, or their dates are text. Compare cell **B6** with the plain count in **B7**. The difference is the number of orders your month table missed.

### An order on the last day of the month lands in the next month
Check that **A2** holds the first day of the month, not the last. The formula counts from that day up to, but not including, the next first of the month.

### The formula returns an error or #NAME?
Check that the sheet name matches exactly. The formulas expect a tab called **Orders**. If you renamed it, update every reference. A sheet name with spaces needs single quotes, such as `'Order list'!$A$2:$A$200`.

### My month list shows numbers like 46204 instead of dates
The cell holds a date serial number. Format column A as a date.

## Template

[Download the .xlsx](templates/tidy-tabs-count-orders-by-month.xlsx)

The file has three tabs. **Orders** holds the 10 sample orders with date, customer and amount. **Monthly** holds the four month rows with the count and revenue formulas, a total row and the two check cells. **How to use** has short fill-in notes.

The orders and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608?hl=en).
