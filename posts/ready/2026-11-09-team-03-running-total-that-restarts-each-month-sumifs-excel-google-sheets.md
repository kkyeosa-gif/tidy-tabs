---
title: How to Make a Running Total That Restarts Each Month in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 17 workbooks and 203 sorted rows. Checked: the sample, an empty sheet, month and year boundaries (10/31, 11/01, 11/30, 12/31, 01/01, 02/29/2028, 03/01), blank rows in the middle, refunds, five rows on one day, times of day, four random sets up to 60 rows, an unsorted list, a limit of 0, a text or blank limit, a text date and a text amount. Excel and Google Sheets were not used. Not run: more than 60 rows.
search_description: Make a running total that starts over on the 1st of each month with SUMIFS in Excel and Google Sheets. Includes a month-to-date expense template.
image_prompts: A stack of paper receipts under a binder clip, a pencil and a desk calendar page on a wooden desk, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Paper receipts under a binder clip beside a pencil and a desk calendar
threads: A running total that never resets is useless for a monthly budget.\nIn a date-sorted list: =SUMIFS($C$4:C4,$A$4:A4,">="&DATE(YEAR(A4),MONTH(A4),1))\nIn the sample, the 11/18/2026 expense shows $407.25 for November, not $676.75 (made-up expenses).
---
In a date-sorted list, use `=SUMIFS($C$4:C4,$A$4:A4,">="&DATE(YEAR(A4),MONTH(A4),1))` so each row adds up only the amounts from the 1st of its own month down to that row. In the sample, the 11/18/2026 expense shows $407.25 for November, not the $676.75 a plain running total would show (made-up expenses).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (17 workbooks, 203 rows). The formulas use only IF, ISNUMBER, SUMIFS, DATE, YEAR, MONTH and EDATE, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Expense table with month to date, amount left and month total columns for eight sample expenses](images/2026-11-09-team-03-running-total-that-restarts-each-month-sumifs-excel-google-sheets/monthly-running-total-expenses-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. The template has one data tab, **Expenses**: monthly limit in **B1**, then date in **A**, what it was in **B** and amount in **C**, starting at row 4 under the headers in row 3.

1. Type your expenses with the oldest date first. Enter dates as MM/DD/YYYY, such as 11/09/2026.
2. Sort the list by date, oldest first, before you add the formula. Step 3 only adds rows above the current one, so the order matters.
3. In **D4**, type `=SUMIFS($C$4:C4,$A$4:A4,">="&DATE(YEAR(A4),MONTH(A4),1))`.
4. Fill **D4** down the column. The `$C$4` is locked, and the `C4` after the colon moves, so the range grows by one row each time.
5. In **E4**, type `=$B$1-D4` to see how much of the monthly limit is left. Put your limit, such as 450.00, in **B1**.
6. Optional: in **F4**, type `=SUMIFS($C$4:$C$63,$A$4:$A$63,">="&DATE(YEAR(A4),MONTH(A4),1),$A$4:$A$63,"<"&EDATE(DATE(YEAR(A4),MONTH(A4),1),1))`. This is the full month total on every row of that month, not just the part so far.

Here is how step 3 works. `DATE(YEAR(A4),MONTH(A4),1)` turns any date into the 1st of its month. The `">="&` part keeps only rows on or after that day. The growing range stops at the current row, so later dates never count.

The template wraps these in checks so a text date shows Check instead of a wrong number. The versions above are the core.

Rows are locked to a fixed range (row 63 in step 6), so the full-month formula stays short. If your list grows past row 63, extend every `$63` to the same last row.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349), and Google's [SUMIFS](https://support.google.com/docs/answer/3238496). They describe the functions, not our test.

## Example

Sample data is fictional: 8 expenses from 10/03/2026 to 12/04/2026, with a $450.00 monthly limit.

| Date | What | Amount | Month to date | Left | Month total |
|---|---|---|---|---|---|
| 10/03/2026 | Supplies | $120.00 | $120.00 | $330.00 | $269.50 |
| 10/12/2026 | Booth fee | $85.00 | $205.00 | $245.00 | $269.50 |
| 10/27/2026 | Shipping | $64.50 | $269.50 | $180.50 | $269.50 |
| 11/02/2026 | Supplies | $210.00 | $210.00 | $240.00 | $446.00 |
| 11/09/2026 | Packaging | $47.25 | $257.25 | $192.75 | $446.00 |
| 11/18/2026 | Ads | $150.00 | $407.25 | $42.75 | $446.00 |
| 11/30/2026 | Shipping | $38.75 | $446.00 | $4.00 | $446.00 |
| 12/04/2026 | Supplies | $95.00 | $95.00 | $355.00 | $95.00 |

The total of all amounts is $810.50.

Look at 11/02/2026. A plain running total would carry October's $269.50 forward. Here the sum starts over at $210.00. On 11/18/2026 a plain total would read $676.75 (120 + 85 + 64.50 + 210 + 47.25 + 150). The monthly version reads $407.25, which leaves $42.75 of the $450.00 limit.

## Troubleshooting

### Month to date is wrong after you add a row in the middle
The formula only adds rows above the current one. In a test with the dates 11/20, 10/05, 11/02, 12/03 and 11/25, the column showed 100, 150, 110, 1 and 116 when the true month-to-date values were 110, 50, 10, 1 and 115. The template flags the three rows dated before an earlier row with Out of order. Sort the dates oldest first.

### The month total is right but month to date is wrong
Same cause. The full-month formula in step 6 reads the whole range, so row order does not change it. Month to date does depend on order. Sort the list.

### An amount is ignored
`SUMIFS` skips an amount stored as text, such as "abc" or a number with a leading apostrophe, and the template shows Check on that row. Retype it as a number.

### A date shows Check
The date is stored as text, so the formula cannot find its month. Retype it as MM/DD/YYYY. A date with a time of day, such as 10/31/2026 11:30 PM, works.

### Left this month shows Check or goes negative
Check shows when the limit in **B1** is blank or text. A negative Left means the month is over the limit, and the template turns it red. A refund goes in as a negative amount and lowers the total.

## Template

[Download the .xlsx](templates/tidy-tabs-monthly-running-total.xlsx)

The file has two tabs. **Expenses** has 60 rows from row 4 with 8 fictional sample expenses, the columns **Month to date**, **Left this month**, **Month total** and **Row check**, and a side block that totals all amounts, counts rows showing Check or Out of order, counts rows over the limit, and has a **CHECK** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The list is set up for up to 60 rows.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
