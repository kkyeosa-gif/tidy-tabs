---
threads_url: https://www.threads.com/@tin_ylab/post/DeU5cNimkBs
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-rolling-average-of.html
title: How to Calculate a Rolling Average of Monthly Sales in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. The template was recalculated there and every result was checked against hand math. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Rolling 3 month average of monthly sales with =IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4)), filled down, plus a check column. Free .xlsx template.
image_prompts: Craft studio shelf with a stack of twelve monthly paper sales envelopes in a row, a desk calendar page and a small cash tin, soft window light, no screens, no readable text, no logos
image_alt: Row of monthly envelopes beside a wall calendar and a small cash tin
threads: Rolling 3 month average, one formula: =IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4)) in row 4, filled down.\nThe first two months stay blank instead of showing a fake average.\nA =SUM(B2:B4)/3 column next to it checks the math.
---
In the row for the third month, type `=IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4))` and fill it down. Each cell then averages that month and the two before it, and the first two months stay blank.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to their help pages. Neither Excel nor Google Sheets was opened for this post.

![Monthly sales table with a rolling 3 month average column starting in the third month](images/2026-10-10-team-02-rolling-3-month-average-sales-excel-google-sheets/rolling-average-sales-template.png) *PDF export from LibreOffice Calc 24.2, sample data fictional*

## Steps

A rolling average smooths out busy and slow months so you can see the trend. It needs one row per month, in order, with no gaps.

1. In row 1, type these headers in A to C: Month, Sales, 3 month average.
2. Type the months in **A2:A13**, oldest first, for example 01/2026 down to 12/2026.
3. Type each month's sales in **B2:B13** as plain numbers.
4. Leave **C2** and **C3** empty. Two months of data are not enough for a 3 month average.
5. In **C4**, type `=IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4))`.
6. Select **C4**, then drag the fill handle down to row 13. Each row now looks at its own month and the two above it.
7. In **D4**, type `=SUM(B2:B4)/3` and fill it down to row 13. This is a check column. It should equal column C.
8. Select columns B to D and format them as currency with two decimals.

`AVERAGE(B2:B4)` adds the three cells and divides by 3. `COUNT(B2:B4)` counts how many of them hold a number. The `IF` shows a blank instead of an average when fewer than three numbers are there.

The ranges are relative, so there are no `$` signs. When you fill down, `B2:B4` becomes `B3:B5`, then `B4:B6`, and the window slides one month at a time.

To add a new month, type it in row 14, then copy the last C and D cells down one row.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [AVERAGE function](https://support.microsoft.com/en-us/office/average-function-047bac88-d466-426c-a32b-8f33eb960cf6) page and its [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) page.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [AVERAGE](https://support.google.com/docs/answer/3093615) help page. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is twelve months of sales for a small shop.

| Month | Sales | 3 month average |
|---|---|---|
| 01/2026 | $1,200.00 | |
| 02/2026 | $950.00 | |
| 03/2026 | $1,100.00 | $1,083.33 |
| 04/2026 | $2,400.00 | $1,483.33 |
| 05/2026 | $3,100.00 | $2,200.00 |
| 06/2026 | $1,800.00 | $2,433.33 |
| 07/2026 | $1,650.00 | $2,183.33 |
| 08/2026 | $1,400.00 | $1,616.67 |
| 09/2026 | $2,250.00 | $1,766.67 |
| 10/2026 | $2,900.00 | $2,183.33 |
| 11/2026 | $3,800.00 | $2,983.33 |
| 12/2026 | $4,200.00 | $3,633.33 |

Check March by hand: $1,200.00 plus $950.00 plus $1,100.00 is $3,250.00, and divided by 3 that is $1,083.33. April is $950.00 plus $1,100.00 plus $2,400.00, which is $4,450.00, so $1,483.33.

In the recalculated file, all ten average cells matched an independent calculation in Python, and column D equaled column C in every row.

Notice how June drops to $1,800.00 from $3,100.00, yet the average only moves from $2,200.00 to $2,433.33. That is the point of the smoothing. One slow month does not flip the trend.

## Troubleshooting

### The first two months show a number, not a blank
You probably filled a formula into C2 and C3. A formula in row 2 would look at cells above the data, and the `IF` test is meant to start in row 4. Delete C2 and C3 so they are truly empty.

### A month in the middle shows a blank
`COUNT` counts only numbers. If a sales figure is typed as text, or the cell is empty, the window has fewer than 3 numbers and the formula returns a blank. Retype the value as a plain number such as 1100, with no quote mark or letters.

### The average looks too low after a missing month
The formula averages three rows, not three calendar months. If you skip a month, the window quietly reaches back one month further. Add a row for every month and type 0 for a month with no sales, if that is the true figure.

### Column D does not match column C
Check that C and D both start in row 4 and that the ranges are three cells tall. A range such as `B2:B5` covers four months, so the average will be off.

### The newest months stay blank
The formulas stop where you stopped filling. Select the last C and D cells and drag the fill handle down to the new month.

## Template

[Download the .xlsx](templates/tidy-tabs-rolling-average-sales.xlsx)

The file has two tabs. **Monthly sales** holds twelve sample months from 01/2026 to 12/2026 in columns A and B, the rolling average in column C, a check column in D and a Yes or CHECK flag in E. Rows 2 and 3 of columns C and D have no formula on purpose. **How to use** has short fill-in notes. The template has no chart.

The months and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
