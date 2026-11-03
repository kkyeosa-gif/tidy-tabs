---
threads_url: https://www.threads.com/@tin_ylab/post/DfR9RDYG-qx
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-add-up-weekly-sales-from-daily.html
title: How to Add Up Weekly Sales From a Daily Sales Log in Excel and Google Sheets
labels: excel-formulas, small-business, sales-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 14 workbooks and 1,071 comparisons with 0 mismatches. Checked: an empty log, 1, 5, 40 and 200 sales, a first date that is not a Monday, a blank first date, text in an amount cell, and a sale typed with a time of day on a Sunday. Excel and Google Sheets were not used.
search_description: Add up weekly sales from a daily sales log with SUMIFS in Excel and Google Sheets. Free template with weeks from Monday to Sunday and a check cell.
image_prompts: A paper wall calendar with one week of boxes circled in pencil next to a stack of folded market receipts and a small cash tin on a wooden table, soft daylight, no screens, no readable text, no logos
image_alt: Wall calendar with one week circled in pencil beside folded receipts and a cash tin
threads: Daily sales log, but you want the week?\nSUMIFS with a start date and an end date does it: ">="&A6 and "<="&B6, where B6 is A6+6.\nAdd 7 for each row below and you have a week by week table. Type dates only, because a sale with a time on Sunday falls out of the week.
---
Add up one week of sales with `=SUMIFS(Sales!$C$5:$C$204,Sales!$A$5:$A$204,">="&A6,Sales!$A$5:$A$204,"<="&B6)`, where A6 is the Monday that starts the week and B6 is `=A6+6`. Add 7 days to A6 for each row below to get the next week.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. SUMIFS, COUNTIFS and WEEKDAY exist in current Excel and Google Sheets, but neither app was opened for this post.

![Weekly sales table with week starting, week ending, sales, orders and average order columns](images/2026-11-03-team-01-weekly-sales-totals-from-daily-log-sumifs-excel-google-sheets/weekly-sales-totals-by-week-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The log is on a tab named **Sales**: dates in column A, channel in B, amount in C, starting in row 5. The weekly table is on a tab named **Weeks**.

1. On **Weeks**, type the first Monday in **B2**, such as 09/07/2026. To find the Monday on or before a date, use `=A5-WEEKDAY(A5,2)+1`. The `2` makes Monday day 1.
2. In **A6**, type `=B2`. In **A7**, type `=A6+7` and fill it down. Each row is the next Monday.
3. In **B6**, type `=A6+6`. This is the Sunday that ends the week.
4. In **C6**, type the SUMIFS formula from the top of this post. It adds the amounts whose date is on or after the Monday and on or before the Sunday.
5. In **D6**, count the sales in the same week: `=COUNTIFS(Sales!$A$5:$A$204,">="&A6,Sales!$A$5:$A$204,"<="&B6)`.
6. Fill **A6:D6** down. Add `=IF(D6=0,"",ROUND(C6/D6,2))` in **E6** for the average order.

The dollar signs lock the log range so it does not slide as you fill down. The `&` joins the comparison sign to the date in the cell.

Type dates only, with no time of day. A date is a whole number and a time is a fraction. A sale stamped 09/13/2026 2:30 PM is larger than the Sunday date 09/13/2026, so `"<="&B6` leaves it out. In the test, a $100.00 sale stamped that way was missing from the week, and the check cell flagged it.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [WEEKDAY function](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/02/2026; the pages were not read for this post.

## Example

Sample data is fictional: 24 sales over five weeks from an Etsy shop, a Shopify store and a farmers market.

| Week starting | Week ending | Sales | Orders | Average order |
|---|---|---|---|---|
| 09/07/2026 | 09/13/2026 | $453.75 | 5 | $90.75 |
| 09/14/2026 | 09/20/2026 | $414.25 | 5 | $82.85 |
| 09/21/2026 | 09/27/2026 | $404.50 | 4 | $101.13 |
| 09/28/2026 | 10/04/2026 | $456.00 | 5 | $91.20 |
| 10/05/2026 | 10/11/2026 | $433.75 | 5 | $86.75 |

Together the weeks hold $2,162.25, which equals the total of the log. The template lists 16 weeks, so the rows after 10/11/2026 show $0.00 and 0 orders until you log more sales. The best week starts 09/28/2026.

## Troubleshooting

### A week shows $0.00 but you logged sales
Check the dates in column A of the log. If they are text, such as `9/7/26` typed into a Plain text column, the comparison does not see them as dates. Retype them or format the column as a date.

### A sale is missing from its week
It probably has a time of day. Remove the time and keep the date only. The check cell, **H9**, shows Check because **H7** (sales outside the listed weeks) is above $0.00.

### Every week is blank
The first Monday in **B2** is empty or is not a Monday. The template shows blanks instead of wrong weeks and the check cell says Check. Type a Monday, or copy the date from **B3**, which shows the Monday on or before your earliest sale.

### The sales total is smaller than the log total
Some sales fall before the first week or after the 16th. **H7** shows how much. Move **B2** earlier or extend the table by filling the last row down.

### An amount is text
A value like `$12.00` typed as text is not added by SUMIFS. In the test, a text amount left the sheet at Check. Retype the number or convert it.

## Template

[Download the .xlsx](templates/tidy-tabs-weekly-sales-totals.xlsx)

The file has three tabs. **Sales** has room for 200 sales in yellow cells with 24 fictional rows. **Weeks** has the first-Monday cell, a table of 16 weeks, a total of all sales, a total in the weeks listed, sales outside the weeks, the best week and a **CHECK** cell, **H9**. **How to use** has notes on each formula.

**CHECK** says OK when **B2** is a Monday, nothing falls outside the listed weeks and every date has an amount. Fix a Check before you trust a weekly total. Sales past row 204 of the log are not counted.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
