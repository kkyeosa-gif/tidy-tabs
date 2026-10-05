---
title: How to Calculate Percent Change in Excel and Google Sheets
labels: excel-formulas, etsy-sellers, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same functions; the IF link is from Microsoft Support (not tested here).
search_description: Calculate month-over-month percent change with =IF(B2=0,"",(C2-B2)/B2) formatted as Percent. Free .xlsx template with sample shop sales.
image_prompts: A wooden shipping table with a roll of kraft packing paper, a tape gun and two small cardboard mailer boxes tied with twine, a paper notepad with blurred unreadable pencil columns, soft window light, no screens, no logos
image_alt: Packing table with kraft paper roll, tape gun, twine-tied mailer boxes and notepad
threads: $2,400.00 to $2,760.00 is 15.0%, not 13.0%.\nDivide the change by the earlier month: =IF(B2=0,"",(C2-B2)/B2).\nA last month of $0.00 leaves the percent cell blank.
---
Subtract last month from this month and divide by last month: `=IF(B2=0,"",(C2-B2)/B2)`, then format the cell as Percent. Going from $2,400.00 to $2,760.00 shows 15.0%.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions, and the formulas are plain arithmetic and IF. Neither app was opened for this post.

![Sales sheet comparing last month and this month for five sales channels with change in dollars and percent](images/2026-10-06-team-02-percent-change-month-over-month-excel-google-sheets/percent-change-template.png) *LibreOffice Calc 24.2 PDF export of the Sales sheet with sample data (fake).*

## Steps

Use one row per thing you want to compare, such as a sales channel or a product. Last month goes in column B and this month in column C.

1. In row 1, type these headers in **A1:E1**: Sales channel, Last month, This month, Change ($), Change (%).
2. Type the name in **A2**, last month's number in **B2**, and this month's number in **C2**. Enter plain numbers with no text.
3. In **D2** (Change in dollars), type `=C2-B2`.
4. In **E2** (Change in percent), type `=IF(B2=0,"",(C2-B2)/B2)`.
5. Select **E2:E6** and apply the Percent format with one decimal place. The template uses the format `0.0%`.
6. Select **D2** and **E2**, then drag the fill handle down to cover every row.
7. Format B, C, and D as currency, such as `$#,##0.00`.

The formula divides by the earlier month, which is column B. A negative percent means the number went down.

The IF handles a last month of 0. You cannot divide by zero, and a percent change from nothing is undefined. With the IF, the cell stays blank instead of showing an error.

### In Excel

The formulas work as typed. Apply the percent format from the Home tab, in the Number group. Microsoft documents the logic on its [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) page.

### In Google Sheets

The same formulas should work as typed, because they use only subtraction, division, and IF. The percent format is in the toolbar and under the **Format** menu. No Google help page is linked here because none could be verified for this post.

## Example

Sample data below is fictional. It is a small shop's month-to-month sales by channel.

| Sales channel | Last month | This month | Change ($) | Change (%) |
|---|---|---|---|---|
| Whole shop | $2,400.00 | $2,760.00 | $360.00 | 15.0% |
| Etsy orders | $1,800.00 | $1,692.00 | -$108.00 | -6.0% |
| Farmers market | $950.00 | $1,140.00 | $190.00 | 20.0% |
| Workshops (new this month) | $0.00 | $300.00 | $300.00 | (blank) |
| Wholesale | $640.00 | $640.00 | $0.00 | 0.0% |

Check the first row by hand: $2,760.00 minus $2,400.00 is $360.00, and $360.00 divided by $2,400.00 is 0.15, or 15.0%.

Etsy orders went down by $108.00. That is -6.0% of $1,800.00, so the percent is negative.

Workshops had no sales last month. The dollar change is still $300.00, but the percent cell is blank. That blank is the IF doing its job.

Wholesale did not move, so both columns show zero. The percent shows 0.0%, not blank, because last month was not zero.

## Troubleshooting

### The percent is smaller than it should be
You divided by the later month. For the whole shop row, `=(C2-B2)/C2` gives $360.00 divided by $2,760.00, which is about 13.0%, not 15.0%. Divide by column B, the earlier month, as in the formula above.

### The percent cell is blank
Check last month in column B. If it is 0, the IF returns an empty cell on purpose, as in the Workshops row. If B holds a number, make sure it is not text.

### The cell shows 0.15 instead of 15.0%
The cell is formatted as a number. Apply the Percent format. The formula already returns the fraction, so do not multiply it by 100 yourself. This was not reproduced for this post.

### A divide-by-zero error shows up
You are probably using `=(C2-B2)/B2` without the IF, and a last month of 0 breaks it. Wrap it as `=IF(B2=0,"",(C2-B2)/B2)`. The error was not reproduced for this post. The template always uses the IF version.

## Template

[Download the .xlsx](templates/tidy-tabs-percent-change.xlsx)

The file has two tabs. **Sales** holds the five sample rows with the Change ($) and Change (%) formulas, already formatted as currency and percent. **How to use** has short fill-in notes, including the divide-by-the-earlier-month rule and what a blank percent means.

The shop and numbers are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
