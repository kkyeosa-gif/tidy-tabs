---
threads_url: https://www.threads.com/@tin_ylab/post/DekBnSCllV7
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-week-ending-date-for-any.html
title: How to Find the Week Ending Date for Any Date in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python weekday math, including a date already on the week ending day, a year boundary, a Sunday week end, a date typed as text, hours typed as text and a blank date row. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google.
search_description: Week ending date: =A2+MOD(6-WEEKDAY(A2,1),7) turns 10/14/2026 into Friday 10/16/2026. Then sum hours by week. Free .xlsx template.
image_prompts: Wall calendar with a Friday square circled in marker, a spiral timesheet pad and a pencil on a wooden desk, soft window light, no screens, no readable text, no logos
image_alt: Circled Friday on a wall calendar beside a spiral pad and pencil
threads: Find the Friday that ends any date's week: =A2+MOD(6-WEEKDAY(A2,1),7).\nWednesday 10/14/2026 becomes 10/16/2026, and a date already on Friday stays put.\nThe sample log's 12 entries total 47.00 hours across three week ending dates.\nA date typed as text leaves the cell blank.
---
Find the Friday that ends a date's week with `=A2+MOD(6-WEEKDAY(A2,1),7)`. A Wednesday such as 10/14/2026 returns 10/16/2026, and a Friday stays the same.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Work log table with a Friday week ending date beside every entry and total hours](images/2026-10-16-team-01-week-ending-friday-date-weekday-excel-google-sheets/week-ending-date-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

`WEEKDAY(A2,1)` returns 1 for Sunday through 7 for Saturday. `MOD` counts how many days are left until the week ending day, so the formula adds just that many days to the date.

1. On a tab named **Log**, type these headers in row 1: Date, Client, Hours, Week ending.
2. Type the work date as MM/DD/YYYY in column **A**, the client in column **B** and the hours as a plain number in column **C**.
3. In **G2**, type `6`. This is the weekday number the week ends on: 1 Sunday, 2 Monday, up to 7 Saturday. 6 is Friday.
4. In **D2**, type `=IF(ISNUMBER(A2),A2+MOD($G$2-WEEKDAY(A2,1),7),"")`. The `ISNUMBER` test leaves the cell blank when the date is empty or is text.
5. Format column **D** as a date (MM/DD/YYYY) and copy **D2** down to the last row of your log.
6. To total hours by week, list each week ending date in a second tab and use `=SUMIFS(Log!C2:C60,Log!D2:D60,A2)` beside it.

If you always want Friday and never plan to change it, replace `$G$2` with `6`. Keeping it in a cell lets you switch to a Sunday week end by typing `1`.

### In Excel

The formulas work as typed. Microsoft describes them on its [WEEKDAY function](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a) and [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) pages. Set the date format from **Home > Number Format > Short Date**, or use **Format Cells > Custom** with `mm/dd/yyyy`.

### In Google Sheets

The same formulas should work as typed. Google describes the [WEEKDAY function](https://support.google.com/docs/answer/3092985) on its help page. Set the date format with **Format > Number > Date**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a freelancer's hours log for three weeks, with weeks ending on Friday (G2 = 6).

| Date | Client | Hours | Week ending (D) |
|---|---|---|---|
| 09/28/2026 | Maple Street Bakery | 3.50 | 10/02/2026 |
| 09/29/2026 | Harbor Yoga Studio | 4.00 | 10/02/2026 |
| 09/30/2026 | Maple Street Bakery | 2.50 | 10/02/2026 |
| 10/01/2026 | Birch Design Co | 6.00 | 10/02/2026 |
| 10/02/2026 | Birch Design Co | 5.50 | 10/02/2026 |
| 10/03/2026 | Harbor Yoga Studio | 2.00 | 10/09/2026 |
| 10/04/2026 | Lantern Print Shop | 1.50 | 10/09/2026 |
| 10/05/2026 | Lantern Print Shop | 7.00 | 10/09/2026 |
| 10/07/2026 | Maple Street Bakery | 4.25 | 10/09/2026 |
| 10/09/2026 | Birch Design Co | 3.00 | 10/09/2026 |
| 10/12/2026 | Harbor Yoga Studio | 5.00 | 10/16/2026 |
| 10/14/2026 | Lantern Print Shop | 2.75 | 10/16/2026 |

Summing the hours for each week ending date gives this.

| Week ending | Hours |
|---|---|
| 10/02/2026 | 21.50 |
| 10/09/2026 | 17.75 |
| 10/16/2026 | 7.75 |

The total is 47.00 hours, the same as the log. A date already on a Friday (10/02, 10/09) stays put. The Saturday right after (10/03) jumps six days to 10/09.

## Troubleshooting

### The week ending cell is blank
`ISNUMBER` returned FALSE, so the date is probably text. In the recalculated copy, a date typed as text ("10/20/2026") left column D blank and its hours out of the weekly totals, so the weeks tab showed 47.00 against 50.00 in the log. Left-aligned dates are the clue. Retype the date.

### The weekly total is missing some hours
Hours typed as text are skipped without any warning. In the recalculated copy, hours typed as text ("2.5") dropped out of both the log total and the weekly sums, so both showed 44.50 and the match check still said OK. Compare the log total with the number of rows you filled in, and retype any left-aligned hours as a plain number.

### The date landed on the wrong day
Check the number in **G2**. A week ending Sunday needs `1`, not `7`. With G2 set to 1, the recalculated copy gave 10/04/2026 at 25.00 hours, 10/11/2026 at 14.25 and 10/18/2026 at 7.75. A Sunday such as 10/18/2026 moves to 10/23/2026 with G2 at 6, because it is the first day of the next Friday week.

### The weekly list shows empty weeks
A week with no hours between two busy weeks still gets a row, with 0.00 hours. This happened across a year boundary in the recalculated copy: 12/30/2026, a Wednesday, returned 01/01/2027, and the weeks in between showed no hours. That is correct, and you can ignore those rows.

## Limits

This is date math only. It does not know your pay period, overtime rules or time zone. The template lists up to 26 weeks on the weeks tab and 59 log rows, so lengthen the ranges for a bigger log.

## Template

[Download the .xlsx](templates/tidy-tabs-week-ending-date.xlsx)

The file has three tabs. **Log** holds the 12 sample entries in yellow input cells, room down to row 60, the week ending weekday in **G2** and the weekday name in **G3**. **Weeks** lists each week ending date with its hours and a check that the totals match. **How to use** has short fill-in notes. All three tabs print portrait on US Letter, though the Log tab runs onto more than one page when filled.

The entries are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
