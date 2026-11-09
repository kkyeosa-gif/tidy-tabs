---
threads_url: https://www.threads.com/@tin_ylab/post/Dfh7kS8IyLZ
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-get-fiscal-year-and-quarter-from.html
title: How to Get the Fiscal Year and Quarter From a Date in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 146 workbooks and 13,177 date rows with 0 mismatches. Checked: every day from 01/01/2026 to 12/31/2028 (including 02/29/2028) for each start month 1 to 12, dates with a time of day, a date typed as text, blank dates, start month 0, 13, 7.5, text and blank, and a text amount. Excel and Google Sheets were not used. Not run: more than 100 sales rows, a non-US language setting.
search_description: Get the fiscal year and quarter from any date with a custom start month using YEAR, MONTH and MOD in Excel and Google Sheets. Free template.
image_prompts: A paper wall calendar with a few days circled in pencil, a brass paperclip and a small stack of kraft envelopes on a wooden desk, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Wall calendar with circled days beside a paperclip and kraft envelopes on a desk
threads: Your year starts in July, but every date formula assumes January?\nPut the start month in B1, then =INT(MOD(MONTH(A4)-$B$1,12)/3)+1 gives the quarter.\nIn the sample, 11/09/2026 is FY2027 Q2 (made-up sales).
---
Put the first month of your year in **B1** (7 for July), then get the fiscal year with `=YEAR(A4)+IF(AND($B$1>1,MONTH(A4)>=$B$1),1,0)` and the quarter with `=INT(MOD(MONTH(A4)-$B$1,12)/3)+1`. With a July start, 11/09/2026 is FY2027 Q2.

The year is named for the calendar year it ends in. That is a choice, not a rule, and this labels dates only. It is not tax advice.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and the results below were recalculated and compared with a Python model (146 workbooks, 13,177 date rows). The formulas use only YEAR, MONTH, MOD, INT, IF and AND, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Sales dates with fiscal year, fiscal quarter and FY label columns for eighteen sample rows](images/2026-11-09-team-02-fiscal-year-quarter-from-date-custom-start-month-excel-google-sheets/fiscal-year-quarter-from-date-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. Dates go in column **A**, amounts in column **B**, and the results fill **C:E** from row 4.

1. In **B1**, type the number of the first month of your year: 7 for July, 10 for October, 1 for a calendar year.
2. Type your dates in **A4** and down. Use real dates, not text. Sort them if you like, but the formulas do not need it.
3. In **C4**, type `=YEAR(A4)+IF(AND($B$1>1,MONTH(A4)>=$B$1),1,0)`. A date in or after the start month moves to the next year number, so July 2026 becomes FY2027. With a start month of 1, the fiscal year is the calendar year.
4. In **D4**, type `=INT(MOD(MONTH(A4)-$B$1,12)/3)+1`. `MOD` counts how many months the date sits after the start month, wrapping around the calendar. Dividing by 3 and adding 1 turns that into quarter 1 to 4.
5. In **E4**, type `="FY"&C4&" Q"&D4` for a label such as FY2027 Q2.
6. Select **C4:E4** and fill down.
7. To total by quarter, type the labels you want in a summary column, for example FY2027 Q2, and use `=SUMIFS($B$4:$B$103,$E$4:$E$103,A4)` next to each one.

The template wraps these formulas so a blank row stays blank and a bad entry shows Check. The versions above are the core.

Help pages: Microsoft's [YEAR function](https://support.microsoft.com/en-us/office/year-function-c64f017a-1354-490d-981f-578e8ec8d3b9) and [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3), and Google's [MOD](https://support.google.com/docs/answer/3093497). They describe the functions, not our test.

## Example

Sample data is fictional: 18 sales from 06/15/2026 to 07/15/2027, with the first month of the year set to 7.

| Date | Amount | FY | Quarter | Label |
|---|---|---|---|---|
| 06/15/2026 | $140.00 | 2026 | 4 | FY2026 Q4 |
| 08/15/2026 | $120.00 | 2027 | 1 | FY2027 Q1 |
| 11/09/2026 | $47.25 | 2027 | 2 | FY2027 Q2 |
| 01/05/2027 | $95.00 | 2027 | 3 | FY2027 Q3 |
| 06/30/2027 | $99.99 | 2027 | 4 | FY2027 Q4 |
| 07/01/2027 | $110.00 | 2028 | 1 | FY2028 Q1 |

Take 11/09/2026. The month is 11, and 11 is at least 7, so the year is 2026 plus 1, which is 2027. For the quarter, 11 minus 7 is 4, `MOD(4,12)` is 4, 4 divided by 3 rounds down to 1, and adding 1 gives quarter 2.

The summary tab for the full 18 sales:

| Quarter | Total | Sales |
|---|---|---|
| FY2026 Q4 | $140.00 | 1 |
| FY2027 Q1 | $269.75 | 3 |
| FY2027 Q2 | $495.25 | 4 |
| FY2027 Q3 | $386.25 | 4 |
| FY2027 Q4 | $313.99 | 4 |
| FY2028 Q1 | $151.00 | 2 |

The six quarters add up to $1,756.24, the same as all 18 sales. The template shows that check as $0.00 not in a listed quarter.

## Troubleshooting

### Every row says Check
The start month in **B1** must be a whole number from 1 to 12. In the test, 0, 13, 7.5, text such as July, and a blank cell all made every row show Check. Type 7, not July. The template also limits **B1** to a whole number from 1 to 12 with data validation, but a paste can bypass it.

### One row says Check and the others are fine
A date typed as text, such as one with a leading apostrophe, is not a date. Retype it, or convert it first. A date with a time of day works. A text amount also shows Check in the row check column.

### The summary shows a quarter total of $0.00
The summary labels are typed text. If you change the start month, FY2027 Q2 no longer means the same months, so retype the labels in the summary column. The template flags this when the not in a listed quarter amount is no longer 0.

### The fiscal year is one off from what you expected
This sheet names the year after the calendar year it ends in, so a July start puts July 2026 in FY2027. Some businesses name it after the year it starts in. If yours does, subtract 1 from the year formula, or follow your own records. The sheet does not decide which is correct.

### You have more than 100 sales
The template has 100 sales rows and 12 summary rows. Extend the ranges to the same last row before adding more.

## Template

[Download the .xlsx](templates/tidy-tabs-fiscal-year-quarter.xlsx)

The file has three tabs. **Sales** has 100 rows from row 4, 18 fictional sample sales, and a **Row check** column. **Summary** has 12 slots for quarter labels with total and count, and a side block with the total of all sales, the total of listed quarters, the amount not in a listed quarter, rows showing Check, and a **CHECK** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
