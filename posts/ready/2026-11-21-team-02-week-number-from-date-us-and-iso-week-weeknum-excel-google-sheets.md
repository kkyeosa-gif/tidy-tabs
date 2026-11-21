---
title: How to Get the Week Number From a Date (US and ISO Weeks) in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 26 workbooks and 130,052 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Get the ISO week with =WEEKNUM(date,21), its year with =YEAR(date-WEEKDAY(date,2)+4), a 2026-W47 label, and US Sunday-start weeks with a day count.
image_prompts: A large paper wall calendar turned to late December hanging in a small bakery back room, a tear-off weekly planner pad and a metal date stamper with an ink pad on the counter below it, soft morning light, realistic photo, no screens, no readable text, no logos
image_alt: Paper wall calendar at late December above a weekly planner pad and date stamper
threads: For ISO weeks, use =WEEKNUM(date,21) and take the year from =YEAR(date-WEEKDAY(date,2)+4).\nPlain YEAR() labels 01/01/2027 as 2027-W53. It belongs to 2026-W53.\nIn LibreOffice, WEEKNUM(date,1) also said week 1 on 12/28/2026, where a day count gives 53.
---
Get the ISO week (Monday start) with `=WEEKNUM(date,21)` and its year with `=YEAR(date-WEEKDAY(date,2)+4)`, then join them as `=iso_year&"-W"&TEXT(iso_week,"00")` to get labels like "2026-W47". For US weeks (Sunday start, the week of January 1 is week 1) the download counts days with `=INT((date-DATE(YEAR(date),1,1)+WEEKDAY(DATE(YEAR(date),1,1))-1)/7)+1`, so 12/27/2026 is US week 53 but ISO week 52, and 01/01/2027 is US week 1 but ISO week 53 of 2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (26 workbooks, 130,052 cell comparisons, 0 mismatches), plus the edge cases below. The functions used also exist in Excel and Google Sheets, but neither app was opened for this post.

![Fifteen dates from late 2025 to 2027 with US weeks, ISO weeks and labels](images/2026-11-21-team-02-week-number-from-date-us-and-iso-week-weeknum-excel-google-sheets/week-number-us-vs-iso-week-template.png) *Dates tab with 15 fictional dates where 12/27/2026 is US week 53 but ISO week 52 and 01/01/2027 is US week 1 but ISO week 53, labeled 2026-W53, while the YEAR() column wrongly shows 2027-W53. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Dates** tab, one date per row starting in row 5.

1. Type a date in **A5**, such as 11/21/2026. Type it as a real date, not text.
2. In **B5**, type `=INT((A5-DATE(YEAR(A5),1,1)+WEEKDAY(DATE(YEAR(A5),1,1))-1)/7)+1` for the US week. Weeks start on Sunday and the week holding January 1 is week 1, so some years end in week 53, and 12/31/2028 is week 54.
3. In **C5**, type `=WEEKNUM(A5,21)` for the ISO week. Type 21 means Monday start, and week 1 is the week with the year's first Thursday. Microsoft's [WEEKNUM help](https://support.microsoft.com/en-us/office/weeknum-function-e5c43a03-b4ab-426c-b411-b18c13c75340) and Google's [WEEKNUM help](https://support.google.com/docs/answer/3294949) list the type codes.
4. In **D5**, type `=YEAR(A5-WEEKDAY(A5,2)+4)` for the ISO year. It finds the Thursday of the date's ISO week and takes that Thursday's year, so 01/01/2027 belongs to 2026.
5. In **E5**, type `=D5&"-W"&TEXT(C5,"00")` for the year-week label. TEXT pads week 1 to "01" so the labels sort in order.
6. In **F5**, type `=INT((DAY(A5)+WEEKDAY(DATE(YEAR(A5),MONTH(A5),1))-2)/7)+1` for the week of the month. There is no standard for this one; here weeks start on Sunday and the 1st is always in week 1.
7. In **G5**, type `=INT(A5)-WEEKDAY(A5,2)+1` and format it as a date to get the Monday that starts the ISO week. INT drops any time part.
8. Fill **B5:G5** down. In the template every formula is wrapped in `IF($H5<>"OK","",...)`, and **H5** marks text dates and numbers below 1 with a Check label.

ISOWEEKNUM is the newer name for the same number ([Microsoft's ISOWEEKNUM help](https://support.microsoft.com/en-us/office/isoweeknum-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e)); it returned the same weeks in LibreOffice, but the template sticks with WEEKNUM type 21.

Fiscal or retail weeks (such as a 4-5-4 calendar) follow your own year, not these rules. For a custom year start, see the Tidy Tabs post on [getting the fiscal year and quarter from a date](https://tidytabs.blogspot.com/2026/11/how-to-get-fiscal-year-and-quarter-from.html).

## Example

The dates are made up: market days and order dates around New Year and in November.

| Date | US week | ISO week | Year-week label | Week of month | ISO week starts | Label with YEAR() (wrong) |
|---|---|---|---|---|---|---|
| 12/29/2025 | 53 | 1 | 2026-W01 | 5 | 12/29/2025 | 2025-W01 |
| 01/04/2026 | 2 | 1 | 2026-W01 | 2 | 12/29/2025 | 2026-W01 |
| 11/01/2026 | 45 | 44 | 2026-W44 | 1 | 10/26/2026 | 2026-W44 |
| 11/21/2026 | 47 | 47 | 2026-W47 | 3 | 11/16/2026 | 2026-W47 |
| 11/22/2026 | 48 | 47 | 2026-W47 | 4 | 11/16/2026 | 2026-W47 |
| 12/27/2026 | 53 | 52 | 2026-W52 | 5 | 12/21/2026 | 2026-W52 |
| 12/28/2026 | 53 | 53 | 2026-W53 | 5 | 12/28/2026 | 2026-W53 |
| 01/01/2027 | 1 | 53 | 2026-W53 | 1 | 12/28/2026 | 2027-W53 |
| 01/04/2027 | 2 | 1 | 2027-W01 | 2 | 01/04/2027 | 2027-W01 |
| 05/01/2027 | 18 | 17 | 2027-W17 | 1 | 04/26/2027 | 2027-W17 |

Sunday 11/22/2026 starts US week 48 but is still the last day of ISO week 47. With "2026-W47" typed in **M7**, the summary counts 3 of the 15 sample dates in that week.

## Troubleshooting

### The label says 2027-W53 for 01/01/2027
The label was built with `YEAR(date)`, like the comparison column **J**. It also turns 12/29/2025 into "2025-W01". Use the ISO year from `=YEAR(date-WEEKDAY(date,2)+4)`.

### WEEKNUM shows week 1 on 12/28/2026
This is a LibreOffice Calc finding. Over every day from 01/01/2020 to 12/31/2029, `WEEKNUM(date,1)` in LibreOffice 24.2.7.2 returned 1 on 33 days between December 26 and December 31 where the counting formula gives 53 (54 on 12/31/2028), including 12/27/2026 to 12/31/2026. `WEEKNUM(date,2)` differed from a Monday-start count on 30 days. Type 21 had no mismatches. Excel and Google Sheets were not checked, so compare the help pages above with your own late-December dates, or use the counting formula in Step 2.

### A row shows Check: date is text
The date was typed or pasted as text, such as "11/21/2026" from a web form. In LibreOffice, a plain `WEEKNUM` of that text quietly returned 47, so the template flags it instead of guessing. Retype it as a date. Numbers like 0.5 or -3 show Check: not a date.

### Sundays show a different US and ISO week
That is expected: every Sunday starts a new US week but ends an ISO week. The **Calendar check** tab marks those days Differ in orange.

### CHECK in M10 shows Check
A row shows a Check status, the long ISO cross-check in column **I** disagrees with column **C**, a date is typed below row 504, or the tab is empty. **M5**, **M6** and **M9** tell you which.

## Template

[Download the .xlsx](templates/tidy-tabs-week-number-from-date.xlsx)

The file has three tabs. **Dates** holds up to 500 dates in rows 5 to 504 with Date (input in **A**), US week, ISO week, ISO year, Year-week label, Week of month, ISO week starts and Status in **B:H**, plus the long ISO formula without WEEKNUM in **I** and the wrong YEAR() label in **J** for comparison. The summary in **L4:M10** counts dates listed, rows showing Check, ISO weeks that differ from the long formula, rows with the year-week you type in **M7**, dates typed below row 504, and the CHECK.

**Calendar check** lists 62 days from the first day in **B1** (12/01/2026 in the sample) with the weekday, both week numbers, the ISO year, the label and the week of the month, and marks each day Same or Differ. **How to use** explains each formula. Both data tabs print Landscape on US Letter.

Yellow cells are inputs and gray cells are formulas. Every Check turns red and Differ turns orange. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
