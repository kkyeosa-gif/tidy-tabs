---
title: How to Make a Monthly Calendar That Changes With the Month and Year in Excel and Google Sheets
labels: excel-formulas, small-business, calendar
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model and Python calendar, including 59 edited variants (every month of 2026 and 2027, February in 2015, 2021, 2024, 2028, 2100, years 1901 and 9999, month 13, 0, -1, 2.5, text and blank, year as text, 1900 and 10000, an event with a time part, an event typed as text, no events). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Make a monthly calendar that changes with the month and year in Excel and Google Sheets. Free printable template with an events count per day.
image_prompts: A cork board with colorful push pins and a few small blank paper tags beside a ceramic pencil jar and a potted herb on a craft studio wall shelf, soft morning light, no screens, no readable text, no logos
image_alt: Corkboard with colored pins and blank paper tags above a pencil jar and herb pot
threads: Need a printable month grid that changes when you type a new month?\nType a year and a month number, then one long formula fills all 42 cells.\nOctober 2026 starts on a Thursday and leaves the last row empty. The blank cells are the point.
---
Type the year in B1 and the month number in B2, set B3 to `=DATE(B1,B2,1)`, then fill A6:G11 with `=IF(MONTH($B$3-WEEKDAY($B$3)+1+(ROW()-ROW($A$6))*7+COLUMN()-COLUMN($A$6))=$B$2,$B$3-WEEKDAY($B$3)+1+(ROW()-ROW($A$6))*7+COLUMN()-COLUMN($A$6),"")`. Each cell shows its date, or stays blank outside the month, and the grid starts on Sunday.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Monthly calendar grid for October 2026 with the first day on Thursday and the weekday headings](images/2026-10-24-team-02-monthly-calendar-grid-from-month-year-weekday-excel-google-sheets/monthly-calendar-october-2026-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: find the first day of the month, back up to the Sunday on or before it, then count forward one day per cell. A cell keeps its date only if that date falls in the month you typed.

1. In **B1**, type the year, such as 2026.
2. In **B2**, type the month number, 1 to 12. October is 10.
3. In **B3**, type `=DATE(B1,B2,1)` and format it as a date. For 2026 and 10 it shows 10/01/2026.
4. In **A5:G5**, type Sun, Mon, Tue, Wed, Thu, Fri, Sat.
5. In **A6**, type the long formula from the top of this post. `ROW()` and `COLUMN()` tell each cell how far it sits from A6.
6. Copy **A6** to the whole block **A6:G11**. That is 6 rows of 7 cells, enough for any month.
7. Select **A6:G11** and set the number format to `d`. In Excel use **Home > Number Format > More Number Formats > Custom**. In Google Sheets use **Format > Number > Custom number format**. Only the day number shows, but each cell still holds a real date.

How the grid formula reads:

- `$B$3-WEEKDAY($B$3)+1` is the Sunday on or before the 1st. `WEEKDAY` returns 1 for Sunday, so a Thursday the 1st moves back 4 days.
- `(ROW()-ROW($A$6))*7` adds 7 days for each row down.
- `COLUMN()-COLUMN($A$6)` adds 1 day for each column right.
- `MONTH(...)=$B$2` keeps the date only when it belongs to the month. Otherwise the cell is `""`.

Microsoft documents the function on its help page: [WEEKDAY in Excel](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a). That page describes the function, not our test.

### What the template adds

The built cells are longer than the formulas above. **B3** is wrapped in a test, so it stays blank unless the year is a whole number from 1901 to 9999 and the month is a whole number from 1 to 12. Every grid cell is wrapped in `IF(ISNUMBER($B$3),...,"")`, so a blank B3 blanks the grid. The core of each is exactly the formula in the opening.

**A4** shows the title, such as October 2026, with `TEXT(B3,"mmmm yyyy")`. If B1 or B2 is not valid, it asks for a year and a month number instead.

### In Excel

The formulas work as typed. DATE, WEEKDAY, MONTH, ROW, COLUMN and COUNTIF are old functions, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed. The format code for the day of the month is also `d`, but nothing here was run in Sheets.

## Example

Sample data is fictional: a craft seller's events for October 2026. The finished grid for October 2026 starts on Thursday, so A6:D6 are blank and E6 shows 1.

| | Sun | Mon | Tue | Wed | Thu | Fri | Sat |
|---|---|---|---|---|---|---|---|
| Row 6 | | | | | 1 | 2 | 3 |
| Row 7 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |
| Row 10 | 25 | 26 | 27 | 28 | 29 | 30 | 31 |
| Row 11 | | | | | | | |

Rows 8 and 9 are full weeks (11 to 17 and 18 to 24). Row 11 stays empty because October needs only 5 rows.

The second block, **A14:G19**, counts events per day from the **Events** tab with `=IF(A6="","",COUNTIF(Events!$A$2:$A$200,A6))`. Put one event per row, with the date in column A and the description in column B. Days with an event turn green, and a day with none shows blank.

The sample has 13 events, 11 of them in October. Saturday 10/10/2026 has two (a farmers market and a wholesale order shipping), so G15 shows 2. The summary in **I1:J5** shows 31 days, 13 events, 11 counted in the grid, 2 outside the month, and OK in the check cell.

## Troubleshooting

### The calendar is blank
The month is outside 1 to 12, or the year is outside 1901 to 9999. B3 stays blank and the title in A4 tells you what to type. Month 13, 0, 2.5 and text gave a blank grid with no error values in LibreOffice.

### The days show as 10/01/2026 instead of 1
The grid cells need the number format `d`. The cells hold full dates, so a date format shows the whole date. Set the format on A6:G11 as in step 7.

### An event is missing from the counts
The date probably has a time part, such as 10/10/2026 2:00 PM, or was typed as text. COUNTIF matches whole dates only, so neither is counted, and the check cell **J5** says CHECK. Retype the date without a time.

### January or a leap February looks off
Every month of 2026 and 2027, February in 2024, 2028 and 2100, and the 12/2026 to 01/2027 rollover matched Python `calendar` in LibreOffice. Dates before 1901 are blocked on purpose, because Excel and LibreOffice count early 1900 days differently and that was not compared.

## Template

[Download the .xlsx](templates/tidy-tabs-monthly-calendar.xlsx)

The file has three tabs. **Calendar** holds the yellow inputs, the year in **B1** and the month number in **B2**, and the gray formulas: the first day in **B3**, the title in **A4**, the date grid in **A6:G11**, the events per day in **A14:G19** and a summary with the check in **I1:J5**. **Events** holds the date and description in **A2:B200**. **How to use** has short notes.

The events are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. The green event days and the month dropdown on B2 were written into the file but were not seen in a spreadsheet window, and printing was checked as a PDF export only (US Letter, landscape). To open it in Google Sheets, go to **File > Import > Upload** and select the file.
