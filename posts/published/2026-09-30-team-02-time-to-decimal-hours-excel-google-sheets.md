---
threads_url: https://www.threads.com/@tin_ylab/post/DeFcrLFnTUe
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-convert-clock-times-to-decimal.html
title: How to Convert Clock Times to Decimal Hours in Excel and Google Sheets
labels: formulas, time-tracking
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Convert clock in and clock out times to decimal hours with =ROUND(MOD(C2-B2,1)*24,2), including overnight shifts, with a free .xlsx timesheet.
image_prompts: A cramped back room of a small US cafe with a wall rack of blank paper time cards next to an old mechanical punch clock, one card half inserted, warm overhead light, all card text blurred and unreadable, no logos
image_alt: Paper time cards in a wall rack beside a mechanical punch clock
threads: =ROUND(MOD(C2-B2,1)*24,2) turns 8:30 AM to 5:00 PM into 8.5 hours.\nMOD is the part that keeps a 10:00 PM to 6:30 AM shift positive instead of negative.\nIf the result still looks like a clock time, the cell is formatted as Time, not Number.
---
Multiply the time difference by 24: `=ROUND(MOD(C2-B2,1)*24,2)` turns 8:30 AM to 5:00 PM into 8.5 hours, and MOD keeps overnight shifts positive.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template's formulas were recalculated and checked against hand math. The Excel and Google Sheets steps below follow Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Timesheet showing clock times, unpaid breaks, decimal hours, and pay for five sample shifts](images/2026-09-30-team-02-time-to-decimal-hours-excel-google-sheets/time-to-decimal-hours-template.png) *LibreOffice Calc 24.2 PDF export of the Shifts sheet with sample data (fake). The Date and Nearest 15 min columns are hidden so the numbers stay readable.*

## Steps

Excel and Google Sheets store a time as a fraction of a day. Noon is 0.5, and 6:00 AM is 0.25. Subtracting two times gives a fraction of a day, so multiplying by 24 gives hours.

1. Put the clock in time in column B and the clock out time in column C, for example 8:30 AM and 5:00 PM.
2. Put the unpaid break, in minutes, in column D.
3. In E2, type `=MOD(C2-B2,1)-D2/1440` to get time worked. MOD wraps a negative result around a full day, so a shift from 10:00 PM to 6:30 AM comes out as 8:30 instead of a negative number. 1440 is the number of minutes in a day.
4. In F2, type `=ROUND(E2*24,2)` to get decimal hours, rounded to two places.
5. Fill E2 and F2 down for every shift.

If you skip the break column, the one-cell version is `=ROUND(MOD(C2-B2,1)*24,2)`.

### In Excel

1. Format the clock in and clock out cells as Time, so 8:30 AM is stored as a time and not as text.
2. Leave column E in a time format such as **h:mm**.
3. Format column F as a Number with 2 decimal places, not as Time. Without this, Excel may show the result as a time such as 12:00 AM.

Microsoft documents both functions: [MOD](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) and [MROUND](https://support.microsoft.com/en-us/office/mround-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427).

### In Google Sheets

1. Type the times in a form Sheets recognizes, such as 8:30 AM, so they are stored as times and not text.
2. Type the formulas from the steps above. They work the same way.
3. Format column F as a plain number with 2 decimals; Google explains number formats on its [Format numbers in a spreadsheet page](https://support.google.com/docs/answer/56470). Google's help pages for [MOD](https://support.google.com/docs/answer/3093497) and [MROUND](https://support.google.com/docs/answer/3093426) cover the functions.

## Example

Sample data below is fictional. The hourly rate of $22.50 is made up.

| Date | Clock in | Clock out | Break (min) | Time worked | Decimal hours | Pay |
|---|---|---|---|---|---|---|
| 09/21/2026 | 8:30 AM | 5:00 PM | 30 | 8:00 | 8.00 | $180.00 |
| 09/22/2026 | 9:00 AM | 5:45 PM | 30 | 8:15 | 8.25 | $185.63 |
| 09/23/2026 | 7:15 AM | 3:40 PM | 30 | 7:55 | 7.92 | $178.20 |
| 09/24/2026 | 10:00 PM | 6:30 AM | 30 | 8:00 | 8.00 | $180.00 |
| 09/25/2026 | 9:05 AM | 1:20 PM | 0 | 4:15 | 4.25 | $95.63 |

Total decimal hours: **36.42**. Total pay: **$819.46**.

The first row is the opening example: 8:30 AM to 5:00 PM is 8.5 hours, and the 30 minute break brings it to 8.00. The 09/24/2026 row crosses midnight, and MOD keeps it at 8 hours.

Pay is `=ROUND(F2*$K$1,2)`, where K1 holds the hourly rate. The 7:55 shift shows why rounding matters: 7 hours 55 minutes is 7.9166..., which displays as 7.92.

If your payroll rounds to quarter hours, the template also has a Nearest 15 min column with `=MROUND(E2*24,0.25)`. On these five shifts it gives 8, 8.25, 8, 8, and 4.25. Whether to round to two decimals, to a quarter hour, or not at all is your own payroll policy. The file only does the arithmetic.

## Troubleshooting

### The result looks like a time, such as 8:30 AM, not 8.5
The formula cell is still formatted as Time. Change the cell to a Number format with 2 decimals.

### An overnight shift shows a negative number or ####
Without MOD, 6:30 AM minus 10:00 PM is negative. Wrap the subtraction in `MOD(...,1)`, as in `=MOD(C2-B2,1)`.

### The formula returns #VALUE!
One of the times was typed as text, often with a stray space or a period in "a.m.". Retype it as `8:30 AM`, or check that the cell aligns to the right like a number.

### The total is off by a few hundredths
Each row is rounded to two decimals before the total, so the sum of rounded rows can differ from the rounded sum. Decide which one your payroll uses and keep it the same every pay period.

### The hours are wrong when a shift is longer than 24 hours
MOD wraps at one full day, so a span of 24 hours or more will not come out right. Use one row per day, or type the date and time together in each cell.

## Template

[Download the .xlsx](templates/tidy-tabs-time-to-decimal-hours.xlsx)

The file has two tabs. **Shifts** holds five sample shifts with Time worked, Decimal hours, Nearest 15 min, and Pay columns, plus an hourly rate in K1 and totals for hours and pay in K2 and K3. **How to use** explains each formula.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
