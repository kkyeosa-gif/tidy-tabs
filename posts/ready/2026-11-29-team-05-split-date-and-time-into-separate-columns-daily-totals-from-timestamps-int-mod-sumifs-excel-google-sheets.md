---
title: How to Split a Date and Time Into Separate Columns in Excel and Google Sheets
labels: excel-formulas, small-business, dates, sales-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 226 workbooks and 2,348,592 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The 2:29 PM display of a plain MOD, the FALSE time match and the $0.00 and $69.75 date-only totals were seen in LibreOffice Calc only.
search_description: Split "11/28/2026 2:30 PM" into a date and a time with INT and MOD, then total sales by day from timestamps with a one-day SUMIFS window.
image_prompts: A small bakery counter after closing with a long curled paper register tape spilling out beside the register drawer, a short stack of handwritten paper order tickets held by a binder clip, and a wall clock above the counter just past midnight, dim warm light, realistic photo, no screens, no readable text, no logos
image_alt: Curled register tape, clipped paper order tickets and a wall clock past midnight
threads: SUMIFS by date, and only the midnight sales show up. Classic.\nA date is the number for 12:00 AM, so sum a one-day window instead: ">="&A5 and "<"&A5+1.\nIn the sample holiday weekend export, 11/28/2026 comes to $115.00.
---
The date is `=INT(A5)` and the time is `=ROUND(MOD(A5,1)*86400,0)/86400`, the time of day rounded to the nearest second: format the first column as MM/DD/YYYY and the second as h:mm AM/PM, and "11/28/2026 2:30 PM" splits into 11/28/2026 and 2:30 PM. To total a day straight from the export, sum a one-day window with `=SUMIFS(Export!$B$5:$B$2004,Export!$A$5:$A$2004,">="&A5,Export!$A$5:$A$2004,"<"&A5+1)`: the made-up holiday weekend export comes to $85.50 on 11/27/2026 and $115.00 on 11/28/2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (226 workbooks, 2,348,592 cell comparisons, 0 mismatches), plus the edge cases below. The 2:29 PM display of a plain MOD, the FALSE time match and the $0.00 and $69.75 date-only totals were seen in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Pasted sale timestamps split into date and time columns next to two daily sales totals](images/2026-11-29-team-05-split-date-and-time-into-separate-columns-daily-totals-from-timestamps-int-mod-sumifs-excel-google-sheets/split-date-and-time-daily-totals-template.png) *Left, the Export tab of a made-up holiday weekend export: each date-time split into a Date and a Time column, with the one sale pasted as text marked Check in red. Right, the By day tab: $85.50 on 11/27/2026 and $115.00 on 11/28/2026 from the one-day window, next to the $0.00 and $69.75 that matching the date alone gives (orange). PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

A date-time in a spreadsheet is one number. The whole part is the day and the fraction is the time of day, so 11/28/2026 12:00 AM is the same number as 11/28/2026. The steps are the same in Excel and Google Sheets.

1. On the **Export** tab, paste the date-time column of your shop or point-of-sale export into **A5** and the amount column into **B5**, one sale per row (up to row 2004, any order).
2. Column **C** (Date) keeps the whole day: `=IF(AND(A5="",B5=""),"",IF(AND(ISNUMBER(A5),N(A5)>=0),INT(A5),"Check"))`, formatted MM/DD/YYYY. The core is `=INT(A5)`. See the INT help from [Microsoft](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) or [Google](https://support.google.com/docs/answer/3093490).
3. Column **D** (Time) keeps what is left after the whole days: `=IF(AND(A5="",B5=""),"",IF(AND(ISNUMBER(A5),N(A5)>=0),ROUND(MOD(A5,1)*86400,0)/86400,"Check"))`, formatted h:mm AM/PM. Multiplying by 86,400 (seconds in a day), rounding and dividing back snaps the time to the nearest second. In LibreOffice Calc, a plain `=MOD(A5,1)` with the same format showed 2:29 PM for a 2:30 PM sale; the rounded version showed 2:30 PM. See the MOD help from [Microsoft](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) or [Google](https://support.google.com/docs/answer/3093497).
4. Column **E** (Hour) is `=IF(ISNUMBER(D5),HOUR(A5),"")`, and column **F** (Minutes after midnight) is `=IF(ISNUMBER(D5),INT(ROUND(D5*86400,0)/60),"")`. 2:30 PM is 870. To test for a set time, compare minutes: `=F6=870`.
5. Check column **G** (Status). Only OK rows count.
6. On the **By day** tab, type the first day in **B1** (11/27/2026) and how many days to show in **B2** (1 to 62). Column **A** lists one day per row.
7. Column **B** totals each day from 12:00 AM up to, not including, 12:00 AM the next day: `=IF(A5="","",SUMIFS(Export!$B$5:$B$2004,Export!$A$5:$A$2004,">="&A5,Export!$A$5:$A$2004,"<"&A5+1))`. A sale at 12:00 AM belongs to that day, and one at 11:59:59 PM to the day it is on. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
8. Column **C** totals the same day on the split Date column, `=IF(A5="","",SUMIFS(Export!$B$5:$B$2004,Export!$C$5:$C$2004,A5))`, and column **F** says OK when B and C agree. Column **D** counts orders on the same window.
9. Column **E** shows the wrong way once, matching the date alone: `=IF(A5="","",SUMIFS(Export!$B$5:$B$2004,Export!$A$5:$A$2004,A5))`. It only catches sales at exactly 12:00 AM.
10. For sales at one minute, type a time in **B3** (2:30 PM). **I13** turns it into minutes (870), **I14** counts the sales and **I15** totals them.
11. Read **CHECK** in **I11**. It is OK when the days shown add up to the Export total and no Export row shows Check.

In a sheet of your own, the short pattern is `=SUMIFS(B:B,A:A,">="&D5,A:A,"<"&D5+1)` with the day in D5. The template uses fixed ranges (rows 5 to 2004); whole-column ranges were not tested.

Times are taken as typed. The sheet does not change time zones; for that, see [how to convert time zones](https://tidytabs.blogspot.com/2026/10/how-to-convert-time-zones-for-client.html).

Related: [weekly sales totals from a daily log](https://tidytabs.blogspot.com/2026/11/how-to-add-up-weekly-sales-from-daily.html), [your busiest hour of the day](https://tidytabs.blogspot.com/2026/10/how-to-find-your-busiest-hour-of-day.html), [Unix timestamps to dates](https://tidytabs.blogspot.com/2026/11/how-to-convert-unix-timestamp-to-date.html) and [ship dates with a daily cutoff time](https://tidytabs.blogspot.com/2026/11/how-to-calculate-ship-date-with-daily.html).

## Example

The export is made up: eight sales from a small shop's holiday weekend, 11/27/2026 and 11/28/2026.

| Date and time (A) | Amount (B) | Date (C) | Time (D) | Minutes (F) | Status (G) |
|---|---|---|---|---|---|
| 11/27/2026 9:15 AM | $42.00 | 11/27/2026 | 9:15 AM | 555 | OK |
| 11/27/2026 2:30 PM | $18.50 | 11/27/2026 | 2:30 PM | 870 | OK |
| 11/27/2026 11:59 PM (11:59:59) | $25.00 | 11/27/2026 | 11:59 PM | 1439 | OK |
| 11/28/2026 12:00 AM | $60.00 | 11/28/2026 | 12:00 AM | 0 | OK |
| 11/28/2026 10:05 AM | $33.25 | 11/28/2026 | 10:05 AM | 605 | OK |
| 11/28/2026 2:30 PM | $12.00 | 11/28/2026 | 2:30 PM | 870 | OK |
| 11/28/2026 4:45 PM (text) | $20.00 | Check | Check | | Check: text, not a date-time |
| 11/28/2026 (no time) | $9.75 | 11/28/2026 | 12:00 AM | 0 | OK |

On the By day tab:

| Day | Sales (window) | Orders | Matching the date alone (wrong) |
|---|---|---|---|
| Fri 11/27/2026 | $85.50 | 3 | $0.00 |
| Sat 11/28/2026 | $115.00 | 4 | $69.75 |

Seven rows are OK, worth **$200.50**, and both days add up to that. The text row's $20.00 is left out and shown in **I10**, so CHECK says Check until it is fixed. With **B3** at 2:30 PM, I14 shows 2 sales and I15 shows $30.50.

## Troubleshooting

### A SUMIFS by date comes up short or shows $0.00
The export has times in it. A date such as 11/28/2026 is the number for 12:00 AM that day, so matching on it catches only the sales at exactly midnight. On the sample, the date-only SUMIFS in column E gave $0.00 for 11/27/2026 and $69.75 for 11/28/2026 in LibreOffice Calc ($60.00 + $9.75, the two 12:00 AM rows). Use the `">="&A5` and `"<"&A5+1` window, or sum on the split Date column.

### The Time column shows a minute early, like 2:29 PM
In LibreOffice Calc, plain `=MOD(A5,1)` formatted h:mm AM/PM showed 9:14 AM for the 9:15 AM sale and 2:29 PM for both 2:30 PM sales, while column A showed 9:15 AM and 2:30 PM. Use `=ROUND(MOD(A5,1)*86400,0)/86400` as in the template.

### A row that shows 2:30 PM does not match TIME(14,30,0)
The stored time can be a hair off 2:30. In LibreOffice Calc, `=MOD(Export!A6,1)=TIME(14,30,0)` was FALSE on both 2:30 PM rows of the sample, while `=Export!D6=TIME(14,30,0)` on the rounded column and `=F6=870` were TRUE. Compare whole minutes from column F.

### A row says "Check: text, not a date-time"
The date-time was pasted as text, so it is not counted. If the text has the day first, fix that before converting: see [how to fix dates that flip from MM/DD to DD/MM](https://tidytabs.blogspot.com/2026/09/how-to-fix-dates-that-flip-from-mmdd-to.html). Every split cell is guarded with ISNUMBER because in LibreOffice Calc a plain MOD on the text "11/28/2026 4:45 PM" converted it and returned 4:44 PM.

### CHECK says Check but every row is OK
Some OK sales fall outside the days shown (**I7** above $0.00): move **B1** to the earliest date in **Export J8** or raise **B2**. **Export J11** above 0 means rows were typed below row 2004, which the sheet does not read.

## Template

[Download the .xlsx](templates/tidy-tabs-split-date-and-time.xlsx)

**Export** takes 2,000 sales (rows 5 to 2004) with the Date, Time, Hour, Minutes after midnight and Status columns, plus a summary in **I4:J12**. **By day** holds the settings in **B1:B3**, up to 62 days (rows 5 to 66), and the checks in **H4:I15**. **How to use** explains each formula. Data tabs print Landscape on US Letter.

Clear A5:B2004 and paste your own export.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
