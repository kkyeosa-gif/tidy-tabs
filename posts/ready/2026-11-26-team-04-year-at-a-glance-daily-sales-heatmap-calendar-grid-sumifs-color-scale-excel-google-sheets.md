---
title: How to Make a Year-at-a-Glance Daily Sales Heatmap in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 217 workbooks and 1,182,433 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The color scale was seen only in a LibreOffice PDF export.
search_description: Turn a daily sales log into a 12 x 31 grid: DATE for each month, SUMIFS for each day, blanks for dates like 02/30, and a color scale for the best days.
image_prompts: A back room wall in a small candle shop with a large paper year-at-a-glance wall calendar where many day squares are colored in with light and dark green marker, a cup of green markers on a shelf below it, and a stack of folded market booth tablecloths beside the cup, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Wall year calendar with days shaded in green marker above a cup of markers
threads: Twelve rows, 31 columns, one SUMIFS per cell, and a whole year of daily sales fits on one page.\nDATE(year,month+1,1)-1 finds the month's last day, so 02/30 stays blank instead of showing March 1.\nAdd a color scale and the big days stand out. In the sample, 11/21 is the darkest at $1,650.50.
---
Put the months down column A with `=DATE($B$1,1,1)` and the days 1 to 31 across row 4, fill each cell with `=SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$A$5:$A$5004,$A5+B$4-1,'Sales log'!$D$5:$D$5004,"OK")`, and add a color scale. The made-up 2026 log then shows $49,874.73 across 281 selling days on one page, with the darkest cell on 11/21/2026 at $1,650.50.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (217 workbooks, 1,182,433 cell comparisons, 0 mismatches), plus the edge cases below. The color scale was set as a rule in the file and seen only in a LibreOffice PDF export. Excel and Google Sheets were not opened for this post.

![Twelve-month by 31-day sales grid shaded green, darkest on November 21](images/2026-11-26-team-04-year-at-a-glance-daily-sales-heatmap-calendar-grid-sumifs-color-scale-excel-google-sheets/daily-sales-year-heatmap-template.png) *Year grid tab of a fictional shop's 2026 sales log: each cell is one day's sales in whole dollars, darker green for bigger days, dashes for days with no sales, blanks for dates like 02/30, December still empty, and month totals with each month's best day; 11/21/2026 (1,651) is the darkest cell. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The grid covers one calendar year. Every cell holds a formula, so you only type the year and the sales.

### Build the grid (same in Excel and Google Sheets)

1. On the **Sales log** tab, type one sale (or one day's total) per row from row 5: Date in **A**, Amount in **B**, optional Note in **C**. Several rows on one date are added together, and a refund is a negative amount. Type dates only, with no time of day.
2. Column **D** checks each row with `=IF(AND(A5="",B5="",TRIM(C5)=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(B5="","Check: no amount",IF(NOT(ISNUMBER(B5)),"Check: amount is text",IF(YEAR(A5)<>'Year grid'!$B$1,"Other year (not in the grid)","OK"))))))`. Only rows marked OK reach the grid.
3. On the **Year grid** tab, type the year in **B1**. **AN4** checks it: a whole number from 1901 to 9998 shows OK.
4. **A5** starts January with `=IF($AN$4<>"OK","",DATE($B$1,1,1))`, formatted as `mmmm` so it reads "January". **A6** uses `DATE($B$1,2,1)`, and so on down to December in **A16**. See the DATE help from [Microsoft](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) or [Google](https://support.google.com/docs/answer/3092969).
5. Row 4 holds the plain numbers 1 to 31 in **B4:AF4**.
6. **B5**, copied to every cell of **B5:AF16**, is `=IF($A5="","",IF(B$4>DAY(DATE(YEAR($A5),MONTH($A5)+1,1)-1),"",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$A$5:$A$5004,$A5+B$4-1,'Sales log'!$D$5:$D$5004,"OK")))`. `$A5+B$4-1` is the date for that cell, and `DATE(YEAR,MONTH+1,1)-1` is the last day of the month, so 02/30 and 04/31 stay blank instead of showing March 1 or May 1. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496), and DAY from [Microsoft](https://support.microsoft.com/en-us/office/day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101) or [Google](https://support.google.com/docs/answer/3093040).
7. Format the grid as `#,##0;-#,##0;"-"`. Cells show whole dollars, and a day with $0.00 shows a dash.
8. Add the month columns: **AG5** `=IF($A5="","",SUM(B5:AF5))` for the total, **AH5** `=IF($A5="","",COUNTIF(B5:AF5,">0"))` for days with sales, **AI5** `=IF(OR($A5="",COUNTIF(B5:AF5,">0")=0),"",$A5+MATCH(MAX(B5:AF5),B5:AF5,0)-1)` for the best day and **AJ5** `=IF(AI5="","",MAX(B5:AF5))` for its sales. See MAX from [Microsoft](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098) or [Google](https://support.google.com/docs/answer/3094013) and MATCH from [Microsoft](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) or [Google](https://support.google.com/docs/answer/3093378).
9. The summary in **AM4:AN14** totals the grid (**AN5**) and the log for the year (**AN6**, `=IF(AN4<>"OK","",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$D$5:$D$5004,"OK",'Sales log'!$A$5:$A$5004,">="&DATE(B1,1,1),'Sales log'!$A$5:$A$5004,"<"&DATE(B1+1,1,1)))`).
10. **AN9** finds the best day with `=IF(OR(AN4<>"OK",COUNT(AJ5:AJ16)=0),"",INDEX(AI5:AI16,MATCH(MAX(AJ5:AJ16),AJ5:AJ16,0)))` and **AN11** the best month with `=IF(OR(AN4<>"OK",MAX(AG5:AG16)<=0),"",INDEX(A5:A16,MATCH(MAX(AG5:AG16),AG5:AG16,0)))`. **AN8** is the average per day with sales, `=IF(OR(AN4<>"OK",N(AN7)=0),"",ROUND(AN5/AN7,2))`. INDEX help: [Microsoft](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) or [Google](https://support.google.com/docs/answer/3098242).
11. **AN14** is the final CHECK: `=IF(AN4<>"OK","Check",IF(AND(AN13="OK",ROUND(AN5-AN6,2)=0),"OK","Check"))`. It shows OK only when the grid total equals the log total and the Sales log has no Check rows.

### Add the color scale

The template already has a 3-color scale on **B5:AF16**: the number 0 is white, the 75th percentile is light green and the highest value is dark green. Using the 75th percentile as the middle keeps ordinary days from all looking white next to one big holiday market day. A negative day shows white.

To add or change it yourself, follow the app's own help. These menu paths come from those pages; we did not try them:

- **In Excel:** select **B5:AF16**, then **Home > Conditional Formatting > Color Scales**, or **Manage Rules** to set the minimum, midpoint and maximum types. See Microsoft's [conditional formatting help](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-in-excel-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f).
- **In Google Sheets:** select the range, click **Format > Conditional formatting**, open **Color scale**, and pick a minimum, maximum and optional midpoint. See Google's [conditional formatting help](https://support.google.com/docs/answer/78413).

For a one-month version laid out by weekday, see [How to Make a Monthly Calendar That Changes With the Month and Year](https://tidytabs.blogspot.com/2026/10/how-to-make-monthly-calendar-that.html). To see sales by time of day instead, see [How to Find Your Busiest Hour of the Day From Order Timestamps](https://tidytabs.blogspot.com/2026/10/how-to-find-your-busiest-hour-of-day.html).

## Example

The sales are made up: a small shop's log of 329 rows from 01/01/2026 to 11/25/2026, with one Etsy row per open day (closed Mondays and 07/04/2026), a second Market row every Saturday, a $640.00 Holiday market row on 11/21/2026 and a -$38.50 refund on 03/17/2026. **B1** is 2026.

| Month | Total | Days with sales | Best day | Best day sales |
|---|---|---|---|---|
| January | $2,878.94 | 27 | 01/03/2026 | $331.36 |
| February | $2,662.24 | 24 | 02/07/2026 | $318.65 |
| March | $3,432.97 | 26 | 03/21/2026 | $426.38 |
| April | $3,960.34 | 26 | 04/11/2026 | $473.88 |
| May | $5,279.40 | 27 | 05/23/2026 | $640.67 |
| June | $4,234.09 | 25 | 06/27/2026 | $587.36 |
| July | $4,003.81 | 26 | 07/25/2026 | $536.49 |
| August | $4,550.29 | 26 | 08/22/2026 | $552.06 |
| September | $4,939.65 | 26 | 09/19/2026 | $546.97 |
| October | $6,292.00 | 27 | 10/10/2026 | $680.76 |
| November | $7,641.00 | 21 | 11/21/2026 | $1,650.50 |
| December | $0.00 | 0 | | |

In the grid, 01/01 shows 57, 01/03 shows 331, Monday 02/02 shows a dash, 02/29 to 02/31 are blank, and 11/21 shows 1,651 in the darkest green. The summary reads $49,874.73 for both the grid and the log, 281 days with sales, $177.49 per day with sales, best day 11/21/2026 ($1,650.50), best month November ($7,641.00) and CHECK OK.

## Troubleshooting

### CHECK is red but the dates look right
A date has a time of day or is text, so no grid cell matches it. The Sales log shows "Check: date has a time" or "Check: date blank or text" in column D. Retype it as a plain date.

### The whole grid is blank
**B1** is text, blank, not a whole number, or outside 1901 to 9998, so **AN4** shows Check. Type a plain year such as 2026.

### A month looks empty
Its rows are dated in another year, so the Sales log marks them "Other year (not in the grid)". In the tests, 12/31/2025 and 01/01/2027 were kept out of a 2026 grid this way. Change **B1** or fix the dates.

### Adding up the grid cells does not match the month total
The grid only displays each day in whole dollars; the cells still hold the cents, and **AG** and the summary add up the full amounts. A sum of the numbers you see on screen can therefore differ from the totals. Read money from the totals, not from the displayed grid.

### Two days tie for best day
MATCH returns the first one. In a tie test with $100.00 on 01/05 and 06/05, the best day showed 01/05.

## Template

[Download the .xlsx](templates/tidy-tabs-daily-sales-year-heatmap.xlsx)

**Year grid** holds the year in **B1**, the 12 x 31 grid with its color scale, month totals, days with sales, best days and the summary with CHECK. **Sales log** takes 5,000 rows (5 to 5004) with a summary in **G4:H9**: rows OK, total of OK rows, rows from other years, rows showing Check, rows typed below row 5004, and its own CHECK. **How to use** explains each formula. The Year grid prints on one Landscape US Letter page, and both data tabs repeat the header row.

Yellow cells are inputs and gray cells are formulas. Every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file. We have not checked how Sheets or Excel draw the imported color scale.
