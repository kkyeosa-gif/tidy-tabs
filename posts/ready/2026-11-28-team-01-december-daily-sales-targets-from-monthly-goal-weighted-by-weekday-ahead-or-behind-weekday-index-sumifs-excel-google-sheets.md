---
title: How to Set Daily Sales Targets From a Monthly Goal by Day of the Week in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 235 workbooks and 80,535 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Split a monthly sales goal into daily targets by weekday weight, round the running total so days add up to the goal, and see if you are ahead or behind.
image_prompts: A small gift shop checkout counter in December with a paper wall calendar page where the weekend squares are shaded in pencil and one column is crossed out, a short stack of kraft paper shopping bags with twine handles, and a shallow tray of small wrapped holiday gifts, warm afternoon light, realistic photo, no screens, no readable text, no logos
image_alt: Shop counter with a pencil-marked paper calendar, kraft shopping bags and wrapped gifts
threads: Every Tuesday in December looks behind, and every Saturday looks great.\nThat is the even split: a $9,000.00 goal asks $290.32 every day.\nWeight the weekdays and round the running target with =ROUND($B$2*D5/$D$36,2), and Tuesday 12/01 gets $238.10 while Saturday 12/05 gets $595.23.
---
Give each weekday a weight, then set each day's running target to its share of the goal with `=ROUND($B$2*D5/$D$36,2)` (D5 is the running total of weights, D36 the month's total weight) and subtract the row above, so the days always add up to the goal. With a $9,000.00 December goal, Saturdays weighted 2.5, Tuesdays 1 and Mondays closed, Tuesday 12/01/2026 gets $238.10 and Saturday 12/05/2026 gets $595.23.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (235 workbooks, 80,535 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![December grid of weekday-weighted daily targets, actual sales so far and a $223.82 shortfall](images/2026-11-28-team-01-december-daily-sales-targets-from-monthly-goal-weighted-by-weekday-ahead-or-behind-weekday-index-sumifs-excel-google-sheets/december-daily-sales-targets-template.png) *Targets tab of a made-up shop's December 2026 with a $9,000.00 monthly goal: daily targets weighted by weekday (Saturday 2.5, Monday closed at 0.0, a yellow 3.5 override on 12/19 and 0.0 on 12/25), actual sales through 12/12/2026, still needed per day for the rest of the month, and the summary with $3,704.75 sold against $3,928.57 targeted, -$223.82 behind in orange. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The weights are your own guess, or last year's share of sales by weekday (see [How to Total Sales by Day of the Week](https://tidytabs.blogspot.com/2026/10/how-to-total-sales-by-day-of-week-in.html)). A weight of 0 means closed. The targets are a plan for spreading the goal, not a forecast, and the sheet does not tell you what goal to set.

The steps are the same in Excel and Google Sheets.

1. On the **Targets** tab, type the 1st of the month in **B1**, the monthly goal in **B2**, and the last day your sales are counted through in **B3**.
2. Type a weight for each weekday in **M5:M11**, Sunday to Saturday. Saturday at 2.5 and Tuesday at 1 means a Saturday should sell two and a half times a Tuesday.
3. For a holiday, type an override weight in column **B** next to that date, or 0 for closed.
4. Column **A** lists every date of the month. Rows past the month's last day stay blank.
5. **C5** picks each date's weight: `=IF(A5="","",IF(B5<>"",B5,INDEX($M$5:$M$11,WEEKDAY(A5))))`. WEEKDAY returns 1 for Sunday through 7 for Saturday. See the WEEKDAY help from [Microsoft](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a) or [Google](https://support.google.com/docs/answer/3092985).
6. **D5** keeps a running total of weights with `=IF(A5="","",SUM($C$5:C5))`, and **D36** totals the month with `=SUM(C5:C35)`.
7. **E5** gives the cumulative target, `=IF(OR(A5="",$P$5<>"OK"),"",ROUND($B$2*D5/$D$36,2))`, and **F5** turns it into the daily target with `=IF(E5="","",ROUND(E5-N(E4),2))`. Because the running total is rounded instead of each day, the daily targets add up to the goal to the cent. See the ROUND help from [Microsoft](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) or [Google](https://support.google.com/docs/answer/3093440).
8. On the **Sales log** tab, type one sale or one day's total per row: Date in **A**, Amount in **B**, Note in **C**. Column **D** shows each row's Status.
9. Back on **Targets**, **G5** totals each day's OK sales with `=IF(A5="","",SUMIFS('Sales log'!$B$5:$B$1004,'Sales log'!$A$5:$A$1004,A5,'Sales log'!$D$5:$D$1004,"OK"))`. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
10. Read the summary in **O4:P20**: target to date in **P11**, sold to date in **P12**, ahead (+) or behind (-) in **P13** and **P14**, and still needed in **P15**.
11. Column **J** spreads the amount still needed over the days after B3 by the same weights and the same rounding.

**CHECK** in **P20** is OK only when the inputs are valid, the targets add up to the goal, and the grid's sales match the log.

To track the month against a single goal line instead, see [How to Track Progress to a Monthly Sales Goal](https://tidytabs.blogspot.com/2026/10/how-to-track-progress-to-monthly-sales.html).

## Example

The shop and its sales are made up. The goal is $9,000.00 for December 2026, with weights Sunday 1.5, Monday 0, Tuesday 1, Wednesday 1, Thursday 1.2, Friday 1.8 and Saturday 2.5. Overrides: 12/19 at 3.5, 12/24 at 1.0, 12/25 at 0 and 12/31 at 0.8. Sales are counted through 12/12/2026.

| Date | Weight | Cumulative target | Daily target | Actual to date | Still needed per day |
|---|---|---|---|---|---|
| Tue 12/01/2026 | 1.0 | $238.10 | $238.10 | $212.40 | |
| Sat 12/05/2026 | 2.5 | $1,785.71 | $595.23 | $1,695.95 | |
| Sat 12/12/2026 | 2.5 | $3,928.57 | $595.24 | $3,704.75 | |
| Sun 12/13/2026 | 1.5 | $4,285.71 | $357.14 | | $372.90 |
| Sat 12/19/2026 | 3.5 | $6,309.52 | $833.33 | | $870.11 |
| Fri 12/25/2026 | 0.0 | $7,380.95 | $0.00 | | $0.00 |
| Thu 12/31/2026 | 0.8 | $9,000.00 | $190.48 | | $198.88 |
| Month total | 37.8 | | $9,000.00 | $3,704.75 | $5,295.25 |

The month has 26 open days and 5 closed (four Mondays and Christmas Day). An even split would ask $290.32 every day, or $346.15 per open day, so every Tuesday looks behind and every Saturday ahead.

By 12/12/2026 the weighted target to date is $3,928.57 and the shop sold $3,704.75, so it is $223.82 behind (-5.7%). The $5,295.25 still needed is spread over the 19 days left, with 21.3 of the month's weight.

## Troubleshooting

These cases were reproduced in LibreOffice Calc 24.2.7.2.

### The daily targets add up to a few cents more or less than the goal
Each day was rounded on its own. In a scratch copy of the sample, rounding each day's share of $9,000.00 summed to $9,000.03, and rounding each remaining day's share of the $5,295.25 still needed summed to $5,295.21. Round the running total and subtract the row above, as columns E and F do.

### Every target is blank and CHECK says Check
A setting or a weight is wrong: B1 is not the 1st of a month, every weight is 0, or a weekday weight or override is text or negative.

### A day's sales are missing and CHECK is red
The Sales log date has a time of day, so its Status says "Check: date has a time" and it is not counted. Retype it as a plain date.

### A sale shows "Outside the month (not counted)"
Its date is before B1 or after the month's last day. In the sample, an 11/28/2026 sale of $455.60 is listed but not counted toward December.

### Still needed shows $0.00 a day
The goal is already reached. With the sample sales and a $3,000.00 goal, still needed is $0.00, not a negative number.

## Template

[Download the .xlsx](templates/tidy-tabs-december-daily-sales-targets.xlsx)

**Targets** holds the settings in **B1:B3**, one row per day in rows 5 to 35 with the month total in row 36, the weekday weights in **L4:M11** and the summary in **O4:P20**. **Sales log** takes 1,000 rows (rows 5 to 1004) with its own CHECK cell, which also catches rows typed below row 1004. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Closed days and behind-target cells are light orange, days at or above target are light green, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
