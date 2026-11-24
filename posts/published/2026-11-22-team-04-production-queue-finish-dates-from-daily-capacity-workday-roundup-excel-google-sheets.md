---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-estimate-when-each-order-will-be.html
title: How to Estimate When Each Order Will Be Finished From Your Daily Capacity in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 27 workbooks and 20,736 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Get a finish date for every order in your queue with ROUNDUP(units so far/units per day) and WORKDAY(start-1,day,days off), plus a late flag.
image_prompts: A wooden workbench in a small home woodshop with neat stacks of identical unfinished wooden cutting boards in several piles, a paper wall calendar with a few days circled pinned above, and a clipboard order list resting on the end of the bench, soft shop lighting, realistic photo, no screens, no readable text, no logos
image_alt: Workbench with stacks of unfinished cutting boards below a calendar with circled days
threads: Start WORKDAY one day early with =WORKDAY(start-1,day_number,days_off), so day 1 is the start date itself.\nThe day number is =ROUNDUP(units_so_far/units_per_day,0).\nAt 40 candles a day, a 300-unit order promised for 12/16/2026 lands on 12/18/2026.
---
Add up the units through each order, turn that into a work-day number with `=ROUNDUP(units_so_far/units_per_day,0)`, and get the finish date with `=WORKDAY(start-1,day_number,days_off)`, so day 1 is the start date itself. In the made-up queue at 40 candles a work day from Monday 11/23/2026, with 11/26/2026 and 11/27/2026 closed, the order that brings the total to 160 units finishes 11/30/2026, and a 300-unit order promised for 12/16/2026 finishes 12/18/2026, so the download marks it Late by 2 days.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (27 workbooks, 20,736 cell comparisons, 0 mismatches), plus the edge cases below. WORKDAY, NETWORKDAYS and ROUNDUP also exist in Excel and Google Sheets, but neither app was opened for this post.

![Seven candle orders in a queue with start dates, finish dates and late status](images/2026-11-22-team-04-production-queue-finish-dates-from-daily-capacity-workday-roundup-excel-google-sheets/production-queue-finish-dates-template.png) *Queue tab with 7 sample orders made at 40 units a work day from 11/23/2026, where W-104 shows Late by 1 day and W-106 shows Late by 2 days. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Queue** tab, one order per row from row 5, in the order you will make them.

1. Type the first day you can work in **B1** (11/23/2026) and how many units you finish per work day in **B2** (40).
2. On the **Days off** tab, type your shop's own closed days in **A2:A30**. Weekends are skipped already.
3. Back on **Queue**, type the order name in **A5**, the promised date in **B5** and the units in **C5**.
4. In **D5**, keep a running total of units with `=SUM($C$5:C5)`. The template uses `SUMIF($C$5:C5,">0")` so text, zero and negative units are not added.
5. In **F5**, type `=ROUNDUP(ROUND(D5/$B$2,9),0)` for the last work day number. 160 units at 40 a day is day 4. The inner ROUND keeps a value like 3.0000000001 from rounding up to an extra day ([Microsoft's ROUNDUP help](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7)).
6. In **E5**, type `=INT(ROUND((D5-C5)/$B$2,9))+1` for the first work day number. It is the day after the units ahead are done, or the same day if that day has room left.
7. In **G5**, type `=WORKDAY($B$1-1,E5,'Days off'!$A$2:$A$30)` for the start date, and in **H5** the same with **F5** for the finish date. Starting one day before B1 makes day 1 the start date itself when it is a work day. Microsoft's [WORKDAY help](https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33) and Google's [WORKDAY help](https://support.google.com/docs/answer/3093059) describe the holidays argument.
8. In **I5**, flag late orders with the core form `=IF(H5>B5,"Late by "&(H5-B5)&" days","On time")`. The template's version counts calendar days, uses "1 day" in the singular and shows Check messages first.
9. In **J5**, count work days late with `=NETWORKDAYS(INT(B5)+1,H5,'Days off'!$A$2:$A$30)` when **H5** is after **B5**, and 0 otherwise.
10. Fill **D5:J5** down to row 104.

New to WORKDAY? See [adding business days to a date](https://tidytabs.blogspot.com/2026/10/how-to-add-business-days-to-date-in.html), [planning backward from an event date](https://tidytabs.blogspot.com/2026/11/how-to-plan-backward-from-event-date.html), or, for bars, the [simple Gantt chart](https://tidytabs.blogspot.com/2026/10/how-to-make-simple-gantt-chart-with.html).

## Example

The queue starts Monday 11/23/2026 at 40 units a work day, with 11/26/2026, 11/27/2026, 12/24/2026 and 12/25/2026 on the Days off tab.

| Order | Promised | Units | Through | Work days | Start | Finish | Status | Work days late |
|---|---|---|---|---|---|---|---|---|
| W-101 | 11/25/2026 | 60 | 60 | 1 to 2 | 11/23/2026 | 11/24/2026 | On time | 0 |
| W-102 | 11/30/2026 | 100 | 160 | 2 to 4 | 11/24/2026 | 11/30/2026 | On time | 0 |
| W-103 | 12/02/2026 | 20 | 180 | 5 | 12/01/2026 | 12/01/2026 | On time | 0 |
| W-104 | 12/03/2026 | 140 | 320 | 5 to 8 | 12/01/2026 | 12/04/2026 | Late by 1 day | 1 |
| W-105 | 12/11/2026 | 80 | 400 | 9 to 10 | 12/07/2026 | 12/08/2026 | On time | 0 |
| W-106 | 12/16/2026 | 300 | 700 | 11 to 18 | 12/09/2026 | 12/18/2026 | Late by 2 days | 2 |
| W-107 | 12/23/2026 | 40 | 740 | 18 to 19 | 12/18/2026 | 12/21/2026 | On time | 0 |

W-102 runs from Tuesday 11/24/2026 to Monday 11/30/2026 because Thursday, Friday and the weekend are skipped. Orders share days: W-103 and W-104 both start 12/01/2026, and day 18 is split between W-106 and W-107.

The summary in **M4:M13** shows 7 orders, 740 units, 19 work days needed, a last finish date of 12/21/2026, 2 orders late and CHECK OK.

## Troubleshooting

### Every date is one work day late
The formula uses `WORKDAY($B$1,n)` instead of `WORKDAY($B$1-1,n)`. With that mistake W-101 finishes 11/25/2026 instead of 11/24/2026, and every date shifts the same way.

### An order finishes on a closed day
The holidays range is missing from WORKDAY, or the date is not on the Days off tab. Without the list, W-102 finishes on 11/26/2026 and W-103 on 11/27/2026, both closed days in the sample, and the last finish moves to 12/17/2026.

### The first order starts later than the date in B1
B1 is a weekend or a day off, so day 1 is the next work day. A Saturday start of 11/21/2026 gives day 1 on Monday 11/23/2026. A start on the day off 11/26/2026 gives day 1 on Monday 11/30/2026, and W-101 shows Late by 6 days.

### A row has no dates
The units are 0, negative, text or blank, so the row shows a Check message and **D:H** stay blank. If **B1** is not a date or **B2** is not above 0, **M4** shows Check and every row says Check: start date or units per day. A row with units but no order name is still scheduled and shows Check: no order name.

### A day off shows Check: date is text
The date was typed or pasted as text, such as "11/26/2026" from an email. In LibreOffice, WORKDAY still skipped that text date, but **M11** counts it and CHECK shows Check. Excel and Google Sheets were not checked, so retype it as a real date.

## Template

[Download the .xlsx](templates/tidy-tabs-production-queue-finish-dates.xlsx)

The file has three tabs. **Queue** holds **B1** and **B2** and up to 100 orders in rows 5 to 104, with inputs in **A:C** and the formulas from the Steps in **D:J**. The summary in **L4:M13** adds orders late, rows showing Check, days off typed as text, rows typed below row 104 and the CHECK.

**Days off** takes up to 29 closed days in **A2:A30** with the weekday and a Status for each; the sample dates are not a holiday list. **How to use** explains each formula.

The sheet assumes the same number of units every work day, made in the listed order. It does not plan staff, overtime or materials.

Yellow cells are inputs and gray cells are formulas. Every Check turns red and every Late by status turns orange. Both data tabs print Landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
