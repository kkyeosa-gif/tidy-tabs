---
threads_url: https://www.threads.com/@tin_ylab/post/DdzbGrqCREA
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-build-simple-weekly-timesheet.html
title: How to Build a Simple Weekly Timesheet Template in Excel and Google Sheets (with Overtime Calculation)
labels: timesheets, payroll-basics
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A printed weekly timesheet clipped to a clipboard resting on a break room table next to an old-style punch time clock and a wall clock, no readable text or numbers visible
image_alt: Clipboard timesheet and time clock sitting on a small business break room table
search_description: Build a weekly timesheet in Excel or Google Sheets that calculates hours worked and overtime pay automatically, with the exact formulas to copy in.
threads: Stop doing overtime math by hand. Set up one column with (Time Out - Time In)*24 minus your break, then wrap the weekly total in =MIN(total,40) for regular hours and the rest at 1.5x. Works the same in Excel and Google Sheets.
---
Calculate hours worked with `=(TimeOut-TimeIn)*24-Break`, then split the weekly total into regular and overtime with `=MIN(WeeklyTotal,40)` for regular hours and `=MAX(WeeklyTotal-40,0)` for overtime. Multiply each by the hourly rate, with overtime at 1.5x.

> Works in: Excel (Microsoft 365, Excel 2019+) and Google Sheets. Steps from Microsoft/Google help pages; not hands-on tested.

![Spreadsheet example with columns Date, Time In, Time Out, Break (hrs)](images/2026-09-27-zauto-01-how-to-build-a-simple-weekly-timesheet-t/how-to-build-a-simple-weekly-timesheet-template-in-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

1. In row 1, type these headers across columns A through H: **Date**, **Time In**, **Time Out**, **Break (hrs)**, **Hours Worked**, **Rate**, **Regular Pay**, **Overtime Pay**.
2. Select the Time In and Time Out columns and format them as Time.
3. Select the Break, Hours Worked, Rate, Regular Pay, and Overtime Pay columns and format them as Number.
4. In the Hours Worked column, type `=(C2-B2)*24-D2` and copy it down for each day.
5. Below the last day, add a row labeled **Weekly Total** and type `=SUM(E2:E6)` (adjust the range to match your five or seven rows).
6. In a **Regular Hours** cell, type `=MIN(WeeklyTotalCell,40)`.
7. In an **Overtime Hours** cell, type `=MAX(WeeklyTotalCell-40,0)`.
8. In **Regular Pay**, type `=RegularHours*Rate`. In **Overtime Pay**, type `=OvertimeHours*Rate*1.5`.

### In Excel

- Select the Time In and Time Out columns, then go to **Home > Number Format > Time**.
- To avoid negative results from a shift that crosses midnight, add 1 to the Time Out cell: `=(C2-B2+1)*24-D2`.

### In Google Sheets

- Select the Time In and Time Out columns, then go to **Format > Number > Time**.
- Same midnight fix applies: wrap the time difference with `+1` before multiplying by 24.

## Example

Sample data below is fictional, for a hypothetical part-time shop employee named Alicia.

| Date | Time In | Time Out | Break (hrs) | Hours Worked | Rate | Regular Pay | Overtime Pay |
|---|---|---|---|---|---|---|---|
| 09/22/2026 | 9:00 AM | 5:00 PM | 0.5 | 7.5 | $18.00 | | |
| 09/23/2026 | 9:00 AM | 5:00 PM | 0.5 | 7.5 | $18.00 | | |
| 09/24/2026 | 9:00 AM | 6:00 PM | 0.5 | 8.5 | $18.00 | | |
| 09/25/2026 | 9:00 AM | 6:00 PM | 0.5 | 8.5 | $18.00 | | |
| 09/26/2026 | 9:00 AM | 6:30 PM | 0.5 | 9.0 | $18.00 | | |
| **Weekly Total** | | | | **41.0** | | **$720.00** | **$27.00** |

Weekly total is 41 hours: 40 regular hours at $18.00 ($720.00) plus 1 overtime hour at $27.00 ($18.00 x 1.5).

## Troubleshooting

### Hours Worked shows a serial number instead of a decimal
The Hours Worked column is still formatted as Time. Select the column and set it back to Number so the multiplication by 24 shows a normal decimal like 7.5.

### Hours Worked is negative
This happens when a shift crosses midnight, so Time Out is technically earlier than Time In. Add `+1` day to the formula: `=(TimeOut-TimeIn+1)*24-Break`.

### Overtime Pay is a huge number
Check that Regular Hours uses `MIN(WeeklyTotal,40)`, not the raw weekly total. Without the cap, the formula multiplies all hours by the overtime rate.

### Weekly Total doesn't match the days listed
A hidden or filtered row can get skipped by `SUM`, or the range in the formula stops one row short after you added a day. Confirm the SUM range covers every visible day row.

### Time In or Time Out won't calculate at all
The cell was typed as plain text (like "9am" without a space) instead of a recognized time value. Retype it as `9:00 AM` with the colon and a space before AM/PM, or reformat the cell as Time first.

## Copy-paste setup

Headers (row 1, columns A-H):

```
Date | Time In | Time Out | Break (hrs) | Hours Worked | Rate | Regular Pay | Overtime Pay
```

Formulas (row 2, fill down):

- Hours Worked: `=(C2-B2)*24-D2`
- Weekly Total (below last row): `=SUM(E2:E6)`
- Regular Hours: `=MIN(E7,40)` (where E7 is Weekly Total)
- Overtime Hours: `=MAX(E7-40,0)`
- Regular Pay: `=Regular_Hours*F2`
- Overtime Pay: `=Overtime_Hours*F2*1.5`

Number formats: Time In and Time Out as Time, everything else as Number with two decimal places for pay columns.

Related help pages:
- [IF function (Microsoft)](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)
- [Calculate the difference between two times (Microsoft)](https://support.microsoft.com/en-us/office/calculate-the-difference-between-two-times-e1c78778-749f-4136-a052-2f7b1cf9a2f8)
- [Google Sheets function list](https://support.google.com/docs/table/25273)
