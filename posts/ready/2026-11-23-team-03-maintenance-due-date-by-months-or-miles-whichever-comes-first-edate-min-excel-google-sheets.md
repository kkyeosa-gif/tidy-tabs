---
title: How to Track Maintenance Due by Months or Miles, Whichever Comes First, in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 18 workbooks and 6,426 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Find the month limit with EDATE, project the miles limit from use per day, and take the earlier date with MIN to see what service is due first.
image_prompts: A white delivery van parked in a small business driveway with its hood open, a quart bottle of motor oil standing on the front bumper and a blank service reminder sticker in the upper corner of the windshield, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Delivery van with open hood and a motor oil bottle on the bumper
threads: "6 months or 5,000 miles, whichever comes first" is easy to say and annoying to track.\nEDATE gives the month date, use per day projects the miles date, and =MIN(by_months,by_usage) takes the earlier one.\nThe sample van is due 01/06/2027 by miles, four days before its month limit.
---
Get the month limit with `=EDATE(last_service,months)`, the usage limit with `=reading_date+INT((last_reading+units-latest_reading)/units_per_day)`, where units per day is the use since service divided by the days since service, and the due date with `=MIN(by_months,by_usage)`. In the made-up van that is serviced every 6 months or 5,000 miles, last done 07/10/2026 at 48,210 miles and reading 51,890 on 11/20/2026, the miles hit first, so it is due 01/06/2027 instead of 01/10/2027, 44 days after the 11/23/2026 as-of date.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (18 workbooks, 6,426 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Seven equipment tasks with due dates, limit that hits first and days left](images/2026-11-23-team-03-maintenance-due-date-by-months-or-miles-whichever-comes-first-edate-min-excel-google-sheets/maintenance-due-tracker-template.png) *Equipment tab as of 11/23/2026: the delivery van oil change is due 01/06/2027 by mileage (44 days left) instead of 01/10/2027 by months, the sewing machine service and van tire rotation show Due soon, and the walk-in cooler coil cleaning is Overdue by 8 days. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

"Miles" here stands for any unit you count: miles on a van, firings on a kiln, shots on an espresso machine, hours on a sewing machine. Take the intervals from your owner's manual or the maker's schedule. The sample rules below are made up, and the sheet makes no safety or warranty claim.

The steps are the same in Excel and Google Sheets.

1. On the **Equipment** tab, type the as-of date in **B1**, such as 11/23/2026, and the number of days that count as Due soon in **B2**, such as 14.
2. Type one task per row from row 5: Item and task in **A**, Every N months in **B**, Every N units in **C**, Last service date in **D**, Reading at last service in **E**, Latest reading in **F** and Latest reading date in **G**. Leave **B** or **C** blank if the task has only one limit.
3. In **H5**, get the month limit with `=EDATE(D5,B5)`, the same day N months later. A service on 01/31 plus 1 month comes back 02/28. See the EDATE help from [Microsoft](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) and [Google](https://support.google.com/docs/answer/3092974).
4. In **I5**, get the use per day with `=(F5-E5)/(G5-D5)`. For the van that is 3,680 miles over 133 days, 27.67 a day.
5. In **J5**, project the usage limit with `=G5+INT((E5+C5-F5)/I5)`: the reading date plus the units left divided by the use per day. INT rounds down, so the date lands a day early rather than late. If the latest reading is already past the limit, the template returns the reading date.
6. In **K5**, take the earlier limit with `=MIN(H5,J5)`. A task with only one limit leaves the other cell blank, and MIN uses the one date that is there. **L5** says which limit won: Months or Usage.
7. In **M5**, count the days left with `=K5-$B$1`. Format **M** as a number. With a date format, 44 would show as a date in 1900.
8. Fill **H5:N5** down to row 54. **N** shows the Status: Overdue below 0 days, Due soon from 0 to the days in **B2**, OK above that.

**Why the as-of date is typed:** the template does not use TODAY. You type the day you wrote down the readings, so Days left matches the readings you have and stays the same on a printout. When you take new readings, update **F**, **G** and **B1** together.

The usage date is a straight-line projection from the use so far. It moves when the van is driven more or less, so treat it as an estimate and refresh it with each new reading.

## Example

The sample data is made up. As-of date 11/23/2026, Due soon 14 days.

| Item and task | Rule | By months | By usage | Due date | Limit | Days left | Status |
|---|---|---|---|---|---|---|---|
| Delivery van: oil change | 6 mo / 5,000 mi | 01/10/2027 | 01/06/2027 | 01/06/2027 | Usage | 44 | OK |
| Kiln: element check | 12 mo / 200 firings | 03/02/2027 | 01/21/2027 | 01/21/2027 | Usage | 59 | OK |
| Espresso machine: descale | 3 mo / 1,500 shots | 12/15/2026 | 12/28/2026 | 12/15/2026 | Months | 22 | OK |
| Sewing machine: service | 12 mo / 400 hours | 12/01/2026 | 05/26/2027 | 12/01/2026 | Months | 8 | Due soon |
| Box truck: tires | 24 mo | 01/31/2027 | | 01/31/2027 | Months | 69 | OK |
| Delivery van: tire rotation | 6 mo / 7,500 mi | 11/30/2026 | 01/06/2027 | 11/30/2026 | Months | 7 | Due soon |
| Walk-in cooler: coil cleaning | 3 mo | 11/15/2026 | | 11/15/2026 | Months | -8 | Overdue |

The van shows both sides of the rule: the oil change is due by miles, four days before its month limit; the tire rotation, last done 05/31/2026, is due by months on 11/30/2026, the last day of November.

The summary in **P4:Q10** shows 7 items listed, 1 Overdue, 2 Due soon, 0 rows showing Check and CHECK OK.

## Troubleshooting

### 01/31 plus 1 month shows 02/28
EDATE moves to the last day of a shorter month. In LibreOffice Calc, 01/31/2026 plus 1 month gave 02/28/2026 and 08/31/2026 plus 6 months gave 02/28/2027.

### Due by usage is blank
There has been no use since the service, or the latest reading date is the same as the service date, so there is no use per day to project from. A task with a month limit falls back to the months; a usage-only task shows No due date yet.

### The due date is the reading date and the status is Overdue
The latest reading is already past the usage limit. With 65,400 miles against 60,000 plus 5,000, the template returned the reading date 11/10/2026, -13 days, Overdue.

### Status says Check: reading lower than last service
An odometer typo, such as 4,800 typed for 48,000. Fix the reading in **E** or **F**.

## Template

[Download the .xlsx](templates/tidy-tabs-maintenance-due-tracker.xlsx)

The file has two tabs. **Equipment** holds up to 50 tasks in rows 5 to 54: you type **B1:B2** and **A:G**, and **H:N** hold the formulas above. The summary in **P4:Q10** counts Overdue, Due soon and Check rows, flags anything typed below row 54 and ends in a CHECK cell. **How to use** explains each formula. Equipment prints Landscape on US Letter.

Yellow cells are inputs and gray cells are formulas. Overdue turns orange, Due soon light blue, and every Check turns red. This is a reminder list, not a mileage log for taxes; for trips, see [How to Build a Simple Mileage Log Template](https://tidytabs.blogspot.com/2026/09/how-to-build-simple-mileage-log.html).

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
