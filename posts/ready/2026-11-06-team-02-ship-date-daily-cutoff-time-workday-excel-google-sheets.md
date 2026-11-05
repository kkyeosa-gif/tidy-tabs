---
title: How to Calculate a Ship Date With a Daily Cutoff Time in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 32 workbooks and 12,992 comparisons with 0 mismatches. Checked: the sample, no orders, orders at 12:00 AM, 9:00 AM, 1:59:59 PM, 2:00:00 PM, 2:00:01 PM and 11:59:59 PM on weekdays and weekends, cutoffs of 12:00 AM, 11:59:59 PM, 12:00 PM and 9:30:15 AM, 0, 1, 3 and 20 shipping days, closed days in a week, on a weekend, listed twice, none and 20 of them, bad settings (text cutoff, cutoff of 1 or -0.1, days of 1.5, -1, text or blank, text in the closed-day list), a time-only and a text order cell, year-end and leap-year dates, 100 rows, and six random sets of 5 to 100 orders. Excel and Google Sheets were not used. Not run: WORKDAY in Excel or Sheets, TODAY, a non-English language setting.
search_description: Find a ship date with a daily cutoff time in Excel and Google Sheets using WORKDAY, MOD and INT. Free template skips weekends and closed days.
image_prompts: A round analog wall clock with plain blank hands near two o'clock, a taped brown parcel with a blank white label and a roll of packing tape on a shelf above a small home packing table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Wall clock near two o'clock above a taped parcel and packing tape roll
threads: With a 2:00 PM cutoff, an order at 2:00:00 PM counts today and one at 2:00:01 PM counts tomorrow.\nWORKDAY and a closed-day list do the rest: 5:05 PM on Wednesday 11/25/2026 ships Monday 11/30/2026.\nThe free template skips weekends and the closed days you list.
---
Find the day an order counts with `=IF(ROUND(MOD(B7,1)*86400,0)<=ROUND($B$2*86400,0),WORKDAY(INT(B7)-1,1,Holidays!$A$5:$A$24),WORKDAY(INT(B7),1,Holidays!$A$5:$A$24))`, then add the shipping days with `=WORKDAY(C7,$B$3,Holidays!$A$5:$A$24)`. With a 2:00 PM cutoff, an order at 5:05 PM on Wednesday 11/25/2026 ships Monday 11/30/2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every date below was recalculated and compared with a Python model. WORKDAY, MOD, INT and ROUND exist in current Excel and Google Sheets, but neither app was opened for this post.

![Order times with the day each counts and ship date after a 2 PM cutoff](images/2026-11-06-team-02-ship-date-daily-cutoff-time-workday-excel-google-sheets/ship-date-with-cutoff-time-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. On **Orders**, **B2** holds the cutoff time, **B3** holds the business days to ship, and orders start in row 7. On **Holidays**, the days you are closed go in **A5:A24**.

1. In **B2**, type your daily cutoff, such as `2:00 PM`. An order at or before this time counts that day.
2. In **B3**, type the business days it takes you to ship after the order counts. Type 1 to ship the next business day. Type 0 to ship the same day the order counts.
3. On **Holidays**, type each closed date in **A5:A24**, such as `11/26/2026`. Weekends are skipped on their own.
4. On **Orders**, type the order number in **A7** and the order date and time in **B7**, such as `11/25/2026 5:05 PM`. Format **B7** as a date with a time.
5. In **C7**, type the formula from the top of this post. It gives the day the order counts.
6. In **D7**, type `=WORKDAY(C7,$B$3,Holidays!$A$5:$A$24)`. This is the ship date.
7. Format **C7** and **D7** as dates. Fill **C7:D7** down for every order.

Here is how the formula in **C7** works. `MOD(B7,1)` is the time of day, because a date and time is stored as a number and the part after the decimal point is the time. Multiplying by 86400 turns it into seconds, and `ROUND` removes tiny floating point noise so 2:00:00 PM equals the cutoff exactly.

If the order is at or before the cutoff, `WORKDAY(INT(B7)-1,1,...)` finds the first business day on or after the order date. `INT(B7)` drops the time, and stepping back one day first means an order on a business day stays on that day. If the order is after the cutoff, `WORKDAY(INT(B7),1,...)` moves it to the next business day.

Two optional columns: **E7** `=IF(ISNUMBER(D7),TEXT(D7,"ddd"),D7)` shows the day name, and **F7** `=IF(ISNUMBER(D7),D7-INT(B7),D7)` shows calendar days after the order date.

The template also checks your settings. If the cutoff is not a time, the days are not a whole number, or a closed-day cell is text, **B4** says Check and so does every row.

The template does not know carrier pickup times, weekend shipping or delivery time. It gives the day you plan to ship, not the day the package arrives.

Help pages: Microsoft's [WORKDAY function](https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional. The cutoff is 2:00 PM, shipping takes 1 business day, and the closed days are 11/26/2026, 12/25/2026 and 01/01/2027.

| Order | Placed | Counts as | Ships | Days after |
|---|---|---|---|---|
| 5101 | 11/02/2026 9:15 AM | 11/02/2026 | 11/03/2026 Tue | 1 |
| 5103 | 11/02/2026 2:00:00 PM | 11/02/2026 | 11/03/2026 Tue | 1 |
| 5104 | 11/02/2026 2:00:01 PM | 11/03/2026 | 11/04/2026 Wed | 2 |
| 5105 | 11/05/2026 3:30 PM | 11/06/2026 | 11/09/2026 Mon | 4 |
| 5107 | 11/07/2026 10:00 AM (Saturday) | 11/09/2026 | 11/10/2026 Tue | 3 |
| 5109 | 11/25/2026 8:30 AM | 11/25/2026 | 11/27/2026 Fri | 2 |
| 5110 | 11/25/2026 5:05 PM | 11/27/2026 | 11/30/2026 Mon | 5 |
| 5111 | 11/26/2026 11:20 AM (closed) | 11/27/2026 | 11/30/2026 Mon | 4 |
| 5112 | 12/24/2026 3:00 PM | 12/28/2026 | 12/29/2026 Tue | 5 |

An order at exactly 2:00:00 PM counts that day, and one second later counts the next business day. Order 5109 skips Thursday 11/26/2026 because it is a closed day. Order 5112 skips Friday 12/25/2026 for the same reason.

The summary on the template shows 12 orders, a latest ship date of 12/29/2026 and a longest wait of 5 calendar days. Only the 9 orders above are shown here. The other three are in the download.

## Troubleshooting

### Every row says Check
Look at **B4** first. The cutoff in **B2** must be a time from 12:00 AM up to but not including 12:00 AM the next day, and **B3** must be a whole number of 0 or more. Every cell in **Holidays A5:A24** must be a date or empty. In the test, a text cutoff, a cutoff of 1, days of 1.5 or -1, and text in the closed-day list each made B4 and all rows say Check.

### One order says Check
The order cell holds only a time, or text instead of a date and time. In the test, both gave Check for that row. Retype the order as a date and a time.

### A closed date made no difference
A date on a Saturday or Sunday has no effect, because weekends are already skipped. A date typed as text is caught by the settings check instead.

### An order at exactly the cutoff ships a day late
The test is "at or before," so 2:00:00 PM counts on the same day. A cutoff of 12:00 AM only counts orders placed at exactly midnight.

### The day name is not in English
`TEXT(D7,"ddd")` shows English day names in an English setting. It was not tested in any other language setting.

## Template

[Download the .xlsx](templates/tidy-tabs-ship-date-cutoff-time.xlsx)

The file has three tabs. **Orders** has the cutoff in **B2**, the shipping days in **B3**, a settings check in **B4**, and 100 order rows from row 7 with 12 fictional sample orders and yellow input cells. A summary in **I6:I10** shows orders listed, the latest ship date, the longest wait, rows showing Check and a **CHECK** cell. **Holidays** has 20 yellow slots in **A5:A24** and fills in the day name beside each date. **How to use** has notes on each formula.

Fix every Check before you quote ship dates. The template has 100 order rows. A longer list was not tested.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
