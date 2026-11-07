---
title: How to Find the Last Business Day of Each Month in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 18 workbooks and 1,920 comparisons with 0 mismatches. Checked: the sample, an empty holiday list, a Saturday month end (11/30/2024), the 02/29/2028 leap day, a mid-month start, a holiday on a month end, two holidays in a row, three holidays over a weekend, a start date with a time, 120 months, four random sets of 36 months with 20 holidays, and a text holiday, a text start and a holiday with a time (all show Check). Excel and Google Sheets were not used. Not run: WORKDAY with a partly empty holiday range in Excel or Sheets, a non-English weekday name, date serials below 61.
search_description: Find the last business day of every month in Excel and Google Sheets with EOMONTH and WORKDAY, skipping your own closed days. Free template included.
image_prompts: A zippered bank deposit bag, a ring of shop keys and a blank hanging door sign on a wooden counter, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Zippered deposit bag and shop keys beside a blank hanging sign on a counter
threads: 12 month ends, one WORKDAY formula.\n=WORKDAY(B6+1,-1,Holidays!$B$2:$B$21) starts from the day after month end, so a month end that is already a business day stays put.\nIn the sample, 12/31/2026 is a closed day, so December lands on Wednesday 12/30 (fictional list).
---
Get the month end with `=EOMONTH(A6,0)`, then step back to the last business day with `=WORKDAY(B6+1,-1,Holidays!$B$2:$B$21)`, where the Holidays range is your own list of closed days. December 2026 ends on Thursday 12/31/2026, which is on the sample list, so the last business day is Wednesday 12/30/2026 (the closed days are fictional, not a legal list).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (18 workbooks, 1,920 comparisons). The formulas use only IF, ISNUMBER, DATE, YEAR, MONTH, EDATE, EOMONTH, WORKDAY and TEXT, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Twelve month ends with the last open weekday found and days moved back](images/2026-11-08-team-04-last-business-day-of-month-eomonth-workday-excel-google-sheets/last-business-day-of-month-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional.*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Months** tab and a **Holidays** tab. On **Months**, the first month is in **B2**, the headers are in row 5, and the 12 month rows run from row 6 to row 17.

1. On **Holidays**, type your closed days in **B2:B21** as real dates, such as 12/24/2026. Column **A** is for an optional name. Weekends are already skipped, so you only list weekday closures.
2. On **Months**, type any date in the first month you want in **B2**, such as 11/01/2026.
3. In **A6**, type `=DATE(YEAR($B$2),MONTH($B$2),1)` for the first day of that month. In **A7**, type `=EDATE(A6,1)` and fill down to row 17.
4. In **B6**, type `=EOMONTH(A6,0)`. The `0` means the same month, so you get its last calendar day. Format the column as a date.
5. In **C6**, type `=WORKDAY(B6+1,-1,Holidays!$B$2:$B$21)`. This is the last business day.
6. In **D6**, type `=TEXT(C6,"dddd")` to show the weekday name.
7. In **E6**, type `=B6-C6` to show how many days earlier than the month end it landed.
8. Fill **B6:E6** down to row 17. To cover more months, keep dragging row 17 down.

The trick is `B6+1`. `WORKDAY` with `-1` steps back one working day from the date you give it. Starting from the day after the month end means a month end that is already a business day stays put, and a Saturday or Sunday month end moves back to Friday.

The weekend is always Saturday and Sunday. `WORKDAY` has no way to change that. If you need other days off, Microsoft's `WORKDAY.INTL` takes a weekend setting, but this template does not use it.

Help pages: Microsoft's [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628) and [WORKDAY function](https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33). They describe the functions, not our test. We did not verify a Google help page for these functions.

## Example

Sample data is fictional. The first month is 11/01/2026, and the holiday list holds 10 made-up closed days: 11/26/2026, 12/24/2026, 12/25/2026, 12/31/2026, 01/01/2027, 01/18/2027, 05/31/2027, 07/05/2027, 09/06/2027 and 11/25/2027.

| Month end | Last business day | Weekday | Days earlier |
|---|---|---|---|
| 11/30/2026 | 11/30/2026 | Monday | 0 |
| 12/31/2026 | 12/30/2026 | Wednesday | 1 |
| 01/31/2027 | 01/29/2027 | Friday | 2 |
| 02/28/2027 | 02/26/2027 | Friday | 2 |
| 03/31/2027 | 03/31/2027 | Wednesday | 0 |
| 04/30/2027 | 04/30/2027 | Friday | 0 |
| 05/31/2027 | 05/28/2027 | Friday | 3 |
| 06/30/2027 | 06/30/2027 | Wednesday | 0 |
| 07/31/2027 | 07/30/2027 | Friday | 1 |
| 08/31/2027 | 08/31/2027 | Tuesday | 0 |
| 09/30/2027 | 09/30/2027 | Thursday | 0 |
| 10/31/2027 | 10/29/2027 | Friday | 2 |

Of the 12 months, 6 end on a business day and 6 move back. The most days any month moved is 3. May 2027 ends on Monday 05/31/2027, which is on the list, so it moves back past the weekend to Friday 05/28/2027.

## Troubleshooting

### Every row says Check
A date in the holiday list or the first-month cell is text or has a time of day, such as `12/24/2026 9:00 AM`. The template shows Check on purpose rather than skipping the bad holiday. Retype the date as a plain date. In the **Holidays** tab, column **C** marks the row that is wrong.

### A month end on a closed day did not move
Check that the closed day is inside **Holidays!$B$2:$B$21** and is a real date, not text. A text date in that range makes the template show Check. If you need more than 20 closed days, extend the range in every formula.

### A closed day on a weekend changed nothing
That is expected. Saturdays and Sundays are skipped already. Listing 12/25/2027, a Saturday, does not move December 2027 any further.

### The last business day is the month end, not earlier
That is correct when the month end falls on a weekday that is not on your list. 03/31/2027 is a Wednesday and stays 03/31/2027 with 0 days earlier.

### The weekday shows in another language
`TEXT(C6,"dddd")` follows the spreadsheet's language setting. We only ran it in English.

## Template

[Download the .xlsx](templates/tidy-tabs-last-business-day.xlsx)

The file has three tabs. **Months** has the first-month cell in **B2**, a settings check in **B3**, 12 month rows from row 6 with the month start, month end, last business day, weekday and days earlier, and a summary in **G5:H9**: months listed, month ends that are business days, months moved back, most days earlier, and a **CHECK** cell. **Holidays** has room for 20 closed days with 10 fictional ones filled in and a check column. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The holiday list is yours: the template does not know which days your bank, client or shop is closed, and it gives no legal holiday list.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
