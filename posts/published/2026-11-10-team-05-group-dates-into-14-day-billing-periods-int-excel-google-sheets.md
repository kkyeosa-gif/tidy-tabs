---
threads_url: https://www.threads.com/@tin_ylab/post/DfmjoDKluD6
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-group-dates-into-14-day-billing.html
title: How to Group Dates Into 14-Day Billing Periods in Excel and Google Sheets
labels: excel-formulas, small-business, freelancers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 33 workbooks and 5,712 comparisons. Checked: the sample, an empty sheet, the day before the start, the first and last day of a period, period lengths of 1, 7, 14 and 30, a length of 0, -7, 14.5, text and blank, a start date with a time, dates with a time of day, text dates, blank rows in the middle, a leap day, a year change, 200 consecutive days, and 14 random sets. Excel and Google Sheets were not used. Not run: a non-US language setting, more rows than the slots.
search_description: Number each date's 14-day period with INT and total entries and amounts per period with COUNTIF and SUMIFS in Excel and Google Sheets. Free template.
image_prompts: A pencil, an ink stamp and two stapled paper invoices on a desk pad with two weeks bracketed in pencil, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Pencil, ink stamp and stapled invoices on a desk pad with two weeks bracketed
threads: Want to group entries into 14-day billing periods?\n=INT((A5-$B$1)/$B$2)+1 numbers each date.\nWith an 11/02/2026 start, 11/15/2026 is period 1 and 11/16/2026 is period 2 (sample entries are made up).
---
Number each date's period with `=INT((A5-$B$1)/$B$2)+1`, where B1 is the date the first period starts and B2 is the number of days in a period. With an 11/02/2026 start and 14 days, 11/15/2026 is period 1 and 11/16/2026 is period 2. This groups dates only. It does not say when anyone must bill.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (33 workbooks, 5,712 comparisons). The formulas use only IF, AND, NOT, ISNUMBER, COUNT, COUNTIF, SUMIFS, SUM and INT, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Table of dated freelance entries with period numbers, start dates and end dates for 14-day periods](images/2026-11-10-team-05-group-dates-into-14-day-billing-periods-int-excel-google-sheets/billing-periods-14-day-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. Type the date the first period starts in **B1** and the days in each period in **B2**. Put the entries in columns **A** to **C** from row 5 (Date, What, Amount), with headers in row 4.

1. In **B1**, type the first day of period 1, such as `11/02/2026`. Use a real date, not text, and no time of day.
2. In **B2**, type `14`. Any whole number of 1 or more works, so 7 gives weekly periods and 30 gives 30-day periods.
3. In **D5**, type `=INT((A5-$B$1)/$B$2)+1`. A date is a number of days, so the subtraction counts the days since the start. INT drops the fraction, and the +1 makes the first period number 1 instead of 0.
4. In **E5**, type `=$B$1+(D5-1)*$B$2` for the day the period starts.
5. In **F5**, type `=E5+$B$2-1` for the day it ends. Period 1 runs from B1 through B1 + B2 - 1, so 11/02/2026 through 11/15/2026.
6. Fill **D5:F5** down as far as your entries go.
7. For totals, list period numbers 1, 2, 3 and so on in **H5** down. In **K5**, type `=COUNTIF($D$5:$D$204,H5)`. In **L5**, type `=SUMIFS($C$5:$C$204,$D$5:$D$204,H5)`. Fill both down.

The template wraps D5 in checks. A date before B1 shows Before start, and a text date, or a B1 or B2 that cannot be used, shows Check. The totals skip Check rows.

A date with a time of day counts by its date. In the test, 11/15/2026 23:59 is still period 1. A start date in B1 with a time shows Check, so type it without one.

Help pages: Microsoft's [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef), [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b). They describe the functions, not our test.

## Example

Sample data is fictional: 13 freelance entries, a first period starting 11/02/2026 and a length of 14 days.

| Date | Period no. | Period starts | Period ends |
|---|---|---|---|
| 10/30/2026 | Before start | blank | blank |
| 11/02/2026 | 1 | 11/02/2026 | 11/15/2026 |
| 11/15/2026 | 1 | 11/02/2026 | 11/15/2026 |
| 11/16/2026 | 2 | 11/16/2026 | 11/29/2026 |
| 12/20/2026 | 4 | 12/14/2026 | 12/27/2026 |

The totals table for the whole sample:

| Period | Starts | Ends | Entries | Amount |
|---|---|---|---|---|
| 1 | 11/02/2026 | 11/15/2026 | 4 | $655.00 |
| 2 | 11/16/2026 | 11/29/2026 | 3 | $435.00 |
| 3 | 11/30/2026 | 12/13/2026 | 3 | $585.00 |
| 4 | 12/14/2026 | 12/27/2026 | 2 | $145.00 |
| Before the first period | | | 1 | $120.00 |
| All amounts | | | 13 | $1,940.00 |

Periods 5 to 12 are empty. Periods 1 to 4 hold $1,820.00, and the $120.00 logo revision dated 10/30/2026 falls before the start. The two add up to the $1,940.00 total, and the CHECK cell shows OK.

For 12/20/2026, the days since 11/02/2026 are 48. 48 divided by 14 is 3.43, INT makes it 3, and +1 gives period 4.

## Troubleshooting

### A date shows Check
The date is probably text. In the test, even "11/03/2026" typed as text shows Check. Re-enter it, or see how to convert text to numbers. A left-aligned date is a hint that it is text.

### Every row shows Check
B1 or B2 cannot be used. In the test, a length of 0, a negative number, a fraction such as 14.5, text or a blank all gave Check. A start date with a time of day did too. Fix the setting and the rows recalculate.

### An entry shows Before start
Its date is earlier than B1. Those rows are counted in their own line under the table, so they are not lost. Move B1 earlier if they belong in period 1.

### Entries are missing from the table
The table covers 12 periods. A later date gets a period number above 12 and is counted in a line called After period 12. To extend it, copy the last table row down and adjust the three total formulas. With 14 days, 12 periods reach 04/18/2027 from this start.

### The period ends one day earlier than expected
Period 1 ends on B1 + B2 - 1, so a 14-day period starting 11/02/2026 ends 11/15/2026, not 11/16/2026. The 16th is the first day of period 2.

## Template

[Download the .xlsx](templates/tidy-tabs-billing-periods.xlsx)

The file has two tabs. **Entries** has the first period start in **B1**, the days in each period in **B2**, and 200 slots from row 5 to row 204 with 13 fictional sample entries. Columns **A** to **F** hold the date, what, amount, Period no., Period starts and Period ends. A table in **H4:L16** lists 12 periods with their start, end, entries and amount. Rows below it show periods 1 to 12, entries before the first period, entries after period 12, rows showing Check, all amounts and a CHECK cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The page prints on US Letter, landscape. The template groups dates only, and nothing in it says when a client should be billed.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
