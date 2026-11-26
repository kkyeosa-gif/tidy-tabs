---
title: How to Calculate Average Daily Balance in Excel and Google Sheets
labels: excel-formulas, small-business, cash-flow
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 226 workbooks and 334,480 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Average daily balance from a transaction list: a row per day, SUMIFS for each day's entries, a carried end-of-day balance, and the sum divided by the days.
image_prompts: A small business desk corner with a printed monthly bank statement folded in thirds, an open zippered bank deposit bag with a few bills inside, and a paper desk calendar page with several days circled in pencil, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Folded bank statement, open deposit bag with bills and a pencil-circled calendar page
threads: (Opening + ending) / 2 is not your average daily balance.\nIn the sample November account it gives $3,980.68. Adding up every calendar day's end-of-day balance and dividing by 30 gives $2,554.41.\nOne row per day, SUMIFS for that day's entries, carry the balance forward, then average.
---
Give every day of the period its own row, total that day's entries with `=SUMIFS(Transactions!$C$5:$C$1004,Transactions!$A$5:$A$1004,A8,Transactions!$D$5:$D$1004,"OK")`, carry the balance forward with `=ROUND(D8+B9,2)`, and divide the sum of the end-of-day balances by the number of days. In the made-up November account that starts at $3,250.00, the average daily balance is $2,554.41 over 30 days.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (226 workbooks, 334,480 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Thirty November days of end-of-day balances with the average and two shortcut results](images/2026-11-26-team-03-average-daily-balance-from-transaction-list-end-of-day-balance-sumifs-excel-google-sheets/average-daily-balance-template.png) *Daily tab of a fictional business checking account for 11/01/2026 to 11/30/2026: one row per calendar day with the day's net amount and end-of-day balance, and the summary with the average daily balance of $2,554.41 in green next to the (opening + ending) / 2 shortcut at $3,980.68 and the days-with-entries shortcut at $2,743.02. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

This is the plain method: the balance at the end of each calendar day, added up and divided by the days. Your bank or card issuer may calculate its average daily balance differently (some use the balance at the start of the day, some compound daily, some count fees or new purchases their own way), and its statement is the authority. Use this sheet to follow the math, not to replace the statement.

The steps are the same in Excel and Google Sheets.

1. On the **Daily** tab, type the period's first day in **B1** and last day in **B2**. The last day is included, and the period can run up to 92 days.
2. Type the opening balance in **B3**: the balance at the end of the day before the period starts. Any entry dated before the period belongs in that number, not in the list.
3. On the **Transactions** tab, type one entry per row from row 5: Date in **A**, Description in **B**, Amount in **C**, with deposits positive and payments negative. Any order works, and several entries on one date are added together.
4. Check column **D** on Transactions. It shows OK, a Check message, or "Outside the period (not counted)" for each row.
5. Back on **Daily**, column **A** lists every calendar day with `=IF($K$4<>"OK","",IF($B$1+ROWS($A$8:A8)-1>$B$2,"",$B$1+ROWS($A$8:A8)-1))`. Weekends and days with no entries get a row too.
6. **B8** totals that day's OK entries with the SUMIFS above, and **C8** counts them with COUNTIFS. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
7. **D8** starts the balance with `=IF(A8="","",ROUND($B$3+B8,2))`, and **D9** carries it with `=IF(A9="","",ROUND(D8+B9,2))` down the column. A day with no entries keeps the day before's balance. See the ROUND help from [Microsoft](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) or [Google](https://support.google.com/docs/answer/3093440).
8. Read the average daily balance in **K7**: `=IF(K4<>"OK","",ROUND(K6/K5,2))`, where **K6** is `=SUM(D8:D99)` and **K5** is the day count, `=IF(K4<>"OK","",B2-B1+1)`.
9. Read the rest of the summary in **J4:K17**: ending balance, the lowest end-of-day balance (`MIN`) and the first day it happened (`INDEX` and `MATCH`), days below zero, two shortcuts for comparison, and a final CHECK cell. See the MIN help from [Microsoft](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152) or [Google](https://support.google.com/docs/answer/3094017).
10. Optional: type an annual rate in **B4** and the days in a year your statement uses in **B5** (365, 360 or 366). **K15** then shows a rough interest figure with `=IF(OR(K4<>"OK",B4=""),"",ROUND(K7*B4/B5*K5,2))`. It is an estimate only; the statement's interest charge or earnings are what count.

**CHECK** in **K17** is OK only when the settings are valid, the Transactions CHECK is OK, and the ending balance equals the opening balance plus every OK entry.

To see the balance after each single entry instead of each day, use [How to Add a Running Balance Column](https://tidytabs.blogspot.com/2026/10/how-to-add-running-balance-column-in.html). To look ahead instead of back, see [How to Build a Weekly Cash Flow Forecast](https://tidytabs.blogspot.com/2026/10/how-to-build-weekly-cash-flow-forecast.html).

## Example

The account below is made up: a candle shop's business checking for 11/01/2026 to 11/30/2026, opening at $3,250.00, with 16 entries on 15 days (Etsy payouts, market deposits, rent, suppliers, card fees and a loan payment). Selected days from the **Daily** tab:

| Date | Day's net amount | End-of-day balance |
|---|---|---|
| 11/01/2026 | $0.00 | $3,250.00 |
| 11/02/2026 | $612.40 | $3,862.40 |
| 11/03/2026 | -$845.17 | $3,017.23 |
| 11/10/2026 | -$1,400.00 | $2,473.78 |
| 11/12/2026 | -$1,312.60 | $1,161.18 |
| 11/18/2026 | -$1,247.32 (2 entries) | $1,104.21 |
| 11/21/2026 | $726.50 | $1,830.71 |
| 11/25/2026 | -$350.00 | $2,435.01 |
| 11/28/2026 | $1,180.75 | $3,615.76 |
| 11/30/2026 | $1,095.60 | $4,711.36 |

The 30 end-of-day balances add up to $76,632.18, so the average daily balance is $76,632.18 / 30 = **$2,554.41**. The ending balance is $4,711.36, the lowest is $1,104.21 (first on 11/18/2026, held through 11/20), and no day is below zero.

The two shortcuts in the summary show how far off the quick methods land:

| Method | Result |
|---|---|
| Every calendar day (the ADB) | $2,554.41 |
| (opening + ending) / 2 | $3,980.68 |
| Only the 15 days with entries | $2,743.02 |

At 2.00% on a 365-day year, the rough interest in **K15** is $4.20 for the 30 days.

## Troubleshooting

### The average looks too high
You averaged the opening and ending balance. That ignores how long each balance lasted: in the sample, (opening + ending) / 2 is $3,980.68, while the real average is $2,554.41 because the balance sat at $1,161.18 on 11/12 and 11/13 and at $1,104.21 from 11/18 to 11/20. Add up every day's end-of-day balance instead.

### The average is off when you use the running balance column
A transaction list has rows only for days with entries, so averaging its balance column skips the quiet days. In the sample that gives $2,743.02 instead of $2,554.41. A day with no entries still counts at the day before's balance, which is why the Daily tab has a row for every calendar day.

### An entry is missing from the balance
Check its Status on Transactions. "Check: date has a time" and "Check: date blank or text" mean the Daily tab cannot match it to a day, so retype it as a plain date. "Outside the period (not counted)" means its date is before **B1** or after **B2**; an earlier entry belongs in the opening balance.

### Every day is blank and CHECK says Check
**K4** found a bad setting: the period is longer than 92 days, the end is before the start, a date has a time, the opening balance is blank, the rate is not between 0% and 100%, or **B5** is not 360, 365 or 366. Fix the yellow cell and the days fill back in.

### The interest does not match the statement
**K15** uses one simple formula on the day-end average. Issuers may use start-of-day balances, daily compounding, different day counts or their own rules for what goes in the balance. Treat the statement as correct and the sheet as a rough check.

## Template

[Download the .xlsx](templates/tidy-tabs-average-daily-balance.xlsx)

**Daily** holds the settings in **B1:B5**, one row per day for up to 92 days (rows 8 to 99) and the summary in **J4:K17**. **Transactions** takes 1,000 entries (rows 5 to 1004) with its own CHECK cell, which also catches rows typed below row 1004. **How to use** explains each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. The average daily balance is light green, days below zero are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
