---
title: How to Prorate a Partial Month of Rent or a Retainer in Excel and Google Sheets
labels: excel-formulas, invoicing, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. The template was recalculated there and every result was checked against an independent calculation. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Prorate a partial month with =EOMONTH(A2,0)-A2+1 and =ROUND(B2*C2/D2,2): $150.00 from 10/12/2026 is 20 of 31 days, $96.77. Free .xlsx template.
image_prompts: Landlord or studio manager desk with a ring of metal keys, a wall calendar page with a mid month date circled in pencil and a stack of rent receipts, soft window light, no screens, no readable text, no logos
image_alt: Metal keys, a wall calendar with a circled date and a receipt stack
threads: $150.00 starting 10/12 is 20 of 31 days, so $96.77.\n=EOMONTH(A2,0)-A2+1 counts the days left, start day included. Then =ROUND(B2*C2/D2,2) does the charge.\nIf the days show up as 01/20/1900, format the cell as Number.
---
Count the days left in the month with `=EOMONTH(A2,0)-A2+1`, then charge `=ROUND(B2*C2/D2,2)`, where B2 is the monthly fee, C2 the days charged and D2 the days in that month. A $150.00 fee starting 10/12/2026 comes to 20 of 31 days, or $96.77.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against an independent calculation. The Excel and Google Sheets steps use the same functions. Neither app was opened for this post.

![Proration table with start date, monthly fee, days charged, days in month and prorated amount for seven sample rows](images/2026-10-10-team-03-prorate-first-month-subscription-excel-google-sheets/prorate-partial-month-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Use one row per customer or tenant. The start date goes in column A and the full monthly fee goes in column B.

### In Excel

1. In row 1, type these headers in A to E: Start date, Monthly fee, Days charged, Days in month, Prorated amount.
2. In **A2**, type the start date, for example 10/12/2026. In **B2**, type the full monthly fee, for example 150.
3. In **C2**, type `=EOMONTH(A2,0)-A2+1`. `EOMONTH(A2,0)` is the last day of the start month. Subtracting the start date and adding 1 counts the start day itself.
4. In **D2**, type `=DAY(EOMONTH(A2,0))`. This returns 28, 29, 30 or 31, depending on the month.
5. In **E2**, type `=ROUND(B2*C2/D2,2)`. This is the fee times days charged over days in the month, rounded to cents.
6. Select **C2:D2** and set **Home > Number Format** to Number. If they show a date such as 01/20/1900, this is why.
7. Format A as a date, and B and E as currency with two decimals.
8. Select **C2:E2** and drag the fill handle down for more rows.

Microsoft documents the arguments on its [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) pages.

### In Google Sheets

The same formulas should work as typed, since Sheets has EOMONTH, DAY and ROUND. Enter them in the same cells as above. For step 6, select **C2:D2** and choose **Format > Number > Number**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It covers a retainer, a farmers market booth and studio rent, including a short month and a leap year.

| Start date | Monthly fee | Days charged | Days in month | Prorated amount |
|---|---|---|---|---|
| 10/12/2026 | $150.00 | 20 | 31 | $96.77 |
| 02/20/2026 | $150.00 | 9 | 28 | $48.21 |
| 11/01/2026 | $150.00 | 30 | 30 | $150.00 |
| 12/31/2026 | $150.00 | 1 | 31 | $4.84 |
| 09/15/2026 | $85.00 | 16 | 30 | $45.33 |
| 03/10/2026 | $1,200.00 | 22 | 31 | $851.61 |
| 02/20/2028 | $150.00 | 10 | 29 | $51.72 |

Check the first row by hand: 10/31/2026 minus 10/12/2026 is 19, plus 1 is 20 days. Then $150.00 times 20 divided by 31 is 96.774..., shown as $96.77.

In the template these sit in rows 2 to 8 of the Proration tab. All seven rows matched an independent calculation using calendar days and rounding half up.

A start on the 1st returns the full fee, because every day of the month is charged. A start on the last day returns one day.

## Troubleshooting

### Days charged shows a date like 01/20/1900
Excel and Sheets can format the result of a date subtraction as a date. Select the C and D cells and change the format to Number. The value underneath was already correct, only the display was wrong.

### The result is one day short
The formula adds 1 so the start day counts. If you type `=EOMONTH(A2,0)-A2` your total leaves out the first day. Whether the start day should be charged is your call, so use the version that matches your agreement.

### The amount comes out as #VALUE!
The start date is probably text, not a real date. Retype it as 10/12/2026, or select the column and apply a date format. A left-aligned date usually means text.

### A February start gives an odd divisor
The template reads the real month length, so February is 28 days, or 29 in a leap year such as 2028. Check D if the number looks wrong. If D is hard-coded as 30, that cell no longer follows the calendar.

### Your lease or contract disagrees with the answer
This post is arithmetic only. Some leases and contracts count a flat 30 day month or use other rules, so follow your own agreement. This is not legal or landlord-tenant advice.

## Template

[Download the .xlsx](templates/tidy-tabs-prorate-partial-month.xlsx)

The file has two tabs. **Proration** holds the seven sample rows above with the three formula columns shaded, plus a short label for each case. **How to use** has fill-in notes and the day-count formulas.

The dates and fees are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
