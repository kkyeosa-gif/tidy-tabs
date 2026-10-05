---
threads_url: https://www.threads.com/@tin_ylab/post/DeICGr7gfaG
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-track-subscription-renewals-in.html
title: How to Track Subscription Renewals in Excel and Google Sheets
labels: excel-formulas, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Track subscription renewals with EDATE for the next renewal date, days left, and a 30-day Renews soon flag, plus monthly and annual cost. Free .xlsx template.
image_prompts: A home office desk corner with a wall calendar showing a few circled dates, a small stack of opened envelopes and a plain credit card beside a notepad with handwritten blurred lines, a potted succulent, no laptop, all text unreadable, no logos
image_alt: Wall calendar with circled dates above a stack of opened bills and a credit card
threads: =EDATE(B2,C2) gives the next renewal date, and =E2-TODAY() counts the days left.\nAn IF on that number flags anything due within 30 days.\nA 6 month plan charged 08/31 renews 02/28/2027, because February has no 31st. No error.
---
Find each renewal date with `=EDATE(B2,C2)`, which adds whole months to the last charge date, then count the days left with `=E2-TODAY()`. Add an IF on the days left to flag anything due within 30 days.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every formula and result below was recalculated and checked against hand math, including EDATE on a month-end date. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Subscriptions sheet showing next renewal dates, days left, and two rows flagged Renews soon](images/2026-10-02-team-02-subscription-renewal-tracker-excel-google-sheets/subscription-renewal-tracker-template.png) *LibreOffice Calc 24.2 PDF export of the Subscriptions sheet with sample data (fake). The cost columns are hidden so the dates stay readable.*

## Steps

Set up one tab with one row per subscription. The As of date lives in cell **K1** so the sample numbers stay put.

1. Use these headers in row 1: Subscription, Last renewed, Months per cycle, Cost per cycle, Next renewal, Days left, Status, Monthly cost, Annual cost.
2. For each subscription, type the date you were last charged in **B** (Last renewed). Type **1** in **C** for a monthly plan, **12** for a yearly plan, **6** for every six months. Type the charge in **D**.
3. In **J1**, type `As of date`. In **K1**, type your report date, for example 10/02/2026. Type `=TODAY()` instead if you want the sheet to count from the day you open it.
4. In **E2** (Next renewal), type `=IF(B2="","",EDATE(B2,C2))`. EDATE returns the same day of the month, C2 months later.
5. In **F2** (Days left), type `=IF(E2="","",E2-$K$1)`.
6. In **G2** (Status), type `=IF(F2="","",IF(F2<0,"Past due",IF(F2<=30,"Renews soon","OK")))`. Change the 30 if you want a longer or shorter warning window.
7. In **H2** (Monthly cost), type `=IF(D2="","",D2/C2)`. In **I2** (Annual cost), type `=IF(D2="","",D2*12/C2)`.
8. Fill E2:I2 down for every subscription row.
9. In **J2** type `Total monthly` and in **K2** type `=SUM(H2:H200)`. In **J3** type `Total annual` and in **K3** type `=SUM(I2:I200)`.
10. Format B, E, and K1 as dates and D, H, I, K2, and K3 as currency.
11. Select **G2:G200** and add a conditional formatting rule with the formula `=OR($G2="Renews soon",$G2="Past due")` to shade those rows.

When a subscription renews, type the new charge date in **B**. The Next renewal date moves forward on its own.

### In Excel

The formulas work as typed. For the shading rule, go to **Home > Conditional Formatting > New Rule > Use a formula to determine which cells to format**. Microsoft explains the volatile date function on its [TODAY function](https://support.microsoft.com/en-us/office/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9) page.

### In Google Sheets

The same formulas work as typed. For the shading rule, select G2:G200, choose **Format > Conditional formatting**, set the rule to **Custom formula is**, and paste the formula from step 11. Google documents the arguments on its [EDATE function](https://support.google.com/docs/answer/3092974) page.

## Example

Sample data below is fictional. The As of date is 10/02/2026.

| Subscription | Last renewed | Months | Cost | Next renewal | Days left | Status | Monthly | Annual |
|---|---|---|---|---|---|---|---|---|
| Domain name | 03/15/2026 | 12 | $18.00 | 03/15/2027 | 164 | OK | $1.50 | $18.00 |
| Email marketing plan | 09/15/2026 | 1 | $20.00 | 10/15/2026 | 13 | Renews soon | $20.00 | $240.00 |
| Online shop plan | 09/28/2026 | 1 | $10.00 | 10/28/2026 | 26 | Renews soon | $10.00 | $120.00 |
| Bookkeeping software | 11/05/2025 | 12 | $180.00 | 11/05/2026 | 34 | OK | $15.00 | $180.00 |
| Photo storage | 08/31/2026 | 6 | $30.00 | 02/28/2027 | 149 | OK | $5.00 | $60.00 |

Totals: $51.50 per month and $618.00 per year.

Two rows are worth a second look. Bookkeeping software is 34 days out, which is past the 30-day window, so it shows OK. Photo storage was charged on 08/31, and February has no 31st, so EDATE returns 02/28/2027 instead of an error. In LibreOffice Calc, `=EDATE(DATE(2026,1,31),1)` also returned 02/28/2026.

## Troubleshooting

### Next renewal shows a number like 46461
The cell has a number format instead of a date format. Select column E and change the format to Date.

### Next renewal shows #VALUE!
The Last renewed cell holds text that only looks like a date, often a pasted date left-aligned in the cell. Retype the date, or check that Months per cycle is a number and not text.

### Days left never changes
K1 holds a typed date, so the sheet counts from that day. Replace K1 with `=TODAY()` to count from today. Keep a typed date when you want a snapshot that stays put.

### Status says Past due for something you already paid
Last renewed still has the old charge date. Type the new charge date in column B.

### Monthly cost shows #DIV/0!
Months per cycle is blank or 0 on a row that has a cost. Type 1 for monthly, 12 for yearly.

### New subscriptions are missing from the totals
The totals stop at row 200. Extend the ranges in K2 and K3, for example `H2:H500`.

## Template

[Download the .xlsx](templates/tidy-tabs-subscription-renewal-tracker.xlsx)

The file has two tabs. **Subscriptions** holds the five sample rows with the formulas from the steps above, shades the formula columns (Next renewal through Annual cost), and highlights Renews soon and Past due rows. The As of date is in K1 and the monthly and annual totals are in K2 and K3. **How to use** has short fill-in notes.

The subscriptions and prices are fictional, and the file does not connect to any billing account. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
