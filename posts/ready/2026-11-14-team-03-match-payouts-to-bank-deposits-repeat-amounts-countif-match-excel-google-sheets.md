---
title: How to Match Payouts to Bank Deposits When the Same Amount Repeats in Excel and Google Sheets
labels: excel-formulas, small-business, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 19 workbooks and 41,933 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Pair each Square, Etsy or Stripe payout with its bank deposit, even when $412.50 repeats, using a COUNTIF copy number and MATCH. Free .xlsx template.
image_prompts: Three matching rubber-banded bundles of paper receipts lined up beside an open bank deposit slip book and a pen on a small shop counter, all print blurred and unreadable, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Three matching receipt bundles beside an open deposit slip book on a counter
threads: MATCH on the amount alone sends all three $412.50 payouts to the same deposit.\nGive each a copy number with =ROUND(C5*100,0)&"-"&COUNTIF($C$5:C5,C5) in both lists and MATCH on that, so the second payout finds the second deposit.\nSort both lists oldest first, or the copies pair out of order.
---
Give each amount a copy number with `=ROUND(C5*100,0)&"-"&COUNTIF($C$5:C5,C5)` in both lists, then pair each payout with its deposit using `=MATCH(F5,Bank!$F$5:$F$204,0)`, so the second $412.50 payout matches the second $412.50 deposit instead of the first. In the made-up sample with 3 days allowed, 7 of 10 payouts match, a $1,150.00 Stripe payout that lands 4 days later shows Check date, and 2 show Not in bank.

The file only tells you which rows to look at. It does not decide how anything is booked.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (19 workbooks, 41,933 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither was opened for this post; COUNTIF with a growing range, MATCH on text keys and ROUND of C5*100 are the open questions there.

![Three $412.50 payouts numbered 1 to 3 and each paired with its own deposit](images/2026-11-14-team-03-match-payouts-to-bank-deposits-repeat-amounts-countif-match-excel-google-sheets/match-payouts-to-bank-deposits-template.png) *Ten payouts with 3 days allowed: seven Matched, one Check date at 4 days and two Not in bank, with the summary at right. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has three tabs: **Payouts**, **Bank** and **How to use**. The template splits the one-cell key into a Cents column and a Copy no. column.

1. On the **Payouts** tab, type the most days a payout may take to reach the bank in **B1**. The sample uses 3.
2. Paste your payouts into **A5:C204**: payout date in A, source (Square, Etsy, Stripe) in B and amount in C.
3. On the **Bank** tab, paste your deposits into **A5:C204**: date in A, description in B and amount in C. Keep both lists oldest first.
4. In **D5** on both tabs, turn the amount into whole cents: `=ROUND(C5*100,0)`. $412.50 becomes 41250.
5. In **E5** on both tabs, number the copies: `=COUNTIF($D$5:D5,D5)`. The range grows as you copy down, so the first $412.50 gets 1, the second 2.
6. In **F5** on both tabs, build the match key: `=D5&"-"&E5`.
7. On **Payouts**, in **G5**, find the deposit row with the same key: `=IFERROR(MATCH(F5,Bank!$F$5:$F$204,0),"")`.
8. In **H5**, pick that deposit's date with `=INDEX(Bank!$A$5:$A$204,G5)`, and in **I5** subtract the payout date: `=INT(H5)-INT(A5)`.
9. In **J5**, set the status: `=IF(G5="","Not in bank",IF(OR(I5<0,I5>$B$1),"Check date","Matched"))`.
10. On **Bank**, in **G5**, flag deposits with no payout: `=IF(ISNA(MATCH(F5,Payouts!$F$5:$F$204,0)),"Not a payout","Matched")`.
11. Fill each formula down to your last row.

The template wraps these core formulas. **D5** shows Check for a text amount or fractions of a cent, **H5** shows Check when the matched deposit has no real date, and **J5** shows Check when a date is text or blank or **B1** is blank, text or negative.

Days allowed counts calendar days, so a Friday payout that lands on Monday is 3 days.

Help pages: Microsoft's [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [MATCH](https://support.google.com/docs/answer/3093378).

## Example

Sample data is fictional: 10 made-up payouts and 10 made-up deposits, with 3 days allowed.

| Payout date | Source | Amount | Copy no. | Bank date | Days to bank | Status |
|---|---|---|---|---|---|---|
| 11/02/2026 | Square | $412.50 | 1 | 11/03/2026 | 1 | Matched |
| 11/02/2026 | Etsy | $286.40 | 1 | 11/04/2026 | 2 | Matched |
| 11/03/2026 | Stripe | $318.60 | 1 | 11/05/2026 | 2 | Matched |
| 11/04/2026 | Square | $412.50 | 2 | 11/05/2026 | 1 | Matched |
| 11/05/2026 | Stripe | $1,150.00 | 1 | 11/09/2026 | 4 | Check date |
| 11/06/2026 | Square | $98.75 | 1 | 11/09/2026 | 3 | Matched |
| 11/09/2026 | Etsy | $199.99 | 1 | | | Not in bank |
| 11/10/2026 | Stripe | $75.00 | 1 | | | Not in bank |
| 11/10/2026 | Square | $498.00 | 1 | 11/11/2026 | 1 | Matched |
| 11/11/2026 | Square | $412.50 | 3 | 11/12/2026 | 1 | Matched |

The three $412.50 Square payouts (11/02, 11/04, 11/11) pair with the three $412.50 deposits in order, 1 day each. The Friday 11/06 payout of $98.75 lands Monday 11/09, 3 days, so it is Matched.

On the **Bank** tab, the $640.00 cash deposit from the market and an Etsy deposit of $199.90 show Not a payout. That $199.90 deposit and the $199.99 payout are 9 cents apart, so the file does not pair them.

The summary in **L4:M10** shows 7 payouts matched, 1 to check (date), 2 not in bank, $3,589.25 found in the bank, $274.99 not in bank, $839.90 of deposits that are not payouts, and CHECK OK.

With the same sample, 1 day allowed gives 4 matched and 4 check date, and 4 days gives 8 matched and 0 check date.

## Troubleshooting

### Every repeat matches the same deposit
MATCH on the amount alone always finds the first one. In LibreOffice, `=MATCH(A20,$B$20:$B$22,0)` on three $412.50 rows returned 1 for all three. The copy number key returns 1, 2 and 3.

### A payout shows Check date though it arrived
Either it took longer than Days allowed (the Stripe $1,150.00 took 4 days) or the lists are not in the same order. Copies pair in the order listed, so with the payouts out of date order, a $412.50 on 11/09 listed before one on 11/02 got -6 and 8 days, both Check date. Sort both lists oldest first.

### A payout is Not in bank and a deposit is Not a payout
The amounts differ by a few cents, often a fee taken out. $199.99 and $199.90 do not pair. Amounts must match to the cent, and the file does not guess.

### A row shows Check
The amount is text, such as "$412.50" pasted with the dollar sign, or has fractions of a cent, such as 12.345. A date can also be text or blank, or **B1** is not a number of 0 or more.

## Template

[Download the .xlsx](templates/tidy-tabs-match-payouts-to-deposits.xlsx)

**Payouts** has Days allowed in **B1**, inputs in **A:C** (Payout date, Source, Amount) for rows 5 to 204, formulas in **D:J** (Cents, Copy no., Match key, Bank row, Bank date, Days to bank, Status) and the summary in **L4:M10**. **Bank** has Date, Description and Amount in **A:C**, then Cents, Copy no., Match key and Status in **D:G**. **How to use** explains each column. It prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Matched is green, Check date is orange, and Not in bank, Not a payout and Check are red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
