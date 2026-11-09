---
title: How to Convert Cents to Dollars in a Payments Export in Excel and Google Sheets
labels: excel-formulas, small-business, payments
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent exact Decimal model over 27 workbooks and 8,067 cell comparisons. Checked: the sample, an empty sheet, 0, 1 and 5 cents, a refund, 10 digit amounts, a fraction of a cent, text, a number stored as text, a blank fee, a text fee, a negative fee, 200 rows, a payout one cent off, a blank or text payout, 14 random lists up to 200 rows, and dollars typed in the reverse block. Excel and Google Sheets were not used. Not run: a non-US language setting, more rows than the template holds.
search_description: Convert a payments export from cents to dollars with =ROUND(C4/100,2), add fee and net columns, and check the total against your payout. Free template.
image_prompts: A handheld card reader beside a small metal cash box and folded paper bags on a farmers market stall counter, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Card reader beside a small cash box and folded paper bags on a market counter
threads: Payments export full of cents?\n=ROUND(C4/100,2) turns 1850 into $18.50 and -2400 into -$24.00.\nGoing back, 19.99*100 can come out as 1998.9999999999998 without the ROUND (made-up payments).
---

Divide the amount in cents by 100 and round to 2 places with `=ROUND(C4/100,2)`. An amount of 1850 becomes $18.50, and a refund of -2400 becomes -$24.00.

To go back from dollars to cents, use `=ROUND(J4*100,0)`. Without the ROUND, 19.99*100 can come out as 1998.9999999999998.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with an exact Decimal model (27 workbooks, 8,067 comparisons). The formulas use only IF, ISNUMBER, ROUND and basic arithmetic, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Payments table converting amounts and fees from cents to dollars with net amounts and row checks](images/2026-11-10-team-04-convert-cents-to-dollars-payments-export-round-divide-100-excel-google-sheets/cents-to-dollars-payments-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

All payments, fees and the payout in this post are made up. The fees say nothing about what any payment processor charges.

## Steps

Payment exports often store money as whole cents so no one has to deal with decimals. The sample puts the export in columns A to D, with the amount in cents in C and the fee in cents in D.

1. Paste your export into the yellow cells. Keep the amount in cents in column C and the fee in cents in column D.
2. In E4, type `=ROUND(C4/100,2)`. This is the amount in dollars.
3. In F4, type `=ROUND(D4/100,2)` for the fee. The template also treats a blank fee as 0, so a payment with no fee still gets a net amount.
4. In G4, type `=ROUND(E4-F4,2)` for the net amount.
5. Format columns E to G as currency with 2 decimals.
6. Copy E4:G4 down to the last row of the export.
7. To convert dollars back to cents, type the dollars in J4 and use `=ROUND(J4*100,0)` in K4.

The ROUND in step 4 keeps tiny floating point leftovers out of the net column, so totals match what you see.

Help pages: Microsoft's [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and Google's [ROUND](https://support.google.com/docs/answer/3093440). They describe the function, not our test.

## Example

Sample payments (fictional). The template adds a Row check column that says OK or Check for each row.

| Cents | Fee (cents) | Amount | Fee | Net |
|---|---|---|---|---|
| 1850 | 84 | $18.50 | $0.84 | $17.66 |
| 12500 | 405 | $125.00 | $4.05 | $120.95 |
| 5 | blank | $0.05 | $0.00 | $0.05 |
| -2400 | blank | -$24.00 | $0.00 | -$24.00 |
| 4999 | 175 | $49.99 | $1.75 | $48.24 |
| 100000 | 3120 | $1,000.00 | $31.20 | $968.80 |
| 2175 | 94 | $21.75 | $0.94 | $20.81 |
| 899 | 56 | $8.99 | $0.56 | $8.43 |

The summary block on the right shows 8 payments converted, $1,200.28 in amounts, $39.34 in fees and $1,160.94 net. Type the payout your processor reported (the sample uses $1,160.94) in B1 and the net total minus the payout is $0.00, so CHECK shows OK.

Going back to cents, 19.99 gives 1999, 0.07 gives 7, 1250 gives 125000 and 4.35 gives 435.

## Troubleshooting

### A row shows Check and is missing from the totals
The amount is not a whole number of cents. A value like 1850.5, text, or a number stored as text (such as "1850" left-aligned in the cell) fails the row check. The template leaves that row out of every total so a bad row cannot hide in the sum. Fix the cell, or convert text to numbers first with [our text to numbers post](https://tidytabs.blogspot.com/2026/10/how-to-convert-text-with-dollar-signs.html).

### The fee column is blank for some payments
A blank fee counts as 0, so the net equals the amount. A fee cell with text in it shows Check.

### Refunds look wrong
They should not. A refund stays negative: -2400 cents is -$24.00, and the net is also -$24.00 when the fee is blank.

### Dollars times 100 is off by a tiny amount
Floating point math stores 19.99*100 as 1998.9999999999998. Always wrap it as `=ROUND(J4*100,0)`. Type dollars with at most 2 decimals. In LibreOffice, 1.005 gave 100 cents instead of 101 because of the same floating point issue, so the reverse block is only exact for amounts with 2 decimals or fewer.

### CHECK says Check but every row is OK
Compare B1 with the net total. The payout check only tells you whether your net total equals the number you typed. It does not say the export itself is right, and a payout that is one cent off shows Check. A blank B1 shows "No payout in B1".

## Template

[Download the .xlsx](templates/tidy-tabs-cents-to-dollars.xlsx)

The Payments tab has room for 200 payments (rows 4 to 203), 8 fictional samples, the fee and net columns, a Row check column, a small dollars to cents block and a summary with the payout check. Yellow cells are inputs and gray cells are formulas. A How to use tab explains each column.

To open it in Google Sheets, choose **File > Import > Upload**, then pick the file. The formulas were checked in LibreOffice Calc 24.2.7.2 only.
