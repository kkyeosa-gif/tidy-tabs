---
threads_url: https://www.threads.com/@tin_ylab/post/DfCZoHPFUvW
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-make-loan-payment-schedule-with.html
title: How to Make a Loan Payment Schedule With Interest and Principal in Excel and Google Sheets
labels: excel-formulas, small-business, loans
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model that rounds to cents. Checked: the sample loan, 0 percent, a 1 month loan, 60 months, $500.00 and $250,000.00 loans, 4.25 percent, a first payment on 01/31 and 02/29/2028, and bad inputs (rate typed as 7, months 61 or 12.5, amount 0, numbers typed as text). Excel and Google Sheets were not opened.
search_description: Make a loan payment schedule with interest and principal in Excel and Google Sheets using PMT, IPMT and PPMT. Free template with a balance column and a check cell.
image_prompts: A rubber-banded bundle of paper payment coupon booklets beside a brass desk calculator with a paper tape curling out and a small ceramic piggy bank on a wooden table, soft window light, no screens, no readable text, no logos
image_alt: Coupon booklets tied with a rubber band next to a brass calculator and a piggy bank
threads: How much of your $370.53 payment is actually interest?\nIPMT gives the interest, PPMT the principal, and the balance falls by the principal each month. Month 1 of a sample $12,000.00 loan: $70.00 interest, $300.53 principal.\nThe last payment is a few cents off. The post says why.
---
Split each payment into interest and principal with `=-IPMT($C$2/12,A8,$C$3,$C$1)` and `=-PPMT($C$2/12,A8,$C$3,$C$1)`, where C1 is the loan amount, C2 the yearly rate, C3 the number of months and A8 the payment number. Subtract each principal from the balance above it to get the balance that remains.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Schedule sheet listing date, payment, interest, principal and a shrinking balance for each month of a twelve thousand dollar sample loan](images/2026-10-28-team-01-loan-payment-schedule-ipmt-ppmt-remaining-balance-excel-google-sheets/loan-payment-schedule-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: `PMT` gives the same monthly payment for the whole loan. Each month, `IPMT` says how much of it is interest and `PPMT` says how much pays down the loan. The interest part shrinks, the principal part grows, and the balance falls to zero.

This is arithmetic only. It does not say whether to borrow, and it does not cover the tax treatment of interest.

### In Excel

1. Type labels in **A1:A4** (Loan amount, Yearly rate, Months, First payment date). In **C1** type 12000, in **C2** type 7% (the percent sign matters), in **C3** type 36, and in **C4** type 11/28/2026.
2. In row 7, type the headings: Payment no., Date, Payment, Interest, Principal, Balance after.
3. In **I1**, type the monthly payment: `=ROUND(-PMT($C$2/12,$C$3,$C$1),2)`.
4. In **A8**, type 1. In **A9**, type `=A8+1`.
5. In **B8**, type `=EDATE($C$4,A8-1)`. Format the column as a date.
6. In **D8**, type `=ROUND(-IPMT($C$2/12,A8,$C$3,$C$1),2)`.
7. In **C8**, type `=$I$1`.
8. In **E8**, type `=ROUND(C8-D8,2)`.
9. In **F8**, type `=ROUND($C$1-E8,2)`. In **F9** and below the balance comes from the row above: `=ROUND(F8-E9,2)`.
10. Select **A9:F9** and drag the fill handle down to payment 36. Format C, D, E and F as currency.

`PPMT` also gives the principal: `=-PPMT($C$2/12,A8,$C$3,$C$1)`. The template uses payment minus interest (step 8) so each row adds up to the cent.

Microsoft's pages: [PMT function](https://support.microsoft.com/en-us/office/pmt-function-0214da64-9a63-4996-bc20-214433fa6441) and [IPMT function](https://support.microsoft.com/en-us/office/ipmt-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f). They describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 9 above, typed exactly as written. Format the dates with **Format > Number > Date**.
2. Select **A9:F9**, then drag the fill handle down to payment 36.
3. Select columns C to F and choose **Format > Number > Currency**.

Google's help page: [PMT](https://support.google.com/docs/answer/3093185). It describes the function, not our test.

### The last payment

Rounding each month to cents leaves a few cents unaccounted for. The template fixes that in the last row. In the last row only, type **C36** as `=F35+D36`: the balance still owed plus that month's interest, so the final balance is exactly $0.00. For the sample, the last payment is $370.38 instead of $370.53.

## Example

Sample data is fictional: a $12,000.00 equipment loan at 7.00% over 36 months, first payment 11/28/2026.

| No. | Date | Payment | Interest | Principal | Balance |
|---|---|---|---|---|---|
| 1 | 11/28/2026 | $370.53 | $70.00 | $300.53 | $11,699.47 |
| 2 | 12/28/2026 | $370.53 | $68.25 | $302.28 | $11,397.19 |
| 3 | 01/28/2027 | $370.53 | $66.48 | $304.05 | $11,093.14 |
| 35 | 09/28/2029 | $370.53 | $4.29 | $366.24 | $368.23 |
| 36 | 10/28/2029 | $370.38 | $2.15 | $368.23 | $0.00 |

The summary beside the table shows a monthly payment of $370.53, total interest of $1,338.93, total paid of $13,338.93, and a last payment date of 10/28/2029. The CHECK cell **I6** says OK when the principal column adds up to the loan, the last balance is $0.00, and every payment number is present.

Interest is the textbook schedule rounded to cents each month. A real lender may compute from the actual balance and differ by a few cents.

## Troubleshooting

### The table is empty
Read **C5**. The template only fills the table when C5 says OK. Otherwise it shows the reason, such as "Type a number in each yellow cell" when the amount, rate or date was typed as text like "12k" or "7%".

### The interest looks 100 times too big
You typed 7 instead of 7% in the rate cell. The template catches it with "Rate must be 0% to 100%". Retype the rate with the percent sign.

### Months 61 or 12.5 gives a message
The template accepts a whole number of months from 1 to 60. Anything else shows "Months must be a whole number from 1 to 60" and leaves the table empty.

### The last payment is a few cents off
That is the rounding catch-up described above. With a tiny loan it shows more: in the test, $1.00 over 60 months at 12% gave a payment of $0.02 and a last payment of $0.17.

### A 0 percent rate shows $0.00 interest
That is correct. At 0% the template skips `PMT` and `IPMT` and divides the amount by the months. In the test, $12,000.00 over 36 months gave $333.33 payments and a last payment of $333.45.

### Dates after a 01/31 start jump around
`EDATE` counts from the first date each time, so 01/31/2027 gives 02/28, 03/31 and 04/30. The day is not stuck at 28 after February.

## Template

The download has two sheets. **Schedule** holds four yellow input cells (C1:C4), a table of up to 60 payments with blank rows past the term, and the summary with the CHECK cell. **How to use** explains each cell. It has no extra payments, fees or daily interest.

[Download the .xlsx](templates/tidy-tabs-loan-payment-schedule.xlsx)

To use it in Google Sheets, open a sheet and choose **File > Import > Upload**, then pick the file. The Google Sheets import was not tested.
