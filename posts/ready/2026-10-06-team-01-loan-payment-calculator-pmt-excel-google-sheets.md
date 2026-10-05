---
title: How to Calculate a Loan Payment in Excel and Google Sheets
labels: excel-formulas, loans, small-business-finance
tested_in: LibreOffice Calc 24.2.7 (Linux). The Excel step follows Microsoft Support (not tested here). Google Sheets was not opened.
search_description: Calculate a monthly loan payment with =PMT(B2/12,B3*12,-B1), plus total paid and total interest. Free .xlsx template with four sample loans.
image_prompts: A workshop bench with a small bag of clay, a ceramic kiln shelf with two unfired bowls, and a paper loan folder with a pen and a pocket calculator, soft window light, all paper text blurred and unreadable, no screens, no logos
image_alt: Workshop bench with clay, unfired bowls, a paper folder, and a pocket calculator
threads: Put a minus sign in front of the amount and PMT stops returning a negative payment.\n=PMT(C2/12,D2*12,-B2) turns $15,000.00 at 6.50% over 5 years into $293.49 a month.\nTotal interest on that loan: $2,609.53.
---
Use `=PMT(B2/12,B3*12,-B1)` to get the monthly payment, where B1 is the amount, B2 the yearly rate and B3 the years. A $15,000.00 loan at 6.50% for 5 years comes to $293.49 a month.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated. The Excel step links to Microsoft Support. Google Sheets uses a PMT function with the same arguments, but neither Excel nor Google Sheets was opened for this post.

![Loans sheet listing four sample loans with amount, yearly rate, years, monthly payment, total paid, and total interest](images/2026-10-06-team-01-loan-payment-calculator-pmt-excel-google-sheets/loan-payment-calculator-template.png) *LibreOffice Calc 24.2 PDF export of the Loans sheet with sample data (fake).*

## Steps

The template keeps one loan per row. The first paragraph uses a single column of inputs, the same idea turned sideways. Here is the row layout the template uses.

1. In row 1, type these headers: Loan, Amount, Yearly rate, Years, Monthly payment, Total paid, Total interest.
2. In row 2, type a loan name in **A2**, the amount in **B2** (15000), the yearly rate in **C2** (6.5% or 0.065) and the years in **D2** (5).
3. In **E2**, type `=PMT(C2/12,D2*12,-B2)`. The rate is divided by 12 and the years are multiplied by 12 because you pay monthly.
4. Keep the minus sign before **B2**. It makes the payment come out positive. Without it, PMT returns a negative number.
5. In **F2**, type `=E2*D2*12` for the total paid over the life of the loan.
6. In **G2**, type `=F2-B2` for the total interest. That is everything you pay above the amount you borrowed.
7. Format B, E, F and G as currency, and C as a percentage with two decimals.
8. Fill row 2 down for each extra loan, or type a new term in **D2** to compare terms on the same amount.

### In Excel

The formula works as typed. Microsoft documents the arguments on its [PMT function](https://support.microsoft.com/en-us/office/pmt-function-0214da64-9a63-4996-bc20-214433fa6441) page.

### In Google Sheets

Google Sheets has a PMT function that takes the same rate, number of periods and present value. This post does not link a Google help page for it, and the formula was not run in Sheets.

## Example

Sample data below is fictional. It is four loans a small pottery and craft business might price out.

| Loan | Amount | Yearly rate | Years | Monthly payment | Total paid | Total interest |
|---|---|---|---|---|---|---|
| Equipment loan (fictional) | $15,000.00 | 6.50% | 5 | $293.49 | $17,609.53 | $2,609.53 |
| Kiln purchase | $8,000.00 | 7.25% | 3 | $247.93 | $8,925.56 | $925.56 |
| Booth trailer | $22,000.00 | 5.90% | 7 | $320.33 | $26,908.10 | $4,908.10 |
| Inventory line | $5,000.00 | 9.00% | 2 | $228.42 | $5,482.17 | $482.17 |

The first row checks out by hand. The monthly rate is 6.50% / 12, and there are 5 x 12 = 60 payments. The standard formula, amount x rate / (1 - (1 + rate)^-payments), gives 293.4922, which shows as $293.49.

Compare the first and third rows. The trailer payment is $320.33 against $293.49 for the equipment loan, but the trailer costs far more in interest over 7 years: $4,908.10 against $2,609.53. A longer term lowers the payment and raises the total interest.

These numbers are arithmetic only, not lending or financial advice. Real offers add fees and may use different rules.

## Troubleshooting

### The payment shows as a negative number
The amount was entered without the minus sign in front of it. PMT treats the loan as money you receive and the payment as money you pay out, so the result is negative. Use `-B2` in the formula, as in step 3.

### The payment is much too big or too small
The rate or the term was not converted to months. A yearly rate used as is, or years used without multiplying by 12, will give a wrong payment. This mistake was not reproduced in LibreOffice for this post. Check that the formula has `C2/12` and `D2*12`.

### The rate looks like 650%
A rate typed as 6.5 in a cell formatted as a percentage is read as 650%. The template notes say to type 6.5% or 0.065. This was not reproduced here, so compare the cell on screen with the percent you meant.

### The payment is a cent off the lender's statement
Lenders round each month's payment, and the template does not. The template notes warn about this. Expect small differences of a cent or so on a statement and use the sheet to compare options, not to settle a balance.

### Total paid does not match the loan paperwork
Total paid here is the monthly payment times the number of payments. Origination fees, late fees and extra payments are not in the template, so add those yourself if your lender charges them.

## Template

[Download the .xlsx](templates/tidy-tabs-loan-payment-calculator.xlsx)

The file has two tabs. **Loans** holds the four sample loans with the PMT, total paid and total interest formulas in columns E, F and G. **How to use** has short fill-in notes, including why the amount carries a minus sign.

The loans and rates are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
