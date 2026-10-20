---
threads_url: https://www.threads.com/@tin_ylab/post/DeugF3JGayH
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-how-many-months-it.html
title: How to Calculate How Many Months It Will Take to Pay Off a Balance in Excel and Google Sheets
labels: excel-formulas, small-business, loans
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal month by month loop, with 24 edited payoff variants (0 percent rate, payment below the first month's interest, payment above the balance, blank and text inputs, a one cent edge case) and goal examples. No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Calculate months to pay off a balance with =ROUNDUP(NPER(B2/12,-B3,B1),0) in Excel and Google Sheets. Free payoff template with a month by month schedule.
image_prompts: A small tabletop pottery kiln with its lid propped open beside a rack of unglazed mugs and a folded canvas apron on a studio bench, soft natural window light, no screens, no readable text, no logos
image_alt: Open tabletop kiln beside a rack of unglazed mugs in a pottery studio
threads: NPER says 13.095 months. Nobody makes 0.095 of a payment.\n=ROUNDUP(NPER(B2/12,-B3,B1),0) gives 14 for $5,000.00 at 8.00% with $400.00 a month.\nPay less than one month of interest and NPER errors out, so the template shows "Payment too low" instead.
---
Find the number of months with `=ROUNDUP(NPER(B2/12,-B3,B1),0)`, where B1 is the balance, B2 the yearly interest rate and B3 the monthly payment. Round up because a partial last month still needs a payment.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and compared with a Python month by month loop. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Payoff sheet with a balance, rate and payment above a month by month schedule table](images/2026-10-20-team-01-months-to-pay-off-equipment-balance-nper-excel-google-sheets/months-to-pay-off-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

NPER returns how many payments it takes to bring a balance to zero. It needs three things in a fixed order: the rate per month, the payment as a negative number, and the balance.

1. In **B1**, type the balance, for example $5,000.00.
2. In **B2**, type the yearly interest rate, for example 8.00%. Type 0% if there is no interest.
3. In **B3**, type the monthly payment, for example $400.00.
4. In **B4**, type the date of the first payment, for example 11/01/2026. Type it; the sheet does not use TODAY.
5. In **B5**, type `=ROUNDUP(NPER(B2/12,-B3,B1),0)`.
6. In **B6**, type `=EDATE(B4,B5-1)` for the date of the last payment.

Three details trip people up:

- `B2/12` turns the yearly rate into a monthly rate.
- `-B3` is negative because NPER treats money you pay as outflow and the balance as money you received. Leave the minus sign off and the sign convention is wrong, so the result will not be the number of months.
- `ROUNDUP` matters. In LibreOffice, NPER returned 13.095 for the sample, which means 13 full payments and a small 14th.

### Add a guard for a payment that is too small

If the payment does not cover one month of interest, the balance never shrinks and NPER returns #NUM!. Test for that first:

`=IF(B3<=B1*B2/12,"Payment too low",ROUNDUP(NPER(B2/12,-B3,B1),0))`

The template adds more guards on top. A blank, text, zero or negative input shows "Check inputs" instead of a number.

### Find the last payment and the interest

The last payment is smaller than the regular one. The template reads it from a Schedule tab, then uses it for the totals:

- Total paid: `=B3*(B5-1)+B7`, where **B7** is the last payment.
- Interest paid: `=B8-B1`, where **B8** is the total paid.

### In Excel

The formulas work as typed. Microsoft describes the arguments on its [NPER function](https://support.microsoft.com/en-us/office/nper-function-240535b5-6653-4d2d-bfcf-b6a38151d815) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) pages.

### In Google Sheets

The same formulas should work as typed. Google lists NPER on its [function list](https://support.google.com/docs/table/25273). Sheets was not opened for this post, so confirm the result against the schedule below.

## Example

Sample data is fictional: a $5,000.00 kiln balance at 8.00% a year, paid at $400.00 a month starting 11/01/2026.

| Cell | Result |
|---|---|
| B5 Months to pay off | 14 |
| B6 Last payment date | 12/01/2027 |
| B7 Last payment | $38.19 |
| B8 Total paid | $5,238.19 |
| B9 Interest paid | $238.19 |

The Schedule tab shows where the 14 comes from. Interest each month is the balance times the yearly rate divided by 12, rounded to cents.

| Month | Start balance | Interest | Payment | End balance |
|---|---|---|---|---|
| 1 | $5,000.00 | $33.33 | $400.00 | $4,633.33 |
| 14 | $37.94 | $0.25 | $38.19 | $0.00 |

Months 2 to 13 follow the same pattern. Row 16 of the Schedule tab, month 15, is blank because nothing is left to pay.

**B10** shows OK when the schedule has exactly the NPER number of payments and ends at $0.00.

The template has a second block for saving toward a goal. A $3,000.00 booth deposit at $250.00 a month and a 2.00% yearly savings rate gives `=ROUNDUP(NPER(B15/12,-B14,0,B13),0)`, which is 12 months. Here the payment is still negative, the present value is 0 and the goal is the future value. **B17** checks it with FV and shows OK.

## Troubleshooting

### I get #NUM! instead of a number
The payment is not larger than one month of interest, so the balance can never reach zero. On $6,000.00 at 12.00%, one month of interest is $60.00. A $60.00 payment and a $59.00 payment both show "Payment too low" in the template, and $61.00 works.

### NPER returns a decimal such as 13.095
That is expected. NPER solves for a fraction of a month. Wrap it in `ROUNDUP(...,0)` as above. Exact cases come out clean: $4,800.00 at 0% and $400.00 gave 12, not 13, in LibreOffice.

### The months and the schedule disagree and B10 says CHECK
NPER does not round interest to cents, and the schedule does. They can split by one month when the payment is a few cents short of the balance plus the first month's interest. With $5,000.00 at 8.00%, a payment of $5,033.33 makes NPER say 2 months while the cent schedule finishes in 1. Trust the schedule, and treat the CHECK as a prompt to look.

### The schedule stops and B7 says Over 60 months
The Schedule tab has 60 rows. A $6,000.00 loan at 12.00% with a $61.00 payment takes 414 months. B5 still gives the number, but B7 to B10 say Over 60 months. Copy the last Schedule row down to extend it.

### My lender's statement is a few cents off
The sheet charges interest monthly on the declining balance, with each payment at the end of the month. A lender's own schedule can differ by cents, and in rare cases by a month. It also ignores fees, skipped payments and rate changes.

## Template

[Download the .xlsx](templates/tidy-tabs-months-to-pay-off.xlsx)

The file has three tabs. **Payoff** holds the yellow inputs in **B1:B4** and **B13:B15**, the results in **B5:B9** and **B16**, and the checks in **B10** and **B17**. **Schedule** has the 60 month table. **How to use** has short notes. All tabs print portrait on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. It is arithmetic only, not borrowing advice, and it does not follow any lender's APR rules. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
