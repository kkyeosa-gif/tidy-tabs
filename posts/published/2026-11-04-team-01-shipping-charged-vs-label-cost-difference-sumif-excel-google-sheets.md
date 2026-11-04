---
threads_url: https://www.threads.com/@tin_ylab/post/DfUiDvREaH7
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-check-that-your-shipping-charges.html
title: How to Check That Your Shipping Charges Cover Your Label Costs in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 16 workbooks and 6,560 comparisons with 0 mismatches. Checked: an empty log, text in a charged amount, a blank label cost, an order number with no amounts, all-zero amounts, half-cent amounts, a tie for the biggest shortfall, negative amounts, a log with no shortfall, and six random logs of 1 to 200 orders. Excel and Google Sheets were not used.
search_description: Compare the shipping you charged with each label cost in Excel and Google Sheets. Free template shows orders where shipping fell short and the total lost.
image_prompts: A small digital postal scale holding a kraft mailer box beside a roll of blank white label paper and a roll of packing tape on a wooden packing table, soft daylight, no screens, no readable text, no logos
image_alt: Postal scale with a kraft mailer box beside a roll of blank label paper
threads: Does the shipping you charge cover the label? Most of us never check.\nSubtract per order: =ROUND(C5-D5,2). Negative means you paid the difference. Add up only the negatives with =ROUND(-SUMIF(E5:E204,"<0"),2).\nIn the sample, 9 of 24 orders came up short.
---
Subtract the label cost from the shipping you charged with `=ROUND(C5-D5,2)`. A negative result means the buyer's shipping did not cover the label. Add up only the shortfalls with `=ROUND(-SUMIF(E5:E204,"<0"),2)`.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. ROUND, SUM, SUMIF and COUNTIF exist in current Excel and Google Sheets, but neither app was opened for this post.

![Orders table with shipping charged, label cost, difference and status, and a summary of shortfalls](images/2026-11-04-team-01-shipping-charged-vs-label-cost-difference-sumif-excel-google-sheets/shipping-charged-vs-label-cost-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. Put one order on each row, starting at row 5, with the order number in **A**, the date in **B**, the shipping you charged in **C** and the label cost in **D**.

1. In **E5**, type `=ROUND(C5-D5,2)`. This is the shipping you charged minus the label.
2. In **F5**, type `=IF(E5<0,"Short",IF(E5=0,"Even","Covered"))`.
3. Fill **E5:F5** down as far as your orders go.
4. In **I8**, type `=ROUND(SUM(C5:C204)-SUM(D5:D204),2)` for the net result across all orders.
5. In **I9**, type `=COUNTIF(E5:E204,"<0")` for the number of orders that fell short.
6. In **I10**, type `=ROUND(-SUMIF(E5:E204,"<0"),2)`. SUMIF adds only the negative differences, and the minus sign makes the total positive.
7. In **I11**, type `=INDEX(A5:A204,MATCH(MIN(E5:E204),E5:E204,0))` for the order with the biggest shortfall.

A free-shipping order has 0.00 in column C, so its whole label cost shows as a shortfall. That is the point: the sheet shows what free shipping costs you.

The sheet compares shipping charged with the label price only. It does not include packaging, marketplace fees or your time.

Help pages: Microsoft's [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b), [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/03/2026; the pages were not read for this post.

## Example

Sample data is fictional: 24 orders from 10/19/2026 to 11/03/2026.

| Order # | Shipping charged | Label cost | Difference | Status |
|---|---|---|---|---|
| 1041 | $5.00 | $4.35 | $0.65 | Covered |
| 1042 | $5.00 | $7.82 | -$2.82 | Short |
| 1043 | $0.00 | $6.15 | -$6.15 | Short |
| 1044 | $8.95 | $8.95 | $0.00 | Even |
| 1047 | $12.00 | $10.55 | $1.45 | Covered |

Across all 24 orders, buyers paid $141.80 for shipping and the labels cost $160.07. The net is -$18.27. Nine orders came up short for a total of $29.97, and the biggest was order 1043 at -$6.15, a free-shipping order.

## Troubleshooting

### A row shows Check
The charged amount or the label cost is blank or typed as text, such as `five`. In the test, a text amount and a blank label cost each gave Check on that row and left the **CHECK** cell saying Check.

### The net does not match your payout
The sheet compares two columns only. Marketplace fees, refunds and postage insurance are not in it, so the net here will not match a payout statement.

### A half-cent difference rounds the wrong way
**ROUND** rounds halves away from zero. In the test, $7.35 charged against a $7.355 label gave -$0.01 and $5.00 against $4.355 gave $0.65. Type label costs in whole cents to avoid this.

### Two orders tie for the biggest shortfall
**I11** names the first of them in your list. In the test, with a tie, it returned the earlier row.

### The order number cell is filled but nothing else
A row with an order number and no amounts shows Check, because the amounts are missing. Delete the number or add the amounts.

## Template

[Download the .xlsx](templates/tidy-tabs-shipping-charged-vs-label-cost.xlsx)

The file has two tabs. **Orders** has 200 rows from row 5 (yellow inputs in **A:D**, formulas in **E:F**) and a summary in **I5:I14**: orders counted, shipping charged, label cost, net, orders short, total lost, the order with the biggest shortfall, label cost as a percent of shipping charged, rows showing Check and a **CHECK** cell. **How to use** has notes on each formula.

**CHECK** says OK when no row shows Check and the differences add up to the net. Fix every Check before you read the totals.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
