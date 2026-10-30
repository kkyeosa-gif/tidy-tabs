---
title: How to Track Deposits and Balance Due on Custom Orders in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model of paid, balance, status and the summary, in 27 workbooks (6,981 cell comparisons, 0 mismatches). Checked: three part payments, no payments, a payment equal to the total, a payment $0.01 over, a due date equal to the as-of date and one day earlier, a moved as-of date, capital letters, a trailing space, a typo order number, a text amount, blank rows, a $0.00 total, a text total, a duplicate order number and 60 full rows. TODAY() in the as-of cell was not run. Excel and Google Sheets were not used.
search_description: Track deposits and balance due on custom orders with SUMIF in Excel and Google Sheets. Free template with part payments, past due labels and a check cell.
image_prompts: A wooden tray holding a few handwritten order tags and a small stack of folded cash and receipts next to a half-finished embroidered tote bag on a craft table, soft daylight, no screens, no readable text, no logos
image_alt: Craft table with order tags, folded cash and a half-finished embroidered tote bag
threads: Customer paid a $60.00 deposit, then $140.00 later. Which order is that again?\nKeep payments on their own sheet and let SUMIF add them up per order number.\nThe balance, a Past due label and an Overpaid flag all come from that one number.
---
Add up what each customer has paid with `=SUMIF(Payments!$A$5:$A$104,A5,Payments!$C$5:$C$104)`, then subtract it from the order total with `=ROUND(D5-F5,2)` to get the balance due. Keep payments on their own sheet so a $50.00 deposit and a later $120.00 final payment both count toward the same order number.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. SUMIF, ROUND and COUNTIF exist in Excel and Google Sheets, but neither app was opened for this post.

![Twelve custom orders with order total, due date, paid so far, balance due and status, plus a summary](images/2026-10-31-team-01-track-custom-order-deposits-balance-due-sumif-excel-google-sheets/custom-order-deposits-balance-due-template.png) *caption: the Orders sheet of the template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

The layout matches the template. **Orders** has one row per order starting in row 5: order number in column A, customer in B, item in C, order total in D and due date in E. **Payments** has one row per payment starting in row 5: order number in A, date in B, amount in C and method in D. **Settings B1** holds the as-of date.

Type the as-of date instead of using `TODAY()`, so the labels stay fixed when you open the file later. Swapping in `=TODAY()` makes Past due move every day; that swap was not run in the tests.

The steps are the same in Excel and Google Sheets.

1. On **Settings**, type 10/31/2026 in **B1**.
2. On **Payments**, enter each payment in columns A to D. Type the order number exactly as it appears on **Orders**, and enter the amount as a number.
3. On **Orders**, list each order once, with its total and due date.
4. In **F5**, type the paid-so-far formula below.
5. In **G5**, type `=IF(A5="","",ROUND(D5-F5,2))`. ROUND keeps cent fractions from leaving a balance like $0.00000001.
6. In **H5**, type the status formula below.
7. Select **F5:H5** and drag the fill handle down to your last order.
8. Format columns D, F and G as currency.

The paid-so-far formula in **F5**:

`=IF(A5="","",SUMIF(Payments!$A$5:$A$104,A5,Payments!$C$5:$C$104))`

SUMIF adds every amount on **Payments** whose order number matches A5. Three part payments on one order are three rows, and they all land in the same total. 
The status formula in **H5**:

`=IF(A5="","",IF(G5<0,"Overpaid",IF(G5=0,"Paid in full",IF(F5=0,IF(E5<Settings!$B$1,"Past due","Awaiting deposit"),IF(E5<Settings!$B$1,"Past due","Deposit received")))))`

An order is Past due only when money is still owed and the due date is before the as-of date. A due date equal to the as-of date is not Past due. The template adds guards for blank or text cells, so its formula is longer.

For a deposit percent, use `=ROUND(F5/D5,4)` formatted as a percent.

Help pages: Microsoft's [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), and Google's [SUMIF page](https://support.google.com/docs/answer/3093583). Those pages describe the functions, not our test.

## Example

Sample data is fictional: 12 made up customers and 20 payments, as of 10/31/2026.

| Order | Total | Due date | Paid so far | Balance due | Status |
|---|---|---|---|---|---|
| CO-1001 | $480.00 | 11/20/2026 | $250.00 | $230.00 | Deposit received |
| CO-1002 | $85.00 | 10/24/2026 | $40.00 | $45.00 | Past due |
| CO-1003 | $200.00 | 10/29/2026 | $200.00 | $0.00 | Paid in full |
| CO-1004 | $120.00 | 11/08/2026 | $0.00 | $120.00 | Awaiting deposit |
| CO-1005 | $95.00 | 10/31/2026 | $0.00 | $95.00 | Awaiting deposit |
| CO-1009 | $130.00 | 11/03/2026 | $135.00 | -$5.00 | Overpaid |

CO-1003 is the two-payment case: $60.00 on 10/05/2026 and $140.00 on 10/27/2026. CO-1005 is due on the as-of date, so it is not Past due yet. CO-1009 was overpaid by $5.00.

The summary beside the list shows 12 orders, $1,880.00 ordered, $970.00 received and $910.00 still owed. Three orders are Past due, with $240.00 between them. The check cell says OK.

Move the as-of date to 11/25/2026 and the Past due count goes from 3 to 8. This arithmetic does not say whether a deposit is refundable or what a contract requires.

## Troubleshooting

### A payment is missing from Paid so far
The order number does not match exactly. In the tests, a trailing space ("CO-1003 ") matched nothing, while capital letters did not matter ("co-1003" matched). The Check column on **Payments** says No order for that row, the summary counts it, and the check cell shows CHECK. Retype the number or remove the space.

### A payment amount is skipped
The amount is stored as text, such as "60" with a leading apostrophe. SUMIF skips text. The Check column says Check amount and the check cell shows CHECK. Retype the amount as a number.

### The balance is negative
The customer paid more than the order total, so the status says Overpaid. Check the payment amounts and the total.

### Every status says Check settings
The as-of date is text. **Settings B2** shows the reason. Retype the date as a real date such as 10/31/2026.

### The check cell says CHECK
It shows CHECK for an unmatched payment or text amount, an order number listed twice, a missing or text total, or bad settings. The summary cells beside it point to which one.

## Template

[Download the .xlsx](templates/tidy-tabs-custom-order-deposits.xlsx)

The file has four tabs. **Orders** has 12 fictional orders in **A5:E64** (60 slots), the paid, balance, status and deposit percent columns, and a summary in **K** and **L**. **Payments** holds 20 fictional payments in **A5:D104** (100 slots) plus a Check column. **Settings** has the yellow as-of date and a settings check. **How to use** has short notes.

Limits: more than 60 orders or 100 payments, and order numbers with wildcard characters, were not tested. The file was recalculated only in LibreOffice Calc 24.2.7.2. Orders prints landscape and Payments prints portrait, both on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
