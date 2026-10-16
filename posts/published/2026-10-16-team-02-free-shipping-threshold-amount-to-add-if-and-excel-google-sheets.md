---
threads_url: https://www.threads.com/@tin_ylab/post/DekWNdxD9R9
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-check-free-shipping-threshold.html
title: How to Check a Free Shipping Threshold and Show the Amount to Add in Excel and Google Sheets
labels: excel-formulas, small-business, ecommerce
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against a Python Decimal comparison, including a subtotal exactly at the threshold, a different threshold, a subtotal typed as text, a blank subtotal, a blank delivery cell and a pickup order above the threshold. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Free shipping check: =IF(AND(C2="Ship",B2>=$I$2),"Free","Pay"), plus the amount to add with MAX. Fictional sample data and a free .xlsx template.
image_prompts: Small brown shipping box on a packing scale beside a roll of kraft tape and a folded padded mailer on a wooden packing table, soft window light, no screens, no readable text, no logos
image_alt: Shipping box on a small scale beside kraft tape and a padded mailer
threads: Cart at $49.99 and free shipping starts at $50.00?\nThe sample sheet shows $0.01 to add, and counts the 3 orders within $10.00 of the threshold.\n=IF(AND(C2="Ship",B2>=$I$2),"Free","Pay") sets the label, and pickup orders stay out of it.
---
Label an order with `=IF(AND(C2="Ship",B2>=$I$2),"Free","Pay")`, then show how much more the customer needs to add with `=IF(AND(C2="Ship",B2<$I$2),MAX(0,$I$2-B2),0)`. Type your own threshold in I2; pickup orders never get free shipping.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Order list marking each order Free or Pay with the dollar amount still needed](images/2026-10-16-team-02-free-shipping-threshold-amount-to-add-if-and-excel-google-sheets/free-shipping-threshold-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The label needs two tests at once: the order ships, and the subtotal reaches the threshold. AND combines them and IF turns the answer into a word.

1. On a tab named **Orders**, type these headers in row 1: Order, Subtotal, Delivery, Status, Amount to add.
2. Type the order number in column **A**, the subtotal as a plain number in column **B**, and the word `Ship` or `Pickup` in column **C**.
3. In **I2**, type your threshold, such as `50`. Every formula points at this one cell.
4. In **D2**, type `=IF(OR(NOT(ISNUMBER(B2)),C2=""),"",IF(C2="Pickup","Pickup",IF(AND(C2="Ship",B2>=$I$2),"Free","Pay")))`. This shows Free, Pay or Pickup, and stays blank when the subtotal or delivery cell is empty or the subtotal is text.
5. In **E2**, type `=IF(OR(D2="Pay",D2="Free"),MAX(0,ROUND($I$2-B2,2)),"")`. This is the amount the customer still needs to add. MAX keeps it from going below $0.00, and ROUND keeps cents exact.
6. Select **D2:E2** and drag the fill handle down to the last order row.
7. Put a "close" gap in **I3**, such as `10`. In **I6**, type `=COUNTIFS(D2:D31,"Pay",E2:E31,"<="&I3)` to count unpaid-shipping orders within $10.00 of the threshold.

The `$I$2` with dollar signs stays fixed as you fill down. The shorter formulas in the first paragraph work when every row is Ship or Pickup, but a pickup order would read "Pay". The longer D2 avoids that.

### In Excel

The formulas work as typed. Microsoft describes them on its [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) and [AND function](https://support.microsoft.com/en-us/office/and-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9) pages. To restrict column C to two choices, use **Data > Data Validation**, choose List, and type `Ship,Pickup`.

### In Google Sheets

The same formulas should work as typed. For the dropdown, use **Data > Data validation**, add a rule, choose Dropdown, and enter `Ship` and `Pickup`. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. The threshold in I2 is $50.00 and the close gap in I3 is $10.00.

| Order | Subtotal | Delivery | Status | Amount to add |
|---|---|---|---|---|
| ORD-5001 | $32.50 | Ship | Pay | $17.50 |
| ORD-5002 | $50.00 | Ship | Free | $0.00 |
| ORD-5003 | $49.99 | Ship | Pay | $0.01 |
| ORD-5004 | $78.25 | Ship | Free | $0.00 |
| ORD-5005 | $24.00 | Pickup | Pickup | |
| ORD-5006 | $45.00 | Ship | Pay | $5.00 |
| ORD-5007 | $112.40 | Ship | Free | $0.00 |
| ORD-5008 | $18.75 | Ship | Pay | $31.25 |
| ORD-5009 | $61.00 | Pickup | Pickup | |
| ORD-5010 | $41.50 | Ship | Pay | $8.50 |

| Result | Cell | Value |
|---|---|---|
| Free shipping orders | I4 | 3 |
| Pay orders | I5 | 5 |
| Pay orders within $10.00 | I6 | 3 |
| Pickup orders | I7 | 2 |
| Every order counted once | I8 | OK |

ORD-5002 sits exactly on the threshold and gets Free, because the test is `>=`. ORD-5003 is one cent short and shows $0.01. The three close orders (5003, 5006 and 5010) are the ones worth a "add $8.50 for free shipping" nudge.

## Troubleshooting

### A pickup order says Pay
You used the short formula from the first paragraph. It has no Pickup branch, so anything that is not Ship and above the threshold falls to Pay. Use the longer D2 from the steps, which returns Pickup and leaves the amount blank even for ORD-5009 at $61.00.

### An order has no status and I8 shows CHECK
The row has a subtotal typed as text (such as `'55`), a blank subtotal, or a blank delivery cell. In the recalculated copy, an 11th order with any of those got no status and no amount, and I8 changed to CHECK. Fill the cell, or retype the subtotal as a plain number.

### Everything says Pay even though the subtotal is over the threshold
The delivery word must be exactly `Ship` or `Pickup`. A typo such as "Shipping" matches neither test, so the order falls to Pay or stays blank. Use the dropdown in column C to keep the words consistent.

### A $49.99 order shows $0.01, not $0.00
That is correct. The ROUND in E2 rounds the gap to cents, so $50.00 minus $49.99 shows exactly $0.01 in the recalculated copy. Only a subtotal at or above the threshold shows $0.00.

### I changed the threshold and the counts did not match my guess
Every formula reads I2, so the counts move together. In the recalculated copy, a $45.00 threshold with a $5.00 gap gave Free 5, Pay 3, close 1 and Pickup 2. Recheck I2 and I3 before you question the formulas.

## Template

[Download the .xlsx](templates/tidy-tabs-free-shipping-threshold.xlsx)

The file has two tabs. **Orders** holds the 10 sample orders in yellow input cells, room down to row 31, the threshold in **I2**, the close gap in **I3**, and the five results in **I4** to **I8**. Column C has a Ship or Pickup dropdown. **How to use** has short fill-in notes. Both tabs print portrait on US Letter.

The file uses your own threshold only. It does not know carrier rates or marketplace shipping rules. The orders are fictional, and it was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
