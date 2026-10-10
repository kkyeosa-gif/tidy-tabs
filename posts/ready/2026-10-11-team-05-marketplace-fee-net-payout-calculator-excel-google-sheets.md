---
title: How to Calculate Net Payout After Marketplace Fees in Excel and Google Sheets
labels: excel-formulas, pricing, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was compared with hand or Python math, plus two edge cases run on edited copies. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Net payout after marketplace fees with =B2-ROUND(B2*($H$2+$H$3)+$H$4,2), plus a reverse formula for the price that nets a target. Free .xlsx template with fictional rates.
image_prompts: Craft studio packing table with two wrapped candles, a small kraft shipping box and a roll of twine, soft window light, no screens, no readable text, no logos
image_alt: Two wrapped candles beside a small kraft shipping box and a roll of twine
threads: Sold a $1.00 gift tag set and kept $0.65?\nNet payout is one formula: =B2-ROUND(B2*($H$2+$H$3)+$H$4,2). Two percentage fees plus a fixed fee per order.\nThe fixed fee eats small orders. A $0.25 sticker nets -$0.02 in the sample.
---
Subtract the fees from the price with `=B2-ROUND(B2*($H$2+$H$3)+$H$4,2)`, where H2 and H3 are the percentage fees and H4 is the fixed fee per order. The sample uses made up rates; type the rates from your own seller account.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against hand or Python math. The Excel and Google Sheets steps use the same functions and link to their help pages. Neither Excel nor Google Sheets was opened for this post.

![Six sample orders with marketplace fees, net payout and price for a target](images/2026-10-11-team-05-marketplace-fee-net-payout-calculator-excel-google-sheets/marketplace-net-payout-calculator-template.png) *Render of the Net payout tab of the template, a PDF export from LibreOffice Calc 24.2 using fictional sample data and made up fee rates*

## Steps

Most marketplaces take a percentage of the price, a payment processing percentage, and a small fixed fee on every order. This sheet adds those up per order and shows what you keep.

1. In row 1, type these headers in A to E: Order, Sale price, Fees, Net payout, Net % of price.
2. Type each item name in **A2:A7** and its sale price in **B2:B7** as plain numbers.
3. In **G2:G4**, type labels for the three fees: marketplace fee, payment processing, fixed fee per order.
4. In **H2**, type the marketplace fee as a percentage, such as 6.5%. In **H3**, type the payment processing percentage. In **H4**, type the fixed fee in dollars.
5. In **C2**, type `=ROUND(B2*($H$2+$H$3)+$H$4,2)`. This is the fee on that order, rounded to cents.
6. In **D2**, type `=B2-C2`. This is your net payout.
7. In **E2**, type `=D2/B2` and format it as a percentage. This is the share of the price you keep.
8. Select **C2:E2** and drag the fill handle down to row 7.
9. In row 8, add totals with `=SUM(B2:B7)` and the same formula for columns C and D. In **E8**, type `=D8/B8`.

The `$` signs lock H2, H3 and H4, so every row reads the same three fee cells. Change a rate once and every order updates.

### Work backward to a price

To find the price that nets a target, type the target in **H7**, for example $15.00. Then, in **H8**, type `=ROUNDUP((H7+H4)/(1-H2-H3),2)`. In **H9**, type `=H8-ROUND(H8*(H2+H3)+H4,2)` to check it.

H8 adds the fixed fee to the target, then divides by the share of the price you keep after both percentages. `ROUNDUP` rounds up to the next cent so the price never falls short.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) page and its [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) page.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [ROUND](https://support.google.com/docs/answer/3093440) help page. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional, and so are the rates: 6.5% marketplace fee, 3.0% payment processing and $0.25 per order. They are not any marketplace's current fees.

| Order | Sale price | Fees | Net payout | Net % of price |
|---|---|---|---|---|
| Soy candle, 8 oz | $18.00 | $1.96 | $16.04 | 89.1% |
| Cedar soap bar | $8.50 | $1.06 | $7.44 | 87.5% |
| Brass hoop earrings | $24.00 | $2.53 | $21.47 | 89.5% |
| Gift tag set | $1.00 | $0.35 | $0.65 | 65.0% |
| Candle gift box | $60.00 | $5.95 | $54.05 | 90.1% |
| Sticker sample | $0.25 | $0.27 | -$0.02 | -8.0% |
| Total | $111.75 | $12.12 | $99.63 | 89.2% |

Check the candle by hand: 6.5% plus 3.0% is 9.5%, and 9.5% of $18.00 is $1.71. Add the $0.25 fixed fee for $1.96, and $18.00 minus $1.96 is $16.04.

The cedar soap bar shows how rounding works. $8.50 times 9.5% is $0.8075, plus $0.25 is $1.0575, which rounds to $1.06.

The fixed fee hurts small orders most. The $1.00 gift tag set keeps only 65.0% of its price. The $0.25 sticker sample ends at a net payout of -$0.02, which is a loss.

For the reverse formula, a $15.00 target gives a price of $16.86. A $20.00 target gives $22.38, and the check shows exactly $20.00.

One cent caveat: the $15.00 target checks out at $15.01, not $15.00. Fees are rounded to cents, so $16.85 would net exactly $15.00 by hand, while the formula rounds up to $16.86. If you want the lowest price, try the price one cent lower in column B and read the net payout.

## Troubleshooting

### The net payout is negative on a cheap item
The fixed fee is larger than what the percentages leave behind. In the sample, a $0.25 sticker owes $0.27 in fees. Raise the price or bundle the item with another.

### The reverse price check is one cent over the target
This is expected. Fees round to cents, so the check in H9 can land a cent above the target. It is never below, because H8 rounds up.

### H8 and H9 show #DIV/0!
The two percentage fees add up to 100%, so nothing is left to divide by. Check H2 and H3. A fee typed as 6.5 instead of 6.5% also breaks the math, so make sure the cells hold percentages.

### The fees are off by a few cents from the marketplace statement
This sheet rounds each order's fee to cents. Your marketplace may round differently, or may charge on shipping or tax too. Compare one order against your statement and adjust the settings in H2:H4.

### A zero fixed fee changes the small orders
With H4 at $0.00, the sticker sample fee drops to $0.02 and its net becomes $0.23. If your marketplace has no per-order fee, set H4 to 0.

## Template

[Download the .xlsx](templates/tidy-tabs-marketplace-net-payout.xlsx)

The file has two tabs. **Net payout** holds six sample orders in columns A and B, the fees, net payout and net percentage in C to E, a totals row, yellow fee settings in H2:H4 and the target price calculator in H7:H9. **How to use** has short fill-in notes. It prints landscape on one US Letter page.

The sheet covers only the fees you type in. It does not include shipping, sales tax collected, advertising, refunds or income tax. Marketplaces change their fees, so use the current rates from your own seller account.

The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
