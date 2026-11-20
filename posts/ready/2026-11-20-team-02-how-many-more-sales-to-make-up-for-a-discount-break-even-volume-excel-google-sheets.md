---
title: How to Calculate How Many More Sales You Need to Make Up for a Discount in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 38 workbooks and 20,330 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Units to sell on sale for the same gross profit: =units*(price-cost)/(price*(1-discount)-cost). A 60% margin at 25% off needs 71.43% more. Free template.
image_prompts: A wooden shop shelf lined with hand-poured candles in amber glass jars, a few jars marked with plain round orange sale stickers, a half-used sheet of the same blank stickers lying on the shelf edge, and an open cardboard carton of extra candles on the floor below, warm boutique lighting, realistic photo, no screens, no readable text, no logos
image_alt: Amber jar candles on a shop shelf, some marked with blank orange sale stickers
threads: 25% off a $40.00 candle that costs $16.00. How many more do you need to sell?\nNot 25% more. 71.43% more, 172 units instead of 100, because the whole $10.00 comes out of your $24.00 profit.\n=units*(price-cost)/(price*(1-discount)-cost)
---
Find the units you must sell on sale to keep the same gross profit with `=units*(price-cost)/(price*(1-discount)-cost)`, or the percent more with `=margin/(margin-discount)-1`. In the made-up example, a $40.00 candle that costs $16.00 (a 60% margin) and sells 100 units needs 171.43 units at 25% off, so 172 whole units or 71.43% more, not 25% more; at 60% off the profit per unit is $0.00 and the download says "Discount wipes out profit".

This is the break-even volume for a discount: the sales it takes to offset the price reduction. It is math on your numbers, not advice on whether to run a sale or a claim about how many extra orders a discount brings.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (38 workbooks, 20,330 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (IF, ROUND, ROUNDUP, COUNTIF, SUMPRODUCT) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Candle, mug and sticker pack discounts with the extra units needed to break even](images/2026-11-20-team-02-how-many-more-sales-to-make-up-for-a-discount-break-even-volume-excel-google-sheets/discount-break-even-extra-sales-template.png) *Fictional products: a $40.00 candle costing $16.00 that sells 100 units at full price needs 171.43, so 172 units, at 25% off (71.43% more), while 60% and 70% off show Discount wipes out profit (the Profit now, Profit at whole units and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. On the **Products** tab, use one product and discount per row from row 5. Formulas are shown for row 5; the numbers are the sample's 25% candle in row 6.

1. Type the product in **A5**, the regular price in **B5** as a plain number (40, not "$40.00" as text) and your unit cost in **C5**.
2. Type the discount in **D5** as a percent: 25%, not 25.
3. Type the units you sell now at the regular price in **E5**, such as 100.
4. Read the margin in **F5**: `=(B5-C5)/B5`, here 60.0%.
5. Read the profit per unit at the sale price in **G5**: `=B5*(1-D5)-C5`. At 25% off the candle sells for $30.00 and leaves $14.00, not $24.00.
6. Read the units needed in **H5**: `=E5*(B5-C5)/(B5*(1-D5)-C5)`. That is profit now ($2,400.00) divided by profit per unit on sale ($14.00): 171.43.
7. Read **I5** for the whole units you need, rounded up (172), and **J5** for the percent more than now (71.43%).
8. Compare **K5** (Profit now) with **L5** (Profit at whole units).

Why not 25% more? A 25% discount takes $10.00 off the price, but all $10.00 comes out of your $24.00 profit, not out of the cost. That is why the percent more depends on the margin: `=margin/(margin-discount)-1` gives 0.60/(0.60-0.25)-1, or 71.43%. When the discount reaches the margin, the sale price is at or below cost and the sheet shows **Discount wipes out profit** in orange. No number of extra sales makes up for that.

For rounding up, see Microsoft's [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) page or Google's [ROUNDUP](https://support.google.com/docs/answer/3093443) page. To show a discount as 25% in Sheets, see [Format numbers in a spreadsheet](https://support.google.com/docs/answer/56470).

This is gross profit on unit cost only: fees, shipping and fixed costs are not included. For those, see [a break-even calculator for fixed costs](https://tidytabs.blogspot.com/2026/10/how-to-build-break-even-calculator-in.html), [net payout after marketplace fees](https://tidytabs.blogspot.com/2026/10/how-to-calculate-net-payout-after.html) and [working out a sale price and discount](https://tidytabs.blogspot.com/2026/10/how-to-calculate-sale-price-and.html).

## Example

The products, prices and costs are fictional.

| Product | Price | Cost | Discount | Units now | Margin | Profit per unit on sale | Units needed | Whole units | More units |
|---|---|---|---|---|---|---|---|---|---|
| Candle | $40.00 | $16.00 | 20% | 100 | 60.0% | $16.00 | 150.00 | 150 | 50.00% |
| Candle | $40.00 | $16.00 | 25% | 100 | 60.0% | $14.00 | 171.43 | 172 | 71.43% |
| Candle | $40.00 | $16.00 | 10% | 100 | 60.0% | $20.00 | 120.00 | 120 | 20.00% |
| Mug | $25.00 | $15.00 | 10% | 80 | 40.0% | $7.50 | 106.67 | 107 | 33.33% |
| Mug | $25.00 | $15.00 | 30% | 80 | 40.0% | $2.50 | 320.00 | 320 | 300.00% |
| Sticker pack | $12.00 | $3.00 | 15% | 200 | 75.0% | $7.20 | 250.00 | 250 | 25.00% |
| Candle | $40.00 | $16.00 | 60% | 100 | 60.0% | $0.00 | Discount wipes out profit | | |
| Candle | $40.00 | $16.00 | 70% | 100 | 60.0% | -$4.00 | Discount wipes out profit | | |

The 40% mug needs 33.33% more units at 10% off and 300.00% more at 30% off. The summary in **O4:P8** shows Rows listed 8, Rows showing Check 0, Discount wipes out profit 2, Typed below row 54 0, and CHECK OK.

The **Table** tab gives the percent for any margin and discount pair. The 60% margin row runs from 9.09% at 5% off to 500.00% at 50% off. The 40% row reaches 700.00% at 35% off and shows No from 40% off.

## Troubleshooting

### Status shows Check: discount
The discount was typed as 25 instead of 25%, typed as the text "25%", left blank, or is above 100%. Retype it as 25%.

### Units needed shows Discount wipes out profit
The discount is as big as the margin or bigger. On the 60% candle, 60% off gave $0.00 per unit and 70% off gave -$4.00, and both rows showed Discount wipes out profit.

### Status shows Check: price
The price was typed as the text "$40.00" or as 0. Retype it as a plain number, such as 40.

### Profit at whole units is more than Profit now
Whole units rounds up, so the 25% candle shows 172 units and $2,408.00 against $2,400.00 now. 171 would leave you short, so the extra $8.00 is rounding, not an error.

## Template

[Download the .xlsx](templates/tidy-tabs-discount-break-even-sales.xlsx)

The file has three tabs. **Products** has rows 5 to 54 with the five inputs in **A:E** and the results in **F:M**, ending in a Status of OK or a Check (such as Check: cost at or above price). The summary in **O4:P8** counts rows listed, Check rows, wipes-out rows and entries typed below row 54, plus the CHECK.

**Table** lists margins 20% to 80% in column A and discounts 5% to 50% in row 4; type your own in those yellow cells. **How to use** explains each formula. Products and Table print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Discount wipes out profit and the Table's No cells turn orange, and Check rows and a CHECK that is not OK turn red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
