---
title: How to Build a Break-Even Calculator in Excel and Google Sheets
labels: formulas, pricing, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps follow Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Build a break-even calculator with =ROUNDUP(D2/E2,0) for units and units x price for revenue. Free .xlsx template with booth and Etsy sample products.
image_prompts: A craft fair table with a small cash box and a stack of blank paper booth-fee receipts next to three soy candles in glass jars, a pencil and a pocket calculator, tent shade in the background, all paper text blurred and unreadable, no logos
image_alt: Cash box and booth receipts beside candles and a pocket calculator
threads: 11 candles pays off a $120.00 booth fee, not 10.\nFixed costs divided by profit per unit, rounded up: =ROUNDUP(D2/E2,0).\nAt $11.50 profit it is 10.43, and ten candles leave you $5.00 short.
---
Divide your fixed costs by the profit per unit and round up: `=ROUNDUP(D2/E2,0)`, where E2 is price minus variable cost. A $120.00 booth fee and an $11.50 profit per candle means you need to sell 11 candles to break even.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every formula and result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Break-even sheet listing four handmade products with price, fixed costs, profit per unit, and break-even units](images/2026-10-02-team-01-break-even-calculator-excel-google-sheets/break-even-calculator-template.png) *LibreOffice Calc 24.2 PDF export of the Break-even sheet with sample data (fake).*

## Steps

Use one row per product. Fixed costs are what you pay once no matter how many you sell, such as a booth fee or a table rental. Variable cost is what one sale costs you, such as materials, packaging, and payment fees.

1. In row 1, type these headers in **A1:G1**: Product, Price, Variable cost per unit, Fixed costs, Profit per unit, Break-even units, Break-even revenue.
2. Type the product name in **A2**, the selling price in **B2**, the variable cost for one unit in **C2**, and the fixed costs for that product in **D2**.
3. In **E2** (Profit per unit), type `=B2-C2`.
4. In **F2** (Break-even units), type `=IF(E2<=0,"No break-even",ROUNDUP(D2/E2,0))`. ROUNDUP, not ROUND, because 10.43 candles means you still need an 11th sale to cover the fee.
5. In **G2** (Break-even revenue), type `=IF(ISNUMBER(F2),F2*B2,"")`. This is units times price.
6. Select **E2:G2** and drag the fill handle down to cover every product row.
7. Format B, C, D, E, and G as currency.

The IF in step 4 handles a price that is not higher than the variable cost. Without it, the formula returns a negative number or a divide-by-zero error. With it, you see "No break-even" and the revenue cell stays blank.

### In Excel

The formulas work as typed. Select the cells, then choose **Home > Number Format > Currency** for the dollar columns. Microsoft documents the arguments on its [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) page.

### In Google Sheets

The same formulas work as typed. Select the dollar columns, then choose **Format > Number > Currency**. Google documents the arguments on its [ROUNDUP function](https://support.google.com/docs/answer/3093444) page.

## Example

Sample data below is fictional. Fixed costs are per product: for example, the share of a booth fee you assign to candles.

| Product | Price | Variable cost | Fixed costs | Profit per unit | Break-even units | Break-even revenue |
|---|---|---|---|---|---|---|
| Lavender soy candle, 8 oz | $18.00 | $6.50 | $120.00 | $11.50 | 11 | $198.00 |
| Oatmeal goat milk soap bar | $8.00 | $3.25 | $75.00 | $4.75 | 16 | $128.00 |
| Brass hoop earrings | $24.00 | $9.60 | $150.00 | $14.40 | 11 | $264.00 |
| Letterpress birthday card | $6.00 | $1.50 | $60.00 | $4.50 | 14 | $84.00 |

The raw division gives 10.43 candles, 15.79 soap bars, 10.42 earrings, and 13.33 cards. ROUNDUP turns each into a whole unit, which is why the revenue is units times price, not the fixed cost itself.

Check the candle row by hand. Ten candles earn 10 x $11.50 = $115.00, which is $5.00 short of the $120.00 fee. Eleven earn $126.50, so the fee is covered.

## Troubleshooting

### Break-even units shows "No break-even"
Price is less than or equal to the variable cost, so each sale loses money or earns nothing. Raise the price or lower the cost. A $6.00 candle with a $6.50 cost is an example.

### The result has decimals, like 10.43
The formula uses ROUND or no rounding. Use `ROUNDUP(D2/E2,0)` so the answer is a whole unit. ROUND would show 10 candles, which leaves you $5.00 short.

### Profit per unit looks wrong
Variable cost left out something that comes with every sale, such as packaging or a card processing fee. Add those into column C. Do not put one-time costs there.

### Revenue shows $0.00 or blank
A blank is expected when units show "No break-even". If units are a number and revenue is $0.00, check that the price in column B is a number and not text. A left-aligned price is usually text.

### Break-even does not include my own pay
The template covers costs only. If you want to earn a target amount, add it to the fixed costs in column D before the formula runs.

## Template

[Download the .xlsx](templates/tidy-tabs-break-even-calculator.xlsx)

The file has two tabs. **Break-even** holds the four sample products with the formulas from the steps above, and shades the formula columns (Profit per unit, Break-even units, Break-even revenue). **How to use** has short fill-in notes.

The products, prices, and costs are fictional. The file is arithmetic only and does not include taxes or your own pay. It was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets), and the "No break-even" guard was checked there by lowering a price below its cost. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
