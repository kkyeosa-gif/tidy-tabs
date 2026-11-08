---
threads_url: https://www.threads.com/@tin_ylab/post/DffJCg_oHkC
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-profit-by-product-from.html
title: How to Calculate Profit by Product From a Sales Log in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 16 workbooks and 432 comparisons with 0 mismatches. Checked: the sample, no sales, a discount larger than the line, a misspelled and a double space product name, a quantity typed as text, a blank row between sales, a lowercase product name, 60 sales rows, a text price, a duplicate product row, half cent prices, and six random sets. Excel and Google Sheets were not used. Not run: wildcard characters inside product names, more than 60 sales rows or 20 products.
search_description: Total units, revenue after discounts, cost and profit for each product from a sales log with SUMIFS in Excel and Google Sheets. Free template.
image_prompts: A stack of small kraft price tags, a pencil and a canvas tote bag on a wooden market table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Kraft price tags and a pencil beside a canvas tote on a wooden table
threads: Which product actually makes you money?\nUnits: =SUMIFS(Sales!$D$2:$D$61,Sales!$C$2:$C$61,$A2). Profit: =E2-F2.\nIn the sample, Soy Candle 8 oz sells 6 units, costs $38.40 and earns $68.10 (made-up products).
---
Total the units with `=SUMIFS(Sales!$D$2:$D$61,Sales!$C$2:$C$61,$A2)`, take revenue after discounts with `=ROUND(D2*B2-SUMIFS(Sales!$E$2:$E$61,Sales!$C$2:$C$61,$A2),2)`, and profit with `=E2-F2`, where F2 is the cost `=ROUND(D2*C2,2)`. In the sample, Soy Candle 8 oz sells 6 units for $106.50 after a $1.50 discount, costs $38.40 and earns $68.10, a 63.9% margin (made-up products).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (16 workbooks, 432 comparisons). The formulas use only IF, AND, ISNUMBER, SUMIFS, ROUND and SUM, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Product table showing units, revenue, cost, profit and margin for six sample products](images/2026-11-08-team-01-profit-by-product-from-sales-log-sumifs-excel-google-sheets/profit-by-product-from-sales-log-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional.*

## Steps

The steps are the same in Excel and Google Sheets. The template has two data tabs. **Sales** is the log: date in **A**, order in **B**, product in **C**, quantity in **D**, discount in dollars in **E**, from row 2. **Products** is your price list: product in **A**, price in **B**, unit cost in **C**, from row 2. The summary columns **D:H** sit next to the list.

1. On **Products**, type each product name once in **A**, with its price in **B** and what one unit costs you in **C**.
2. On **Sales**, log each line: the product name exactly as it appears on the price list, the quantity, and any dollar discount for that line. Type 0 or leave nothing if there was none.
3. In **D2** of **Products**, type `=SUMIFS(Sales!$D$2:$D$61,Sales!$C$2:$C$61,$A2)`. This adds the quantities on every Sales row whose product matches **A2**.
4. In **E2**, type `=ROUND(D2*B2-SUMIFS(Sales!$E$2:$E$61,Sales!$C$2:$C$61,$A2),2)`. This is units times price, minus the discounts for that product.
5. In **F2**, type `=ROUND(D2*C2,2)` for the cost of goods.
6. In **G2**, type `=E2-F2` for profit.
7. In **H2**, type `=IF(E2=0,"",G2/E2)` and format the column as a percentage. A product with no revenue stays blank instead of showing #DIV/0!.
8. Fill **D2:H2** down. In the total row, add **D**, **E**, **F** and **G** with `=SUM(D2:D21)` and so on, and type `=IF(E22=0,"",G22/E22)` for the overall margin.

The template wraps these in checks so a text price or an unknown product shows Check instead of a wrong number. The version above is the core.

Rows are locked to a fixed range ($2:$61) so the formulas stay short. If your log grows past row 61, extend every range to the same last row.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), and Google's [SUMIFS](https://support.google.com/docs/answer/3238496). They describe the functions, not our test.

## Example

Sample data is fictional: 6 products and 12 sales lines from 10/01/2026 to 10/28/2026.

| Product | Units | Revenue | Cost | Profit | Margin |
|---|---|---|---|---|---|
| Soy Candle 8 oz | 6 | $106.50 | $38.40 | $68.10 | 63.9% |
| Lip Balm | 7 | $24.50 | $6.65 | $17.85 | 72.9% |
| Tote Bag | 4 | $94.00 | $36.40 | $57.60 | 61.3% |
| Enamel Pin | 7 | $51.00 | $16.10 | $34.90 | 68.4% |
| Sticker Pack | 10 | $42.00 | $7.00 | $35.00 | 83.3% |
| Gift Box | 0 | $0.00 | $0.00 | $0.00 | blank |
| Total | 34 | $318.00 | $104.55 | $213.45 | 67.1% |

Soy Candle 8 oz sells 3 units on 10/01/2026, 2 on 10/09/2026 with a $1.50 discount, and 1 on 10/25/2026. That is 6 units at $18.00, minus $1.50, which is $106.50. Cost is 6 times $6.40, which is $38.40.

Gift Box has no sales, so its margin is blank. The 67.1% overall margin is total profit divided by total revenue, not the average of the product margins, which would be about 70.0%. Whether a margin is good depends on your business, and the template does not judge it.

## Troubleshooting

### A product shows 0 units but you sold it
The name on the Sales row does not match the price list. A trailing space or a double space inside the name ("Tote  Bag") counts as a different name. In the test, 3 units under that spelling left the unmatched units cell at 3 and the Sales row said Check product. Retype the name or pick it from the list.

### The units total on Products is lower than on Sales
Compare the two totals. The template shows the gap as unmatched units, and it should be 0. A positive gap means some Sales names are not on the price list. A negative gap means a product is listed twice on the price list, so its units were counted twice.

### A quantity typed as text is ignored
`SUMIFS` skips a quantity stored as text, such as "3" with a leading apostrophe, so the unmatched units check does not catch it. The Sales row check shows Check qty. Retype the number.

### Revenue for a product is negative
A discount bigger than price times quantity gives negative revenue. The Sales row shows Check discount. Fix the discount or the quantity.

### A lowercase or wildcard name still matches
`SUMIFS` ignores case, so "tote bag" matches "Tote Bag". It also treats `*` and `?` as wildcards. We did not test wildcard characters inside product names, so keep them out of your list.

## Template

[Download the .xlsx](templates/tidy-tabs-profit-by-product.xlsx)

The file has three tabs. **Sales** has 60 log rows from row 2 with 12 fictional sample lines and a **Row check** column. **Products** has slots for 20 products, the summary columns, a total row, and a side block that counts products listed, products with no sales, units, unmatched units, rows showing Check, and a **CHECK** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. Revenue and cost are rounded to cents per product.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
