---
title: How to Sort Products Into Four Groups by Sales and Margin in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 226 workbooks and 239,786 cell comparisons, 0 mismatches, plus the edge cases listed in the post; the scatter chart was seen in LibreOffice's PDF export. Not tested in Excel or Google Sheets.
search_description: Split products at the median of units sold and of margin, label four groups with IF, then count them and total profit per group with COUNTIF and SUMIF.
image_prompts: A craft fair booth table seen from the side with neat rows of wrapped soap bars and small tin candles on one end, a nearly bare wooden crate with only two beeswax wraps left on the other end, and a small metal cash box partly open between them, late afternoon outdoor light, realistic photo, no screens, no readable text, no logos
image_alt: Craft fair table with stacked soap bars, tin candles, a sparse crate and cash box
threads: Ranking by profit alone puts a busy, thin-margin product next to a slow, rich one.\nSplit at the median of units sold and of margin, like =MEDIAN($E$6:$E$205), and IF gives each product one of four groups.\nIn the sample, 3 of 10 products bring $15,403.20 of the $27,356.70 profit.
---
Split your products at the median of units sold and the median of margin, then label each one with `=IF(B6>=$B$2,IF(E6>=$B$3,"High units, high margin","High units, low margin"),IF(E6>=$B$3,"Low units, high margin","Low units, low margin"))`, where B2 is the units median (`=MEDIAN($I$6:$I$205)`) and B3 is the margin median (`=MEDIAN($E$6:$E$205)`). In the made-up candle and soap maker's 10 products, 3 sell a lot at a good margin and bring $15,403.20 of the $27,356.70 profit.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (226 workbooks, 239,786 cell comparisons, 0 mismatches), plus the edge cases below. The scatter chart was seen in LibreOffice's PDF export. Excel and Google Sheets were not opened for this post.

![Ten handmade products with units sold, margin, profit and four color-coded sales and margin groups](images/2026-11-29-team-04-product-mix-quadrant-units-sold-vs-profit-margin-median-split-countif-sumif-excel-google-sheets/product-mix-quadrant-template.png) *Products tab of a made-up candle and soap maker's 10 products: the units sold line at the median of 255 and the margin line at the median of 55.00%, then units sold, revenue, cost, margin, profit and the group for each product, colored by group (High units, high margin in green for the cedar candle, lavender soap and holiday cards). PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

Ranking by profit alone puts a busy, thin-margin product next to a slow, rich one. A line for units sold and a line for margin turn your list into a small product profitability matrix. Margin is (revenue - cost) / revenue from the cost you type, with no inventory or tax accounting.

The steps are the same in Excel and Google Sheets.

1. On the **Products** tab, type one product per row from row 6: name in **A**, units sold in **B**, revenue in **C** and your total cost for those units in **D**, all for the same period (for example 01/01/2026 to 12/31/2026). If you only have a sales log, build these columns first with [profit by product from a sales log](https://tidytabs.blogspot.com/2026/11/how-to-calculate-profit-by-product-from.html).
2. Check column **H** (Status). Only OK rows get a margin and a group and count in the medians and totals.
3. Column **E** works out the margin with `=IF(H6<>"OK","",ROUND((C6-D6)/C6,4))` and column **F** the profit with `=IF(H6<>"OK","",ROUND(C6-D6,2))`. Cost above revenue gives a negative margin, and the row still counts.
4. Read the lines in **B2** (units sold) and **B3** (margin). **C2** holds `=IF(COUNT($I$6:$I$205)=0,"",MEDIAN($I$6:$I$205))`, where helper column **I** repeats the units of OK rows only, and **C3** holds `=IF(COUNT($E$6:$E$205)=0,"",MEDIAN($E$6:$E$205))`. See the MEDIAN help from [Microsoft](https://support.microsoft.com/en-us/office/median-function-d0916313-4753-414c-8537-ce85bdd967d2) or [Google](https://support.google.com/docs/answer/3094025).
5. To use your own line instead of a median, type a number in **D2** (units) or **D3** (margin, such as 50%). **B2** is `=IF(D2="",C2,D2)` (B3 likewise); clear the cell to go back to the median.
6. Column **G** labels each product with the group formula above (blank when Status is not OK). Because of `>=`, a product exactly on a line goes in the high group. **L7** and **L8** count the products sitting on each line.
7. On the **Summary** tab, **B5** counts a group with `=COUNTIF(Products!$G$6:$G$205,A5)` and **F5** totals its profit with `=ROUND(SUMIF(Products!$G$6:$G$205,A5,Products!$F$6:$F$205),2)`. Units, revenue and cost use the same SUMIF on columns B, C and D. See the SUMIF help from [Microsoft](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) or [Google](https://support.google.com/docs/answer/3093583).
8. Column **G** shows each group's share of profit, `=IF($F$9<=0,"",F5/$F$9)`, and **H** its own margin, profit over revenue.
9. **CHECK** in **K8** is OK only when the four groups hold every OK product, their profit matches the Products tab, and no row shows Check.

The group names describe sales and margin only. They do not tell you which products to keep, drop or reprice.

Related: [products that make up 80 percent of your sales](https://tidytabs.blogspot.com/2026/10/how-to-find-products-that-make-up-80.html), [median order value](https://tidytabs.blogspot.com/2026/10/how-to-find-median-order-value-in-excel.html) and [comparing options with a weighted score](https://tidytabs.blogspot.com/2026/10/how-to-compare-vendors-with-weighted.html).

## Example

The products are made up: a candle and soap maker's 2026 line of 10 items.

| Product | Units sold | Revenue | Cost | Margin | Profit | Group |
|---|---|---|---|---|---|---|
| Cedar candle 8 oz | 412 | $9,876.00 | $3,460.80 | 64.96% | $6,415.20 | High units, high margin |
| Lavender soap bar | 655 | $5,240.00 | $1,572.00 | 70.00% | $3,668.00 | High units, high margin |
| Holiday card 6 pack | 380 | $6,840.00 | $1,520.00 | 77.78% | $5,320.00 | High units, high margin |
| Wool dryer balls | 95 | $2,185.00 | $617.50 | 71.74% | $1,567.50 | Low units, high margin |
| Beeswax wraps | 140 | $2,520.00 | $1,260.00 | 50.00% | $1,260.00 | Low units, low margin |
| Ceramic mug | 210 | $6,300.00 | $3,570.00 | 43.33% | $2,730.00 | Low units, low margin |
| Linen tea towel | 88 | $1,584.00 | $968.00 | 38.89% | $616.00 | Low units, low margin |
| Lip balm | 820 | $3,280.00 | $1,640.00 | 50.00% | $1,640.00 | High units, low margin |
| Gift box set | 120 | $5,400.00 | $2,160.00 | 60.00% | $3,240.00 | Low units, high margin |
| Wax melts | 300 | $2,400.00 | $1,500.00 | 37.50% | $900.00 | High units, low margin |

With 10 products, MEDIAN averages the two middle values: the units line is 255 (210 and 300) and the margin line is 55.00% (50.00% and 60.00%). No product sits on either line; Lip balm and Beeswax wraps, at exactly 50.00%, fall under the margin line.

| Group | Products | Profit | Share of profit |
|---|---|---|---|
| High units, high margin | 3 | $15,403.20 | 56.31% |
| High units, low margin | 2 | $2,540.00 | 9.28% |
| Low units, high margin | 2 | $4,807.50 | 17.57% |
| Low units, low margin | 3 | $4,606.00 | 16.84% |
| Total | 10 | $27,356.70 | 100.00% |

Lip balm (820 units at 50.00%) and Wool dryer balls (95 units at 71.74%) made close profits but land in different groups.

Typing your own lines changes the picture. With 300 in D2 and 50% in D3, the groups hold 4, 1, 3 and 2 products, and Wax melts at exactly 300 units counts as High units.

## Troubleshooting

### The groups always split about half and half
That is what a median does: about half the products land on each side of each line, whatever the year. Type a line of your own in **D2** or **D3**, such as the units you need to justify a product or the margin you aim for.

### A product right on a line counts as high
The group formula uses `>=`, so a product equal to the line goes in the high group. In LibreOffice, dropping Wax melts from the sample left 9 products and a units line of 210, and the Ceramic mug at 210 counted as High units. **L7** and **L8** show how many products sit on each line.

### A product has no group
Its Status is not OK. A product with $0.00 revenue has no margin and shows "Check: revenue 0 or less (no margin)"; delete its row. Units stored as text, such as "210", show "Check: units blank or text"; retype the number.

### No product has a group and CHECK says Check
**L10** says D2 or D3 holds text. In LibreOffice, "55%" stored as text in D3 blanked every group, while the number 0.5 worked (typing 55% by hand was not tried). Clear the cell and retype the number.

### Products past row 205 are missing
The template reads rows 6 to 205. **L9** counts anything typed below, and CHECK stays red until you move those rows up.

## Template

[Download the .xlsx](templates/tidy-tabs-product-mix-quadrant.xlsx)

**Products** takes 200 products (rows 6 to 205) with the lines in **B2:D3** and its checks in **K5:L11**. **Summary** repeats the lines in **B1:B2**, holds the four-group table in **A4:H9** with CHECK in **J4:K8**, and has a scatter chart of units against margin with both lines dashed (seen in LibreOffice's PDF export; not checked in Excel or Google Sheets). **How to use** explains each formula. Products and Summary print Landscape on US Letter.

Yellow cells are inputs, gray cells are formulas, each group has its own color, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
