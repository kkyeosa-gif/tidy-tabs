---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-find-which-products-are-bought.html
title: How to Find Which Products Are Bought Together Most Often in Excel and Google Sheets
labels: excel-formulas, small-business, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 18 workbooks and 92,465 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Count orders that contain two products with COUNTIFS and SUMPRODUCT, see what each item sells with most, and avoid counting item lines instead of orders.
image_prompts: An open small shipping box on a packing table holding a lavender jar candle and a metal wick trimmer side by side, with a second open box next to it holding a soap bar and a lip balm tube, and a few folded gift boxes stacked behind, soft natural light in a home candle studio, realistic photo, no screens, no readable text, no logos
image_alt: Two open order boxes, one with a candle and wick trimmer, one with soap
threads: Counting item lines is not counting orders.\nMark each order 1 or 0 per product with COUNTIFS, then SUMPRODUCT counts the orders that hold both.\nIn the sample shop, Lavender candle goes out with Wick trimmer in 3 of its 5 orders.
---
Turn your sold-items lines into one row per order with a 1 or 0 for each product, `=IF(COUNTIFS(order_col,order,product_col,product)>0,1,0)`, then count the orders that contain both of two products with `=SUMPRODUCT((INDEX(grid,0,col_a)=1)*(INDEX(grid,0,col_b)=1))`. In the made-up shop's 10 orders, Lavender candle and Wick trimmer appear together in 3 orders, and the download shows Lavender candle is bought most often with Wick trimmer, in 3 of its 5 orders (60.0%).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (18 workbooks, 92,465 cell comparisons, 0 mismatches), plus the edge cases below. The functions used also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six-product pair grid with order counts, most often with partner and share columns](images/2026-11-22-team-05-products-frequently-bought-together-pair-counts-countifs-sumproduct-index-excel-google-sheets/products-bought-together-pairs-template.png) *Pairs tab with 10 fictional orders and 6 products, where Lavender candle and Wick trimmer share 3 orders and Lavender candle shows Wick trimmer as its most frequent partner at 60.0%. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets.

1. On **Pairs**, type your products in **A5:A16**, up to 12. Row 4 and the **Baskets** headings repeat them by formula.
2. On **Items**, paste your sold-items export with one row per item: order number in **A** and product in **B**, rows 5 to 504. The lines of one order do not have to be next to each other. If your export puts every item of an order in one cell, split them into rows first.
3. Number each order once. **Items C5** is `=IF(A5="","",1)` and **C6** is `=IF(A6="","",IF(COUNTIF($A$5:A6,A6)=1,MAX($C$5:C5)+1,""))`, so each order is numbered once, on its first line.
4. List each order once. **Baskets A5** is `=IFERROR(INDEX(Items!$A$5:$A$504,MATCH(ROWS(A$5:A5),Items!$C$5:$C$504,0)),"")`, which lists order 1, 2, 3 down the column.
5. Mark what each order contains. **Baskets C5** is `=IF(OR($A5="",C$4=""),"",IF(COUNTIFS(Items!$A$5:$A$504,$A5,Items!$B$5:$B$504,C$4)>0,1,0))`. It gives 1 when the order holds that product and 0 when it does not, so a product on two lines of one order still counts once. See [COUNTIFS in Excel](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [in Google Sheets](https://support.google.com/docs/answer/3256550).
6. Count each pair. **Pairs C5**, like every grid cell off the diagonal, is `=IF(OR($A5="",C$4=""),"",SUMPRODUCT((INDEX(Baskets!$C$5:$N$304,0,MATCH($A5,Baskets!$C$4:$N$4,0))=1)*(INDEX(Baskets!$C$5:$N$304,0,MATCH(C$4,Baskets!$C$4:$N$4,0))=1)))`. INDEX with row 0 returns a whole product column, and SUMPRODUCT counts the orders where both columns hold 1. See Microsoft's [SUMPRODUCT help](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e).
7. Count orders per product. **O5** is `=IF($A5="","",SUM(INDEX(Baskets!$C$5:$N$304,0,MATCH($A5,Baskets!$C$4:$N$4,0))))`.
8. Find the top partner. **P5** is `=IF(OR($A5="",MAX(B5:M5)=0),"",INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0)))`, **Q5** is `=IF(P5="","",MAX(B5:M5)/O5)` for the share, and **R5** is `=IF(P5="","",COUNTIF(B5:M5,MAX(B5:M5)))` for how many partners tie.

These are counts of orders that contain both products. They do not say that a bundle will sell more.

## Example

The orders are fictional: 10 orders, 5001 to 5010, with 22 item lines and 6 products. Order 5004 lists Cedar candle twice, and the Items Status shows "Repeat in this order (counted once)" on its second line.

| Product | Lavender | Cedar | Wick | Soap | Lip balm | Gift box | Orders | Most often with | Share |
|---|---|---|---|---|---|---|---|---|---|
| Lavender candle | | 1 | 3 | 0 | 0 | 2 | 5 | Wick trimmer | 60.0% |
| Cedar candle | 1 | | 2 | 0 | 0 | 0 | 2 | Wick trimmer | 100.0% |
| Wick trimmer | 3 | 2 | | 0 | 0 | 1 | 4 | Lavender candle | 75.0% |
| Soap bar | 0 | 0 | 0 | | 3 | 1 | 4 | Lip balm | 75.0% |
| Lip balm | 0 | 0 | 0 | 3 | | 1 | 3 | Soap bar | 100.0% |
| Gift box | 2 | 0 | 1 | 1 | 1 | | 3 | Lavender candle | 66.7% |

Read across a row: Gift box is in 3 orders, and 2 of them also hold a Lavender candle. Shares differ by direction. Every Cedar candle order has a Wick trimmer (100.0%), but only 2 of the 4 Wick trimmer orders have a Cedar candle.

## Troubleshooting

### Counts are too high
A plain COUNTIF on the item lines counts lines, not orders. On the sample, COUNTIF of "Cedar candle" on **Items** gave 3 lines, while **O6** shows 2 orders. Use the 1-or-0 grid on **Baskets** so each product counts once per order.

### Every pair cell shows #VALUE!
The two INDEX columns were multiplied directly, as in `SUMPRODUCT(INDEX(...)*INDEX(...))`. In LibreOffice that gave #VALUE! because the empty Baskets rows hold "". Keep the `=1` test on each column.

### A product is not counted
The name on **Items** has a trailing space ("Soap bar ") or is missing from the Pairs list. **Items** Status shows Check: extra spaces or Check: product not on Pairs list. Upper and lower case did not matter in LibreOffice ("soap bar" counted as Soap bar); Excel and Google Sheets were not checked.

### Most often with picks one of several equal partners
On a tie, P shows the first product in the list. In a test where Lavender candle shared 2 orders each with Soap bar, Lip balm and Gift box, P showed Soap bar and **R** showed 3.

### CHECK in F24 shows Check
The orders on **Items** (**F19**) and on **Baskets** (**F20**) differ, a row shows Check, or rows sit below row 504. An order number stored as text ("5001" next to the number 5001) also left that line out of the counts in LibreOffice; Status flags it as Check: order # is text.

## Template

[Download the .xlsx](templates/tidy-tabs-bought-together-pairs.xlsx)

The file has four tabs. **Items** holds up to 500 item lines in rows 5 to 504 (inputs in **A:B**), Order no. and Status in **C:D**, and a summary in **F4:G8**. **Baskets** lists up to 300 orders in rows 5 to 304 with a 1 or 0 for each product in **C:N**; the 1 cells are light green.

**Pairs** holds up to 12 products in **A5:A16**, the pair grid in **B5:M16** with an empty diagonal, the partner columns in **O:S**, and checks in **A18:F24** ending with CHECK in **F24** (OK on the sample). **How to use** explains each formula. The data tabs print Landscape on US Letter.

To go past 12 products or 300 orders you would add rows and columns and widen the ranges; that larger layout was not tested.

Yellow cells are inputs and gray cells are formulas. Every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
