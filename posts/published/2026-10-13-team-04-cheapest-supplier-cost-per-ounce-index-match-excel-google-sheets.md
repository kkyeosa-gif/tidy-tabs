---
threads_url: https://www.threads.com/@tin_ylab/post/DeeWf9PIChK
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-cheapest-supplier-by-cost.html
title: How to Find the Cheapest Supplier by Cost Per Ounce in Excel and Google Sheets
labels: excel-formulas, handmade-business, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math, including a zero package size, a blank price, a text price, a tie for lowest and an all-blank case. Excel and Google Sheets were not opened.
search_description: Cost per ounce is =(B2+D2)/C2. =INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)) names the cheapest supplier. Sample: $0.2667/oz. Free .xlsx.
image_prompts: Three bags of soy wax flakes in different sizes on a small shipping scale beside a stack of paper packing slips, soft window light, no screens, no readable text, no logos
image_alt: Three bags of wax flakes of different sizes on a small shipping scale
threads: Three wax suppliers, three package sizes, and shipping on top. Which one is actually cheapest?\nDivide (price + shipping) by ounces, then =INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)) names the winner.\nIn the sample it is $0.2667 per oz.
---
Divide price plus shipping by package size with `=(B2+D2)/C2`, then return the cheapest name with `=INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0))`. Of five fictional wax quotes, Harbor Craft Wholesale wins at $0.2667 per ounce.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Five wax suppliers compared by package price, size and shipping with cost per ounce and a cheapest supplier result](images/2026-10-13-team-04-cheapest-supplier-cost-per-ounce-index-match-excel-google-sheets/cheapest-supplier-cost-per-ounce-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: a bigger bag is not always cheaper, and shipping changes the answer. Put every quote on the same footing, one cost per ounce, then let MIN and MATCH find the lowest.

1. In row 1, type these headers in A to G: Supplier, Package price, Package size (oz), Shipping, Cost per oz, Cost for amount needed, Rank.
2. In **A2:A6**, type one supplier per row.
3. In **B2:B6**, type the price of one package as a plain number, such as 42.50.
4. In **C2:C6**, type the ounces in that package, such as 160.
5. In **D2:D6**, type the shipping charged for that package. Type 0 if shipping is free.
6. In **E2**, type `=IF(OR(B2="",C2<=0),"",(B2+D2)/C2)` and fill it down to **E6**. Show it to 4 decimals, because the differences between suppliers are small.
7. In **I2**, type how many ounces you need to buy, such as 200.
8. In **F2**, type `=IF(E2="","",ROUND(E2*$I$2,2))` and fill it down to **F6**. This is what your amount would cost from each supplier.
9. In **G2**, type `=IF(E2="","",RANK(E2,$E$2:$E$6,1))` and fill it down to **G6**. Rank 1 is the lowest cost per ounce.
10. In **I4**, type `=IFERROR(INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)),"Enter prices")`.

Here is how I4 works. MIN finds the lowest cost per ounce. MATCH with 0 as the last argument finds the row where that exact number sits. INDEX returns the supplier name from that row.

Use one unit in every row. If one quote is in pounds, convert it to ounces first.

For the summary, put these in **I5:I8**: `=IF(COUNT(E2:E6)=0,"",MIN(E2:E6))` for the lowest cost per ounce, the same with MAX for the highest, `=IF(COUNT(F2:F6)=0,"",MIN(F2:F6))` for the cheapest total, and `=IF(COUNT(F2:F6)=0,"",MAX(F2:F6)-MIN(F2:F6))` for the gap between the cheapest and priciest total.

### In Excel

The formulas work as typed. Format **E2:E6** and **I5:I6** with 4 decimal places and **F2:F6** and **I7:I8** as Currency from **Home > Number Format**. No Microsoft help page is linked here because the steps are plain formulas. This was not opened in Excel for this post.

### In Google Sheets

The same formulas should work as typed. Format the cells with **Format > Number > Custom number format** and enter `$0.0000` for the per-ounce cells. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a candle maker comparing five wax quotes for 200 oz.

| Supplier | Price | Size (oz) | Shipping | Cost per oz | Cost for 200 oz | Rank |
|---|---|---|---|---|---|---|
| Pine Ridge Wax Co | $42.50 | 160 | $12.00 | $0.3406 | $68.13 | 2 |
| Lakeside Supply | $27.99 | 80 | $9.50 | $0.4686 | $93.73 | 4 |
| Harbor Craft Wholesale | $64.00 | 240 | $0.00 | $0.2667 | $53.33 | 1 |
| Birchwood Candle Goods | $18.50 | 48 | $7.25 | $0.5365 | $107.29 | 5 |
| Maple Lane Supply | $33.00 | 100 | $8.00 | $0.4100 | $82.00 | 3 |

Pine Ridge works out to ($42.50 + $12.00) / 160 = $0.340625 per ounce, shown as $0.3406. The Cost for 200 oz column rounds the unrounded per-ounce figure times 200 to cents.

| Result | Cell | Value |
|---|---|---|
| Cheapest supplier | I4 | Harbor Craft Wholesale |
| Lowest cost per oz | I5 | $0.2667 |
| Highest cost per oz | I6 | $0.5365 |
| Cost of 200 oz, cheapest | I7 | $53.33 |
| Extra paid with the priciest | I8 | $53.96 |

The smallest bag has the lowest sticker price at $18.50 and the highest cost per ounce. Free shipping on the 240 oz bag pushes Harbor to the top.

## Troubleshooting

### The supplier's row shows blank cost, rank and total
The package size in **C** is 0 or the price in **B** is empty, so the formula returns an empty cell on purpose. In the recalculated file, a 0 in **C2** blanked E2, F2 and G2, and that row dropped out of MIN, MATCH and RANK. Enter the real size and price.

### I4 says "Enter prices"
Every price cell is blank, so there is nothing to compare. In the recalculated file, clearing all five prices gave "Enter prices" in I4 and blanks in I5 to I8. Type at least one price and size.

### Two suppliers tie and I4 shows only one name
MATCH returns the first match. In the recalculated file, with Pine Ridge and Harbor both at $0.20, both ranked 1 and I4 returned Pine Ridge Wax Co, the first one listed. Check the RANK column for other rank 1 rows.

### A price typed with a dollar sign does not calculate
A price stored as text can break the math. In the recalculated file, "$64.00" typed into **B4** was converted by LibreOffice and still gave $0.2667, but that depends on the app. Keep prices as plain numbers and format the cells as Currency.

### The cheapest supplier is not the best deal
This sheet compares price and shipping per package only. It leaves out minimum orders, free shipping thresholds, quality, delivery time and shipping on several packages. Use it to narrow the list, then judge the rest yourself.

## Template

[Download the .xlsx](templates/tidy-tabs-supplier-price-compare.xlsx)

The file has two tabs. **Compare** is set to print landscape on one US Letter page. It holds the five sample quotes in A2:D6, the formulas in E2:G6, and the yellow input for the amount needed in I2 with the results in I4:I8. **How to use** has short fill-in notes.

The suppliers and prices are fictional. The file has room for five suppliers, so extend the ranges if you add more. It was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
