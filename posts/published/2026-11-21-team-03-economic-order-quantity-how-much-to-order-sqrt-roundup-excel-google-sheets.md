---
threads_url: https://www.threads.com/@tin_ylab/post/DgFdLXZEX_E
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-find-how-much-to-order-at-once.html
title: How to Find How Much to Order at Once (Economic Order Quantity) in Excel and Google Sheets
labels: excel-formulas, inventory, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 44 workbooks and 24,948 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Find how much to order at once with =SQRT(2*units*cost_per_order/(unit_cost*holding_pct)), then round up to your case size with ROUNDUP.
image_prompts: Cardboard cases of empty glass candle jars stacked on a metal shelf in a garage candle studio, a large bag of soy wax flakes and a roll of packing tape on the concrete floor beside them, soft daylight from a garage window, realistic photo, no screens, no readable text, no logos
image_alt: Stacked cases of empty glass jars on a shelf beside a bag of wax flakes
threads: 707.11 jars is the math answer. 708 is the order.\nEOQ is =SQRT(2*units_per_year*cost_per_order/(unit_cost*holding_pct)), then ROUNDUP to the case of 12.\nThe curve is flat, too: 600 or 800 jars cost $172.00 or $171.00 a year vs $169.71.
---
Find the order size with the lowest yearly total of ordering and holding cost with `=SQRT(2*units_per_year*cost_per_order/(unit_cost*holding_pct))`, then round it up to your case size with `=ROUNDUP(eoq/case,0)*case`. In the made-up supplies list, 2,400 glass jars a year at $25.00 per order, $1.20 each and a 20% holding cost give an EOQ of 707.11, so you order 708 jars (a case of 12) about every 107.7 days for a yearly total of $169.71, and the download's cost table shows 600 or 800 jars costing only $172.00 or $171.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (44 workbooks, 24,948 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (SQRT, ROUNDUP, ROUND, IF, AND, ISNUMBER, MIN, INDEX, MATCH) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Five candle supplies with yearly use, order costs, EOQ, case-rounded order quantity and yearly totals](images/2026-11-21-team-03-economic-order-quantity-how-much-to-order-sqrt-roundup-excel-google-sheets/economic-order-quantity-supplies-template.png) *Supplies tab with five fictional candle-making supplies, where glass jars at 2,400 a year and $25.00 per order give an EOQ of 707.11, rounded up to 708 for cases of 12, for a yearly total of $169.71. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Supplies** tab, one supply per row starting in row 5.

1. In **A5:D5**, type the item, units used per year (not per month), cost per order (shipping, handling, your time) and unit cost.
2. In **E5**, type your holding cost per year as a percent, such as 20%: your own estimate of what a dollar of stock costs to keep for a year (cash tied up, space, breakage). The sheet suggests no rate.
3. In **F5**, type the case size, such as 12, or 1 for single units.
4. In **G5**, type `=D5*E5` for the holding cost per unit per year.
5. In **H5**, type `=SQRT(2*B5*C5/G5)` for the EOQ.
6. In **I5**, type `=ROUNDUP(ROUND(H5/F5,9),0)*F5` to round up to whole cases. The inner ROUND keeps an EOQ of exactly 1,000 at a case of 100 from being pushed up a case by a leftover decimal.
7. In **J5:N5**, type orders per year `=B5/I5`, days between orders `=365/J5`, yearly order cost `=J5*C5`, yearly holding cost `=I5/2*G5` (on average half an order sits on the shelf) and yearly total `=L5+M5`.
8. Fill **G5:N5** down. In the template each is wrapped in `IF($O5<>"OK","",...)`, so results stay blank until Status in **O5** says OK.
9. On **Cost table**, pick an item in **B1**. Rows 5 to 24 list 20 order sizes from **B2** in steps of **B3**, the yearly total in **D**, and mark the cheapest Lowest.

Microsoft's [SQRT function](https://support.microsoft.com/en-us/office/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) pages cover both functions.

Treat EOQ as a starting point. It assumes steady use and the same cost for every order, and it ignores quantity discounts, lead time and safety stock. It is not buying or accounting advice. For price breaks, see [looking up wholesale price tiers by quantity](https://tidytabs.blogspot.com/2026/10/how-to-look-up-wholesale-price-tiers-by.html); for when to place the order, see [calculating a reorder point](https://tidytabs.blogspot.com/2026/10/how-to-calculate-reorder-point-in-excel.html); for minimum orders on top of cases, see [rounding an order quantity up to a case size](https://tidytabs.blogspot.com/2026/11/how-to-round-order-quantity-up-to-case.html).

## Example

The supplies and prices are fictional, for a small candle shop.

| Item | Per year | Per order | Unit cost | Holding | Case | EOQ | Order qty | Days between | Yearly total |
|---|---|---|---|---|---|---|---|---|---|
| Glass jars, 8 oz | 2,400 | $25.00 | $1.20 | 20% | 12 | 707.11 | 708 | 107.7 | $169.71 |
| Candle wicks | 5,000 | $8.00 | $0.10 | 25% | 100 | 1,788.85 | 1,800 | 131.4 | $44.72 |
| Mailer boxes | 1,800 | $15.00 | $0.85 | 20% | 50 | 563.60 | 600 | 121.7 | $96.00 |
| Soy wax, lb | 1,200 | $30.00 | $2.40 | 20% | 50 | 387.30 | 400 | 121.7 | $186.00 |
| Fragrance oil, 16 oz | 300 | $12.00 | $9.50 | 20% | 6 | 61.56 | 66 | 80.3 | $117.25 |

For the jars, 708 means 3.4 orders a year, $84.75 in order cost and $84.96 in holding cost. At the exact EOQ both halves are $84.85, a quick check of the math. The summary shows $613.67 for all five items and CHECK reads OK.

The jars' Cost table (100 in steps of 100) is flat near the EOQ:

| Order quantity | 300 | 500 | 600 | 700 | 800 | 1,000 | 1,500 |
|---|---|---|---|---|---|---|---|
| Yearly total | $236.00 | $180.00 | $172.00 | $169.71 | $171.00 | $180.00 | $220.00 |

700 is marked Lowest. Ordering 600 or 800 costs only $1.29 to $2.29 more a year, so rounding to a case barely matters; 100 at a time costs $612.00. With **B2** at 600 and **B3** at 12, the lowest row is 708 at $169.71.

## Troubleshooting

### Status shows Check: holding % (20% not 20)
The holding cost was typed as 20, which the sheet reads as 2,000%. Unchecked, the jars' EOQ comes out at 70.71 instead of 707.11, ten times too small. Retype it as 20%.

### The EOQ looks about 3.5 times too small but Status says OK
Monthly use went into the yearly column. Typing 200 jars instead of 2,400 gives an EOQ of 204.12 and an order of 216, and the sheet cannot tell. Multiply monthly use by 12 before you type it.

### Status shows Check: cost per order
The cell is 0, blank, or text like "$25.00" pasted with the dollar sign. In LibreOffice an unchecked 0 gave an EOQ of 0 and #DIV/0! in orders per year. Retype it as a plain number.

### Lowest in table is the first row
The EOQ is below the first quantity in the Cost table. The fragrance oil's EOQ is 61.56, so a table starting at 100 marks 100 at $131.00 as Lowest. Lower **B2** and **B3** for small items.

## Template

[Download the .xlsx](templates/tidy-tabs-economic-order-quantity.xlsx)

The file has three tabs. **Supplies** has inputs in **A:F** and results in **G:O** for rows 5 to 54, with a summary in **Q4:R8** (items listed, rows showing Check, yearly total of OK items, entries below row 54, CHECK). **Cost table** has the item picker in **B1**, first quantity and step in **B2:B3**, 20 order sizes in rows 5 to 24, and a summary in **G4:H15** with the item's EOQ, order quantity and the lowest-cost row. **How to use** explains each formula. Both sheets print Landscape on US Letter.

Yellow cells are inputs and gray cells are formulas. Every Check status and a CHECK that is not OK turn red, and the Lowest row turns green. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
