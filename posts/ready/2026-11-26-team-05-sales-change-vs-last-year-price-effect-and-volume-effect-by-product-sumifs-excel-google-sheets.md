---
title: How to See If Sales Grew From Price or Volume in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 220 workbooks and 785,620 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Split each product's sales change vs last year into a volume effect and a price effect with SUMIFS and ROUND; new and dropped products get their own line.
image_prompts: A candle shop stockroom shelf with two rows of glass jar candles, one row wearing small blank orange sale stickers and the other row plain, a handheld price labeler resting on the shelf edge, and a short stack of flat empty shipping boxes on the floor below, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Stockroom shelf of jar candles, some with blank sale stickers, beside a price labeler
threads: Sales up 17.7% while units fell from 1,297 to 1,272. So was it price?\nPer product: the change in units times last year's average price is volume, and the rest of the revenue change is price.\nCedar candle: -$483.41 from volume, +$776.81 from price, up $293.40 overall.
---
Value the change in units at last year's average price, `=ROUND((E7-B7)*D7,2)`, and call the rest of the revenue change the price effect, `=ROUND(H7-I7,2)`, so the two always add up. In the made-up candle shop, Cedar candle revenue rose $293.40 even though it sold 22 fewer units: -$483.41 from volume and +$776.81 from an average price that went from $21.97 to $23.98.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (220 workbooks, 785,620 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Six candle shop products with revenue change split into volume and price effects](images/2026-11-26-team-05-sales-change-vs-last-year-price-effect-and-volume-effect-by-product-sumifs-excel-google-sheets/sales-change-price-volume-template.png) *Compare tab of a fictional candle shop for 01/01 to 10/31 of 2026 vs 2025: each product's units, revenue and average price in both periods, the change split into a volume effect and a price effect, a New product and a Dropped row in orange, and the summary with $180.41 from volume, $1,579.89 from price and $2,283.20 from new and dropped products in green. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

This is a simple price volume analysis per product. "Price" here means average price, which is revenue divided by units, so discounts and returns move it. A price effect is not the same as a list price change. The sheet has no separate mix line; you see the split product by product instead.

The steps are the same in Excel and Google Sheets.

1. On the **Sales log** tab, type one line per row from row 5: Date in **A**, Product in **B**, Units in **C**, Revenue in **D**. A return is negative units and negative revenue. Column **E** shows OK, "Outside both periods" or a Check message.
2. On the **Compare** tab, type the two periods in **B1:B4**: this period from and to, then last period from and to. **O7** says OK only when both periods are real dates, in order, and do not overlap.
3. In **A7:A56**, type each product name once, spelled the same as in the log. Case does not matter; extra spaces do.
4. **B7** totals last period's units with `=IF(TRIM(A7)="","",SUMIFS('Sales log'!$C$5:$C$3004,'Sales log'!$B$5:$B$3004,A7,'Sales log'!$E$5:$E$3004,"OK",'Sales log'!$A$5:$A$3004,">="&$B$3,'Sales log'!$A$5:$A$3004,"<="&$B$4))`. **C7** is the same over the Revenue column, and **E7** and **F7** do the same with **$B$1** and **$B$2** for this period. Only log rows with Status OK count. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
5. **D7** gives last period's average price with `=IF(OR(L7="",LEFT(L7,5)="Check",N(B7)<=0),"",C7/B7)`, and **G7** does the same with F7/E7.
6. **H7** is the revenue change, `=IF(OR(L7="",LEFT(L7,5)="Check"),"",ROUND(F7-C7,2))`.
7. **I7** is the volume effect, `=IF(L7<>"OK","",ROUND((E7-B7)*D7,2))`: the extra or missing units priced at last year's average.
8. **J7** is the price effect, `=IF(L7<>"OK","",ROUND(H7-I7,2))`. Because it is the remainder, volume plus price equals the change to the cent. It matches `(G7-D7)*E7` up to rounding: across all 3,162 OK rows of the test workbooks, the gap was at most $0.005. See the ROUND help from [Microsoft](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) or [Google](https://support.google.com/docs/answer/3093440).
9. **K7** holds the whole change for a New product or a Dropped product, `=IF(L7="New product",H7,IF(L7="Dropped",H7,IF(OR(L7="OK",L7="No sales in either period"),0,"")))`. A product with no sales last year has no old price, so splitting it would show a fake price effect.
10. **L7** is the Status. It checks the periods, a product listed twice (it uses COUNTIF, see [Microsoft](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) or [Google](https://support.google.com/docs/answer/3093480)), net units below 0 and revenue with 0 units, then labels the row OK, New product, Dropped or No sales in either period.
11. Read the summary in **N7:O21**: both revenues, the change in $ and %, units, the three parts (volume, price, new and dropped), Change not explained (should be $0.00), and a final **CHECK** cell.

For the same comparison by calendar month, see [How to Compare This Year's Sales to Last Year Month by Month](https://tidytabs.blogspot.com/2026/10/how-to-compare-this-years-sales-to-last.html). To go from revenue to profit, see [How to Calculate Profit by Product From a Sales Log](https://tidytabs.blogspot.com/2026/11/how-to-calculate-profit-by-product-from.html).

## Example

The sales below are made up: a candle shop with one log row per product per month (on the 15th) for January to October of 2025 and 2026, some lines at 10% or 20% off, a Lavender return of -2 units and -$48.00 on 04/02/2026, and 5 rows on 11/15/2025 that fall outside both periods. This period is 01/01/2026 to 10/31/2026; last period is 01/01/2025 to 10/31/2025.

| Product | Last units | Last revenue | Last avg | This units | This revenue | This avg | Change | Volume | Price | New or dropped | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Cedar candle 8 oz | 410 | $9,009.00 | $21.97 | 388 | $9,302.40 | $23.98 | $293.40 | -$483.41 | $776.81 | $0.00 | OK |
| Lavender candle 8 oz | 352 | $7,737.40 | $21.98 | 369 | $8,836.80 | $23.95 | $1,099.40 | $373.68 | $725.72 | $0.00 | OK |
| Pumpkin candle 4 oz | 260 | $3,637.20 | $13.99 | 301 | $4,211.20 | $13.99 | $574.00 | $573.56 | $0.44 | $0.00 | OK |
| Wax melts 6-pack | 180 | $1,700.50 | $9.45 | 150 | $1,494.00 | $9.96 | -$206.50 | -$283.42 | $76.92 | $0.00 | OK |
| Holiday gift set | 0 | $0.00 | | 64 | $3,038.40 | $47.48 | $3,038.40 | | | $3,038.40 | New product |
| Oat soap bar | 95 | $755.20 | $7.95 | 0 | $0.00 | | -$755.20 | | | -$755.20 | Dropped |

Revenue went from $22,839.30 to $26,882.80, up $4,043.50 (17.7%), while units fell from 1,297 to 1,272. The summary splits that into $180.41 from volume, $1,579.89 from price and $2,283.20 from the new gift set minus the dropped soap, with $0.00 not explained and CHECK OK. Pumpkin shows the other extreme: the price held at $13.99, so almost all of its $574.00 came from volume.

On the Sales log, 101 rows are OK ($49,722.10), 5 are Outside both periods and 0 show Check.

## Troubleshooting

### CHECK is red and log rows say "Check: product not on the Compare tab"
The product is missing from column A of Compare, or the log name has a trailing space. Add the product to the list or fix the name. Those rows are left out of every total until they match.

### Every product row says "Check: period dates (B1:B4)"
The two periods overlap, one runs backward (from after to), or a date was typed as text. Fix **B1:B4** until **O7** says OK.

### A row says "Check: last revenue with 0 units" or "Check: this revenue with 0 units"
A log line has revenue but its units net to 0, such as a $15.00 line typed with 0 units. Fix the units in the log. Net units below 0 get their own Check.

### Two rows say "Check: product listed twice"
The same name is in column A twice, possibly in different case, like "cedar candle 8 oz" and "Cedar candle 8 oz". Delete one row.

### The price effect is large but you did not raise prices
Average price is revenue per unit, so fewer discounted sales or fewer returns raise it too. Look at the log for that product before reading it as a price change. The sheet splits the numbers; it does not tell you what to charge.

## Template

[Download the .xlsx](templates/tidy-tabs-sales-change-price-volume.xlsx)

**Compare** holds the two periods, up to 50 products (rows 7 to 56) with the volume, price and new or dropped columns, and the summary. **Sales log** takes 3,000 rows (5 to 3004) with its own summary and CHECK cell, which also catches rows typed below row 3004. **How to use** explains each formula. The data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. The three effects in the summary are light green, New product, Dropped and Outside both periods are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
