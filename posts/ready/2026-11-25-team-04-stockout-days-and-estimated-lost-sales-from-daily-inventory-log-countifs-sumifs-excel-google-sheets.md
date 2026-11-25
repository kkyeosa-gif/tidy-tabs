---
title: How to Count Stockout Days and Estimate Lost Sales in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 207 workbooks and 2,640,699 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Flag each day a product opened with nothing to sell, count out days with COUNTIFS, and estimate lost sales from the in-stock daily average with SUMIFS.
image_prompts: A wooden retail shelf in a small candle shop with one empty gap between rows of glass jar candles, a single empty cardboard tray left in the gap and a handheld clipboard with a blank count sheet leaning on the shelf edge, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Empty gap on a wooden shelf between rows of glass jar candles
threads: 10 out days in 28, and the quick average says you lost 8 candles.\nAverage only the in-stock days instead: 1.278 a day, not 0.821.\nThat makes it about 13 lost units, or $312.00 at $24.00.
---
Mark a day "Out" when it opened with nothing to sell, `=IF(on_hand+received<=0,"Out","In stock")`, count those days with `=COUNTIFS(sku,s,status,"Out")`, and estimate lost units as out days times the average sold on in-stock days only, `=SUMIFS(sold,sku,s,status,"In stock")/COUNTIFS(sku,s,status,"In stock")`. In the made-up log, a cedar candle out for 10 of 28 days sold 1.278 a day while in stock, so about 13 lost units, or $312.00 at $24.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (207 workbooks, 2,640,699 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Three candle products with out days, in-stock sales rate and estimated lost units and dollars](images/2026-11-25-team-04-stockout-days-and-estimated-lost-sales-from-daily-inventory-log-countifs-sumifs-excel-google-sheets/stockout-lost-sales-template.png) *Summary tab of a fictional 28-day inventory log: the cedar candle was out 10 days and sold 1.278 a day while in stock (0.821 by the all-days shortcut), so the estimate is 13 lost units and $312.00, against 8 units by the shortcut. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The result is an estimate, not a forecast. It assumes the product would have kept selling at its in-stock rate. It leaves out shoppers who bought another scent instead and demand that was rising or falling, and it is not an accounting figure.

The log needs one row per SKU per day, out days included. A day with no row is not counted at all, so it is never an out day.

The steps are the same in Excel and Google Sheets.

1. On the **Products** tab, type one product per row from row 5: SKU in **A**, Name in **B** and Price in **C**, such as CED-8, Cedar candle 8 oz, $24.00.
2. On the **Daily log** tab, type one row per SKU per day from row 6: Date in **A**, SKU in **B**, Units on hand at open in **C**, Units received in **D** and Units sold in **E**. Leave **D** or **E** blank for 0.
3. In **F6**, add what was available to sell that day with `=IF(I6<>"OK","",C6+N(D6))`. N turns a blank into 0.
4. In **G6**, get what was left at close with `=IF(F6="","",F6-N(E6))`.
5. In **H6**, flag the day with `=IF(F6="","",IF(F6<=0,"Out","In stock"))`.
6. Column **I** checks each row (text dates, unknown SKUs, negative numbers, more sold than available, a SKU and date typed twice). Only rows showing OK are counted.
7. On the **Summary** tab, count out days in **E5** with `=IF(A5="","",COUNTIFS('Daily log'!$B$6:$B$3005,A5,'Daily log'!$H$6:$H$3005,"Out"))`. **F5** is the same formula with "In stock", and **D5** adds the two for Days logged. See the [COUNTIFS help from Microsoft](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842).
8. In **G5**, total units sold on in-stock days with `=IF(A5="","",SUMIFS('Daily log'!$E$6:$E$3005,'Daily log'!$B$6:$B$3005,A5,'Daily log'!$H$6:$H$3005,"In stock"))`. See the [SUMIFS help from Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b).
9. In **H5**, get the in-stock daily rate with `=IF(A5="","",IF(F5=0,"No in-stock days",G5/F5))`.
10. In **J5**, estimate lost units with `=IF(A5="","",IF(E5=0,0,IF(F5=0,"No in-stock days",ROUND(E5*H5,0))))`.
11. In **K5**, turn that into dollars with `=IF(A5="","",IF(OR(NOT(ISNUMBER(J5)),C5=""),"",ROUND(J5*C5,2)))`.
12. Column **M** counts in-stock days that ended at 0. Those days sold out partway through and count as in stock, so the estimate runs low.

## Example

The sample data is made up. The log runs 10/26/2026 to 11/22/2026, 28 days per SKU, 84 rows. Lavender starts with 20 and gets 24 on 11/09/2026; cedar starts with 12 and gets 24 on 11/15/2026; pumpkin starts with 30.

| SKU | Price | Days logged | Out days | In-stock days | Sold on in-stock days | Per in-stock day | Per day, all days | Lost units | Lost sales | Ended at 0 |
|---|---|---|---|---|---|---|---|---|---|---|
| LAV-8 | $24.00 | 28 | 0 | 28 | 33 | 1.179 | 1.179 | 0 | $0.00 | 0 |
| CED-8 | $24.00 | 28 | 10 | 18 | 23 | 1.278 | 0.821 | 13 | $312.00 | 1 |
| PUMP-4 | $16.00 | 28 | 1 | 27 | 30 | 1.111 | 1.071 | 1 | $16.00 | 1 |

The cedar candle sold its last unit on 11/04/2026, was out 11/05/2026 to 11/14/2026, and came back on 11/15/2026 when 24 arrived. The pumpkin candle ended at 0 on 11/21/2026 and was out on 11/22/2026.

The totals box in **P4:Q10** shows 11 out days, 14 estimated lost units, $328.00 in estimated lost sales, 0 log rows and 0 product rows showing Check, and CHECK OK.

Edge cases that matched the Python model:

| Case | Result |
|---|---|
| A SKU at 0 on all 28 days | No in-stock days in H and J, K blank, Status Out every logged day |
| Cedar with 11/07 to 11/09/2026 deleted | 25 days logged, 7 out days, 9 lost units, $216.00 |
| 0 on hand plus 24 received | In stock |

## Troubleshooting

### The lost-sales estimate looks too small
You divided total units sold by all days. The out days add zeros to the average you are trying to measure, so cedar drops from 1.278 a day to 0.821 (36% lower) and the estimate drops from 13 units to 8. Column **L** on Summary shows the shortcut's number next to the right one for comparison.

### #DIV/0! for a product that was always out
With no in-stock days, there is nothing to divide by. The template shows No in-stock days in **H** and **J**, leaves the dollar cell blank, and sets Status to Out every logged day.

### A day is missing from the counts
Its row shows a Check message in column **I**, such as Check: sold more than available, Check: SKU and date typed twice (both rows are flagged) or Check: SKU not on Products. **Q7** counts these rows.

### Out days look too low after a gap in the log
A date with no row is skipped, not counted as out. Deleting three of cedar's out days gave 7 out days and 9 lost units instead of 10 and 13. Type a row for every day, with 0 on hand when the shelf was empty.

## Template

[Download the .xlsx](templates/tidy-tabs-stockout-lost-sales.xlsx)

The file has four tabs. **Products** holds up to 50 SKUs. **Daily log** holds up to 3,000 rows: you type **A:E**, and **F:I** hold the formulas above. **Summary** fills one row per product with nothing to type, plus the totals and CHECK box in **P4:Q10**, which reads OK only when days are logged, no row shows Check and nothing is typed below row 3005 on the log. **How to use** explains each formula.

Yellow cells are inputs and gray cells are formulas. Out days turn light orange, lost-sales cells above 0 turn light green, and every Check turns red. The data tabs print Landscape on US Letter with the header row repeated, and Summary fits on one page.

For the forward-looking version, see [How to Calculate Days of Stock Left From Recent Sales](https://tidytabs.blogspot.com/2026/10/how-to-calculate-days-of-stock-left.html).

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
