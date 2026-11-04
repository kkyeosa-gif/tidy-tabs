---
threads_url: https://www.threads.com/@tin_ylab/post/DfVYL4ZmhSB
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-return-rate-by-product.html
title: How to Calculate Return Rate by Product in Excel and Google Sheets
labels: excel-formulas, small-business, products
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 23 workbooks and 1,288 comparisons with 0 mismatches. Checked: no orders, no products, flag rates of 0%, 5%, 10%, 100%, 150%, -10% and text, a rate exactly at the flag, a tie for the highest rate, a product with no sales, a misspelled product, lower and upper case, a trailing space, returns above sales, text and blank units, and five random sets of up to 300 order lines and 12 products. Excel and Google Sheets were not used.
search_description: Calculate return rate by product in Excel and Google Sheets with SUMIFS. Free template flags products above your cutoff and shows the highest rate.
image_prompts: A returned ceramic mug wrapped in tissue paper beside a stack of opened padded mailers and a roll of packing tape on a packing table, soft daylight, no screens, no readable text, no logos
image_alt: Returned mug in tissue paper beside opened padded mailers on a table
threads: Which product do customers send back the most?\nUnits returned divided by units sold, per product: =SUMIFS(returned,product,A6)/SUMIFS(sold,product,A6).\nIn the sample, the enamel mug is 4 of 14 (28.6%) and gets flagged above a 10% line. The sheet also counts units sold under a name that is not on your list.
---
Divide units returned by units sold with `=C6/B6`, where each count comes from `=SUMIFS(Orders!$D$5:$D$304,Orders!$C$5:$C$304,A6)`. If 4 of 14 enamel mugs came back, the return rate is 28.6%.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. SUMIFS, COUNTIF and IF exist in current Excel and Google Sheets, but neither app was opened for this post.

![Product table with units sold, returned units, return rate, status and an overall summary](images/2026-11-04-team-04-return-rate-by-product-sumifs-excel-google-sheets/return-rate-by-product-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Orders** has one row per product on an order from row 5: order number in **A**, date in **B**, product in **C**, units sold in **D**, units returned in **E**. **Products** lists each product once from **A6**.

1. On **Orders**, log each line. Type 0 in column E when nothing came back.
2. On **Products**, type each product name in **A6:A17**, spelled exactly as on Orders.
3. In **B2**, type the return rate you want to flag, such as 10%.
4. In **B6**, type `=SUMIFS(Orders!$D$5:$D$304,Orders!$C$5:$C$304,A6)` for units sold.
5. In **C6**, type `=SUMIFS(Orders!$E$5:$E$304,Orders!$C$5:$C$304,A6)` for units returned.
6. In **D6**, type `=IF(B6=0,"",C6/B6)` and format it as a percent.
7. In **E6**, type `=IF(B6=0,"No sales",IF(D6>$B$2,"High","OK"))`.
8. Fill **B6:E6** down.

The test `B6=0` matters. Without it, a product with no sales divides by zero.

For all products together, **H7** is `=H6/H5`, with `=SUM(Orders!D5:D304)` in **H5** and `=SUM(Orders!E5:E304)` in **H6**.

This counts units on the order line where you typed them. It does not use the date the item came back, refund amounts or reasons.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/03/2026; the pages were not read for this post.

## Example

Sample data is fictional: 44 order lines from 09/01/2026, a flag at 10.0%.

| Product | Units sold | Units returned | Return rate | Status |
|---|---|---|---|---|
| Lavender candle | 10 | 0 | 0.0% | OK |
| Oatmeal soap | 23 | 1 | 4.3% | OK |
| Canvas tote | 11 | 1 | 9.1% | OK |
| Hoop earrings | 11 | 2 | 18.2% | High |
| Enamel mug | 14 | 4 | 28.6% | High |

Across all products, 94 units were sold and 10 returned, an overall rate of 10.6%. The enamel mug has the highest rate. A rate exactly at the flag is OK: in the test, 1 of 10 against a 10% flag gave OK and 3 of 20 (15%) gave High.

## Troubleshooting

### A product shows fewer sales than expected
The name on Orders does not match the name on Products. **H10** counts the units sold under names that are not on the list. In the test, `Enamel mugg` raised **H10** to 2 and the **CHECK** cell said Check.

### Upper and lower case do not matter, spaces do
In the test, `enamel mug` and `ENAMEL MUG` were counted with Enamel mug. A trailing space (`Enamel mug `) was a different product and landed in **H10**.

### A product has more units returned than sold
A line that returns more than it sold is flagged. In the test, typing 99 returned on one Oatmeal soap line gave 100 returned against 23 sold (434.8%) for that product and the **CHECK** cell said Check.

### A product shows No sales and a blank rate
Nothing was sold under that name, so there is no rate to show. In the test, a product with 0 units sold showed No sales and no error.

### B3 shows a message and the statuses say Check settings
The flag in **B2** must be a number from 0% to 100%. In the test, 150%, -10% and text each put a message in **B3** and Check settings in the statuses, with no error value.

## Template

[Download the .xlsx](templates/tidy-tabs-return-rate-by-product.xlsx)

The file has three tabs. **Products** has the flag in **B2**, a settings check in **B3**, 12 product rows from row 6 (yellow names in **A**, formulas in **B:E**) and a summary in **H5:H11**: units sold, units returned, overall rate, the highest rate, products above the flag, units under products not listed and a **CHECK** cell. **Orders** has 300 rows from row 5. **How to use** has notes on each formula.

**CHECK** says OK when every sold unit belongs to a listed product, no line returned more than it sold and the units are numbers. Fix every Check before you compare products.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
