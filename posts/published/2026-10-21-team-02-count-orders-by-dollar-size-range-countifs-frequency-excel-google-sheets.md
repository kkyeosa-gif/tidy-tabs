---
threads_url: https://www.threads.com/@tin_ylab/post/DexOL8ylGo4
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-count-orders-by-dollar-amount.html
title: How to Count Orders by Dollar Amount Range in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model on the sample and about a dozen edited copies (amounts exactly on a limit, $0.00 and a refund, text amounts, blank rows, limits out of order, an amount above the last limit, sub-cent amounts). Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Count orders by dollar amount range with COUNTIFS in Excel and Google Sheets. Free order size template with counts and sales per range.
image_prompts: A cash box with a few folded dollar bills and a roll of kraft paper bags beside a stack of blank cardstock price tags on a farmers market table, soft morning light, no screens, no readable text, no logos
image_alt: Open cash box beside kraft paper bags and blank price tags on a market table
threads: Type only the lower limit of each order size range. The upper limit is the next row's lower limit, so no order is counted twice.\n=COUNTIFS(Orders!$C$2:$C$200,">="&A2,Orders!$C$2:$C$200,"<"&B2)\nIn the sample, 16 of 40 orders are under $25.00.
---
Count the orders in a range with `=COUNTIFS(Orders!$C$2:$C$200,">="&A2,Orders!$C$2:$C$200,"<"&B2)`, where A2 is the lower limit and B2 the upper limit. An order exactly on a limit counts in the higher range.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Order size table with counts, sales and shares for four dollar ranges](images/2026-10-21-team-02-count-orders-by-dollar-size-range-countifs-frequency-excel-google-sheets/order-size-ranges-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: type only the lower limit of each range. The upper limit is the next row's lower limit, so every order lands in exactly one range and none is counted twice.

1. On the **Orders** tab, put the order amounts in column C, from **C2** down. The formulas read down to row 200.
2. On the **Ranges** tab, type the lower limits in **A2:A5**: 0, 25, 50 and 100.
3. In **B2**, type `=A3`. Copy it down to **B4**. In **B5**, type the text "and up", because the last range has no upper limit.
4. In **D2**, type `=COUNTIFS(Orders!$C$2:$C$200,">="&A2,Orders!$C$2:$C$200,"<"&B2)`. Copy it down to **D4**.
5. In **D5**, type `=COUNTIFS(Orders!$C$2:$C$200,">="&A5)`. This last range has only the ">=" test.
6. In **D6**, type `=COUNTIFS(Orders!$C$2:$C$200,"<"&A2)`. This row catches anything below the first limit, such as a refund.
7. In **D7**, type `=SUM(D2:D6)` for the total.

For the dollars in each range, use the same two tests with SUMIFS. In **E2**, type `=SUMIFS(Orders!$C$2:$C$200,Orders!$C$2:$C$200,">="&A2,Orders!$C$2:$C$200,"<"&B2)`.

Three details trip people up:

- The tests are glued to the cell with `&`. `">="&A2` becomes ">=0" and `"<"&B2` becomes "<25", so changing a limit changes the count.
- Use `>=` for the lower limit and `<` for the upper. If both used `<=` and `>=`, an order of exactly $25.00 would be counted twice.
- Lock the amount range with `$`, as in `$C$2:$C$200`, so it does not slide when you copy the formula down.

### Add a cross-check with FREQUENCY

FREQUENCY counts the same ranges a different way. It counts amounts "up to and including" each bin top, so the template sets each bin top to the limit minus $0.01, for example 24.99. In **A18**, type `=IF(ISNUMBER(A2),ROUND(A2-0.01,2),0)` and copy it to **A21** for the other limits.

Select **C18:C22** and type `=IF(COUNT(Orders!$C$2:$C$200)=0,0,IFERROR(FREQUENCY(Orders!$C$2:$C$200,A18:A21),""))` as one array formula. Column E then says OK where the FREQUENCY count equals the COUNTIFS count.

This block is optional. FREQUENCY only matches COUNTIFS for amounts in whole cents, so an amount such as $24.995 shows "Differs".

### In Excel

The formulas work as typed. Microsoft describes them on its [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842), [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [FREQUENCY function](https://support.microsoft.com/en-us/office/frequency-function-44e3be2b-eca0-42cd-a3f7-fd9ea898fdb9) pages. Microsoft notes that older versions of Excel need Ctrl+Shift+Enter to enter FREQUENCY as an array formula. That was not run in Excel for this post.

### In Google Sheets

The same formulas should work as typed. Google documents the function on its [COUNTIFS help page](https://support.google.com/docs/answer/3256550). Sheets was not opened for this post, so compare your counts with the total in **D7**.

## Example

Sample data is fictional: 40 farmers market and Etsy orders, ORD-5001 to ORD-5040, from 09/01/2026 to 10/10/2026, in amounts from $4.00 to $140.00. Three orders sit exactly on a limit: $25.00, $50.00 and $100.00.

| Range | Orders | Sales | Share of orders | Share of sales |
|---|---|---|---|---|
| $0.00 up to $25.00 | 16 | $237.49 | 40.0% | 14.5% |
| $25.00 up to $50.00 | 12 | $417.99 | 30.0% | 25.6% |
| $50.00 up to $100.00 | 9 | $613.50 | 22.5% | 37.5% |
| $100.00 and up | 3 | $365.00 | 7.5% | 22.3% |
| Below $0.00 | 0 | $0.00 | | |
| Total | 40 | $1,633.98 | | |

The $25.00 order is counted in the $25.00 up to $50.00 row, not the $0.00 row. That is the edge rule from the opening formula.

In this fictional file, most orders are small but the $50.00 and up ranges bring in about 60% of the dollars. That describes the sample only, not what a good mix is.

**B13** on the Ranges tab shows OK when the limits are in order, no amount is text, the counts add up to the number of numeric amounts, and the sales add up to the sum of the column. The FREQUENCY column gives 0, 16, 12, 9 and 3 for the sample, and every cross-check cell says OK.

## Troubleshooting

### An order on the limit is in the higher range
That is how the formula is built. An order of exactly $25.00 passes ">=25" and fails "<25", so it belongs to the $25.00 up to $50.00 range. To put it in the lower range instead, you would need to flip the tests, which is not what the template does.

### The counts do not add up to my orders
**B13** shows CHECK, and the cells above it name the cause. **B10** is more than 0 when an amount is stored as text, such as "$12.00" or "n/a", and COUNTIFS skips text. **B11** says Check limits when the lower limits are blank, text, equal or out of order. Retype text amounts as numbers or fix the limits.

### A refund or a $0.00 order seems to be missing
It is counted in the **Below** row (row 6), not in the ranges above it. A refund of -$12.50 and a -$0.01 adjustment both land there with -$12.51 in sales, and the total in **D7** still matches.

### FREQUENCY and COUNTIFS differ
An amount with a fraction of a cent, such as $24.995, falls between the whole-cent bin tops, so FREQUENCY puts it in a different range than COUNTIFS does. The cross-check cell says Differs and **B13** says CHECK. Round the amounts to cents and both agree.

## Template

[Download the .xlsx](templates/tidy-tabs-order-size-ranges.xlsx)

The file has three tabs. **Orders** holds the yellow inputs in **A2:C200** (40 fictional orders in rows 2 to 41). **Ranges** holds the lower limits in **A2:A5**, the counts, sales and shares in **D2:G7**, the summary in **B9:B13** and the FREQUENCY cross-check in **A17:E22**. **How to use** has short notes. All tabs print portrait on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. It has no chart. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
