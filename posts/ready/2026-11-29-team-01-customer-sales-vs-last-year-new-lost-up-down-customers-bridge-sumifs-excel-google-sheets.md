---
title: How to Compare Each Customer's Sales to Last Year in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 240 workbooks and 608,400 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The "<=" end-date result and the * wildcard result were seen in LibreOffice Calc only.
search_description: Sum each customer's sales for the same dates this year and last, label them New, Lost, Up or Down, and build a bridge that adds up to this year's total.
image_prompts: A small candle workshop's packing bench with a row of wholesale shipping cartons, two of them freshly taped and one empty and pushed aside, a clipboard holding a printed customer order list with some lines checked in pencil, and a few amber glass candle jars waiting to be packed, warm morning light, realistic photo, no screens, no readable text, no logos
image_alt: Taped wholesale candle cartons, an empty box set aside and a pencil-checked clipboard
threads: $4,550.60 last year, $4,180.49 this year, and the total does not say who left.\nSum each customer over the same dates in both years and label them New, Lost, Up or Down. The labels walk one total to the other.\nIn the sample, one lost farm stand took $1,545.60 and two new shops brought $1,529.99.
---
Sum each customer's sales over the same dates in both years with `=SUMIFS(Orders!C:C,Orders!B:B,A10,Orders!A:A,">="&$B$4,Orders!A:A,"<"&$B$5+1)`, label each one New, Lost, Up or Down, and add the labels up so last year's total walks to this year's. In the made-up candle maker's wholesale log, $4,550.60 last year plus $1,529.99 from new shops, minus $1,545.60 from one lost farm stand, plus $476.25 up and minus $830.75 down lands on $4,180.49 this year.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (240 workbooks, 608,400 cell comparisons, 0 mismatches), plus the edge cases below. The `"<="` end-date result and the `*` wildcard result below were seen in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Eight wholesale customers labeled New, Lost, Up or Down beside a $4,550.60 to $4,180.49 bridge](images/2026-11-29-team-01-customer-sales-vs-last-year-new-lost-up-down-customers-bridge-sumifs-excel-google-sheets/customer-sales-vs-last-year-bridge-template.png) *Customers tab of a made-up candle maker's wholesale log as of 11/28/2026: last year and this year per customer for the same dates, the change, % change and the status (Up and New in green, Down and Lost in orange), and the bridge from $4,550.60 last year through +$1,529.99 new, +$476.25 up, -$830.75 down and -$1,545.60 lost to $4,180.49 this year in green. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

A yearly total hides who came and who left: one lost customer and one new one of the same size look like no change. The labels describe sales only. They do not say why a customer left or what to do, and nothing here is a forecast.

The steps are the same in Excel and Google Sheets.

1. On the **Orders** tab, paste one order per row from row 5: Date in **A**, Customer in **B**, Amount in **C**, from 01/01 of last year on. A time of day in the date is fine.
2. On the **Customers** tab, type the as-of date in **B1** (for example 11/28/2026). The sheet does not use TODAY.
3. Pick **Same period** (01/01 through the same month and day in both years) or **Full years** (all of last year) in **B2**. To use another last-year end date, type it in **B3**.
4. Gray cells **B4:B7** show the windows: **B4** is `=DATE(YEAR(B1)-1,1,1)`, **B5** is `=DATE(YEAR(B1)-1,MONTH(B1),DAY(B1))` for Same period, **B6** is 01/01 of this year and **B7** is **B1**.
5. From row 10, type each customer's name in column **A**, spelled as in the order log, one row per customer (up to 100). To build the list from the log, see [How to Count Unique Values in a Column](https://tidytabs.blogspot.com/2026/09/how-to-count-unique-values-in-column-in.html).
6. Column **B** (Last year) holds the full version of the formula above: `=IF(OR(TRIM(A10)="",$I$21<>"OK"),"",ROUND(SUMIFS(Orders!$C$5:$C$2004,Orders!$B$5:$B$2004,A10,Orders!$A$5:$A$2004,">="&$B$4,Orders!$A$5:$A$2004,"<"&$B$5+1),2))`. Column **C** (This year) is the same with `$B$6` and `$B$7`. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
7. Column **D** is the change, This year minus Last year, and **E** the % change, blank when last year is $0.00.
8. Column **F** (Status) labels each customer. After its checks for settings, spaces, special characters and duplicates, the label part is `=IF(AND(B10<=0,C10<=0),"No sales",IF(B10<=0,"New",IF(C10<=0,"Lost",IF(C10>B10,"Up",IF(C10<B10,"Down","Same")))))`. Zero or less counts as nothing, so a customer whose refunds cancel the orders is New, Lost or No sales.
9. Read the bridge in **H9:J17**. Each line sums the Change column by label, for example `=SUMIF($F$10:$F$109,"New",$D$10:$D$109)`, and column **J** counts the customers. Down and Lost show as negative amounts, and **I17** adds last year's total to every line. See the SUMIF help from [Microsoft](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) or [Google](https://support.google.com/docs/answer/3093583).
10. Check **CHECK** in **I24**. It is OK only when the settings are valid, no customer row shows Check, the Orders CHECK is OK, and the bridge, the sum of the customer rows and the whole order log agree to the cent for both years.

Each window ends with `"<"&$B$5+1`, before the next day starts, so an order with a time of day on the last day still counts.

Related: [this year vs last year month by month](https://tidytabs.blogspot.com/2026/10/how-to-compare-this-years-sales-to-last.html), [new and returning customers each month](https://tidytabs.blogspot.com/2026/11/how-to-count-new-and-returning.html) and [customers who have not ordered in 90 days](https://tidytabs.blogspot.com/2026/10/how-to-flag-customers-who-have-not.html).

## Example

The data is made up: a candle maker's wholesale log of 23 orders for eight shops, 02/10/2025 to 11/29/2026, as of 11/28/2026 in the Same period view (01/01 to 11/28 in both years).

| Customer | Last year | This year | Change | % change | Status |
|---|---|---|---|---|---|
| Birch & Bloom | $800.00 | $1,220.50 | $420.50 | 52.6% | Up |
| Cedar Lane Cafe | $1,500.00 | $789.25 | -$710.75 | -47.4% | Down |
| Harbor Gifts | $430.00 | $310.00 | -$120.00 | -27.9% | Down |
| Juniper Co. | $0.00 | $649.99 | $649.99 | | New |
| Maple Street Bakery | $275.00 | $330.75 | $55.75 | 20.3% | Up |
| North Pine Studio | $0.00 | $880.00 | $880.00 | | New |
| Oak Hollow Farm | $1,545.60 | $0.00 | -$1,545.60 | -100.0% | Lost |
| Willow Market | $0.00 | $0.00 | $0.00 | | No sales |

The bridge reads:

| Last year to this year | Amount | Customers |
|---|---|---|
| Last year total | $4,550.60 | 8 |
| + New customers | $1,529.99 | 2 |
| + Up | $476.25 | 2 |
| - Down | -$830.75 | 2 |
| - Lost customers | -$1,545.60 | 1 |
| Same | $0.00 | 0 |
| No sales (refunds only) | $0.00 | 1 |
| This year, by the bridge | $4,180.49 | |

The new shops' $1,529.99 nearly replaces Oak Hollow Farm, so the $370.11 drop comes mostly from Cedar Lane Cafe and Harbor Gifts buying less. CHECK shows OK: $4,180.49 is also the sum of the customer rows and of the whole log for 01/01/2026 to 11/28/2026.

Cedar Lane Cafe's $64.00 order at 3:30 PM on 11/28/2026 counts. Harbor Gifts' $80.00 on 11/29/2025 and North Pine Studio's $150.00 on 11/29/2026 fall outside both windows, and the Orders tab marks them "Outside both periods (not counted)".

With Full years in **B2**, **B5** becomes 12/31/2025 and last year totals $5,770.60: Willow Market's $500.00 on 12/01/2025 makes it Lost, and Maple Street Bakery turns Down ($915.00 to $330.75). CHECK stays OK.

## Troubleshooting

### This year's total is short of the order log
The window probably ends with `"<="&B7`. In LibreOffice Calc, `=SUMIFS(amount,date,">="&B6,date,"<="&B7)` on the sample gave $4,116.49 instead of $4,180.49: the 3:30 PM Cedar Lane Cafe order on 11/28/2026 dropped out, because 11/28/2026 3:30 PM is later than 11/28/2026 at midnight. Ending with `"<"&B7+1`, as the template does, kept it.

### CHECK says Check and an order's Status says "Check: customer not on the Customers tab"
The log name matches no name on the list, often because of a trailing space such as "Harbor Gifts " (capital letters do not matter). Fix the spelling or add the customer.

### A customer shows "Check: * ? ~ = < or > in the name"
SUMIFS reads those characters as wildcards or comparisons. In LibreOffice Calc, a list name "birch*" picked up Birch & Bloom's $1,220.50, so the list total came out higher than the log. Take the character out of the name in both tabs.

### A customer that bought last year shows No sales, not Lost
The Same period view only counts orders through the same month and day last year. Willow Market bought only on 12/01/2025, so it has nothing in either window through 11/28. Choose **Full years** in **B2** to count all of last year.

## Template

[Download the .xlsx](templates/tidy-tabs-customer-sales-vs-last-year.xlsx)

**Customers** holds the settings in **B1:B7**, up to 100 customers (rows 10 to 109), the bridge in **H9:J17** and the checks in **H19:I24**. **Orders** takes 2,000 orders (rows 5 to 2004) with a Status for each row and its own CHECK in **G11**, which also catches rows typed below row 2004. **How to use** explains each formula. Data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. New and Up labels are light green, Down and Lost light orange, orders outside both periods light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
