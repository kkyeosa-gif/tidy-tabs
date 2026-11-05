---
threads_url: https://www.threads.com/@tin_ylab/post/DfXoYeJiQO7
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-count-orders-by-status-and-find.html
title: How to Count Orders by Status and Find the Oldest Open Order in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 24 workbooks and 1,654 comparisons with 0 mismatches. Checked: the sample, no orders, a misspelled status, a trailing space, lower and upper case, a blank status, an order dated after the as-of date, a same-day order, a tie for the oldest, 7 and 8 days open, a text date, a blank date, an as-of date typed as text or blank, all orders closed, custom status names, a duplicate status in the list, Maybe instead of Yes or No, and six random sets of 5 to 200 orders. Excel and Google Sheets were not used. Not run: TODAY(), more than 200 orders, and the status dropdown.
search_description: Count orders by status with COUNTIF and name the oldest open order with INDEX and MATCH in Excel and Google Sheets. Free template included.
image_prompts: Several taped parcels on a wooden shelf, a few marked with colored sticky flags, next to a small postal scale in a craft studio, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Packed parcels with colored sticky flags on a shelf beside a packing scale
threads: Name your oldest open order with INDEX, MATCH and MAX.\nIn the sample, order 3103 has been open 20 days as of 11/04/2026.\nThe free template also counts the open orders older than 7 days.
---
Count the orders in one status with `=COUNTIF(Orders!$C$6:$C$205,A5)`, then name the oldest open order with `=INDEX(Orders!A6:A205,MATCH(MAX(Orders!E6:E205),Orders!E6:E205,0))`. In the sample, order 3103 has been open 20 days as of 11/04/2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIF, INDEX, MATCH, MAX and IF exist in current Excel and Google Sheets, but neither app was opened for this post.

![Status counts table showing eleven open orders and the oldest open order at twenty days](images/2026-11-05-team-02-order-status-counts-oldest-open-order-countif-excel-google-sheets/order-status-counts-oldest-open-order-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*


## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Orders** has one row per order from row 6: order number in **A**, order date in **B**, status in **C**. **Summary** lists your five statuses in **A5:A9**, with Yes or No in **B5:B9** for whether the order still needs work.

1. On **Summary**, type your five statuses in **A5:A9**. Next to each, type Yes in **B** if the order is still open, or No if it is done.
2. On **Orders**, type the as-of date in **B2**, such as 11/04/2026. It is typed, not `TODAY()`, so the days do not change when you reopen the file.
3. Log each order in rows 6 to 205. Type the status in **C** exactly as it is on Summary, or add a dropdown with **Data > Data Validation** (Excel) or **Data > Data validation** (Google Sheets) that points to `Summary!$A$5:$A$9`.
4. In **D6**, type `=IF(C6="","",IFERROR(INDEX(Summary!$B$5:$B$9,MATCH(C6,Summary!$A$5:$A$9,0)),"Check"))`. It looks up Yes or No for the status, and says Check when the status is not on the list.
5. In **E6**, type `=IF(D6="Yes",IF(AND(ISNUMBER(B6),ISNUMBER($B$2)),IF($B$2>=B6,$B$2-B6,"Check"),"Check"),IF(D6="Check","Check",""))`. This is days open, shown only for open orders.
6. Fill **D6:E6** down to row 205.
7. On **Summary**, type `=IF(A5="","",COUNTIF(Orders!$C$6:$C$205,A5))` in **C5** and fill it down to **C9**.
8. In **F4**, type `=COUNTIF(Orders!D6:D205,"Yes")` for open orders.
9. In **F5**, type `=IF(COUNT(Orders!E6:E205)=0,"",INDEX(Orders!A6:A205,MATCH(MAX(Orders!E6:E205),Orders!E6:E205,0)))`. MAX finds the longest wait, MATCH finds the row, and INDEX returns the order number from that row.
10. In **F6**, type `=IF(COUNT(Orders!E6:E205)=0,"",MAX(Orders!E6:E205))` for the days that order has waited.
11. In **F7**, type `=COUNTIF(Orders!E6:E205,">7")` for open orders older than 7 days.

If two open orders tie for the longest wait, F5 names the one listed first. COUNTIF ignores upper and lower case but not a space after the status.

This counts days since the order date, not since the status last changed.

Help pages: Microsoft's [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34), Google's [function list](https://support.google.com/docs/table/25273), and for the dropdown Microsoft's [data validation page](https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) and Google's [in-cell dropdown page](https://support.google.com/docs/answer/186103). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 18 orders placed from 10/12/2026 to 11/04/2026, as of 11/04/2026. New, Packing and On hold count as open.

| Status | Open? | Orders |
|---|---|---|
| New | Yes | 6 |
| Packing | Yes | 4 |
| On hold | Yes | 1 |
| Shipped | No | 4 |
| Delivered | No | 3 |
| All orders | | 18 |

That makes 11 open orders. The oldest is order 3103, On hold, placed 10/15/2026 and open 20 days. Four open orders are older than 7 days:

| Order | Status | Ordered | Days open |
|---|---|---|---|
| 3103 | On hold | 10/15/2026 | 20 |
| 3105 | Packing | 10/21/2026 | 14 |
| 3107 | Packing | 10/24/2026 | 11 |
| 3109 | Packing | 10/27/2026 | 8 |

The **CHECK** cell in **F11** says OK.

## Troubleshooting

### A status shows Check in the Open? column
The status is not on the Summary list. In the test, a trailing space (`Packing `) showed Check in **D** and **E**, was not counted in **C5:C9**, and the **CHECK** cell said Check. Retype the status or pick it from the list.

### Days open says Check
The order date is after the as-of date, is text, or is blank, or **B2** is not a date. In the test, an order dated 11/05/2026 against an as-of date of 11/04/2026 showed Check, and so did an as-of date typed as text. An order placed on the as-of date shows 0.

### Two orders have the same longest wait
**F5** names only the first one in the list. In the test, a tie returned the order listed first.

### The CHECK cell says Check
It checks that every status is on the list, no row shows Check, no order is missing a status, the counts add up to the number of orders, and **B5:B9** each hold Yes or No. In the test, a duplicate status in the list and Maybe in place of Yes both gave Check.

The template has 200 order slots; more than 200 orders was not run. The status dropdown was not clicked through, so nothing is claimed about it.

## Template

[Download the .xlsx](templates/tidy-tabs-order-status-counts.xlsx)

The file has three tabs. **Orders** has the typed as-of date in **B2**, 200 rows from row 6 (yellow order number, date and status in **A:C**, formulas in **D:E**) and a status dropdown in **C6:C205**. **Summary** has the five statuses and Yes or No flags in **A5:B9**, the counts in **C5:C9**, and a block in **F4:F11**: open orders, the oldest open order, its days open, open orders older than 7 days, row counts and a **CHECK** cell. **How to use** has notes on each formula.

**CHECK** says OK when every status is on the list, no row shows Check and the counts add up to the number of orders. Fix every Check before you read the oldest order.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
