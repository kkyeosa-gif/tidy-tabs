---
threads_url: https://www.threads.com/@tin_ylab/post/DfIZVx3mTR3
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-flag-customers-who-have-not.html
title: How to Flag Customers Who Have Not Ordered in 90 Days in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model of the last order date and day count, in 27 workbooks (1,932 cell comparisons, 0 mismatches after two fixes). Checked: an order exactly 90 days back, 91 days back, a moved as-of date, cutoffs from 0 to 1000, orders dated after the as-of date, a customer with no orders, capital letters, a trailing space, a duplicate name, an unlisted name, text dates and invalid settings. MAXIFS was checked as the _xlfn.MAXIFS form stored in the .xlsx file. TODAY() in the as-of cell was not run. Excel and Google Sheets were not used.
search_description: Flag customers who have not ordered in 90 days with MAXIFS in Excel and Google Sheets. Free template with a typed as-of date, your cutoff and lapsed spend.
image_prompts: A small wooden box of paper customer index cards with one card pulled halfway out beside a desk calendar with one date circled in pencil, soft daylight, no screens, no readable text, no logos
image_alt: Wooden box of index cards with one pulled out beside a desk calendar with a circled date
threads: Which customers have gone quiet for 90 days?\nMAXIFS finds each one's last order, then one IF compares the days since against a cutoff you type in a cell.\nThe sheet also totals what the lapsed customers spent, so the win-back list has a size.
---
Find each customer's last order with `=MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A5)`, then flag Lapsed when the days since that order are above the cutoff in a cell, as in `=IF(C5>Settings!$B$2,"Lapsed","Active")`. Here C5 is the days since the last order and Settings B2 is 90.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every status and number below was recalculated and compared with a Python model. MAXIFS needs Excel 2019 or newer and is listed in Google Sheets' function list. Neither Excel nor Google Sheets was opened for this post, and no older-Excel alternative was built.

![Fifteen customers with last order date, days since, Active or Lapsed status and spend](images/2026-10-30-team-04-flag-customers-no-order-in-90-days-maxifs-excel-google-sheets/lapsed-customer-flag-template.png) *caption: the Customers sheet of the template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

The layout matches the template. **Orders** has the order date in column A, the customer name in B and the amount in C, starting in row 2. **Customers** has one name per row in column A, starting in row 5. **Settings** holds the as-of date in **B1** and the days without an order in **B2**.

Type the as-of date instead of using `TODAY()`. A typed date keeps the results fixed, so the list you made on 10/30/2026 still says the same thing next week. Swapping in `=TODAY()` makes every status move forward each day. That swap was not run in the tests, so check it yourself before relying on it.

The steps are the same in Excel and Google Sheets.

1. On **Settings**, type 10/30/2026 in **B1** and 90 in **B2**.
2. On **Orders**, enter or paste your orders in columns A to C. Dates must be real dates, not text.
3. On **Customers**, list each customer once in column A, starting at **A5**.
4. In **B5**, type the last-order formula below.
5. In **C5**, type `=IF(AND(ISNUMBER(B5),ISNUMBER(Settings!$B$1)),Settings!$B$1-B5,"")`. This is the days since the last order.
6. In **D5**, type the status formula below.
7. In **E5**, type `=IF(A5="","",SUMIF(Orders!$B$2:$B$200,A5,Orders!$C$2:$C$200))` for lifetime spend.
8. Select **B5:E5** and drag the fill handle down to your last customer.
9. Format column B as a date and column E as currency.

The last-order formula in **B5**:

`=IF(A5="","",IF(COUNTIFS(Orders!$B$2:$B$200,A5,Orders!$A$2:$A$200,">0")=0,"No orders",MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A5)))`

MAXIFS returns the latest date among the rows where the name matches. The COUNTIFS part catches a customer with no dated orders and shows "No orders" instead of a date of 12/30/1899. The test found that bug with a plain COUNTIF, so the template uses COUNTIFS with `">0"`.

The status formula in **D5**:

`=IF(A5="","",IF(Settings!$B$3<>"OK","Check settings",IF(NOT(ISNUMBER(C5)),"No orders",IF(C5>Settings!$B$2,"Lapsed","Active"))))`

The test is "greater than", so a customer exactly 90 days out stays Active and 91 days is Lapsed. Change **B2** to 60 or 120 and every row updates.

Help pages: Microsoft's [MAXIFS function](https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883) and [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842), and Google's [function list](https://support.google.com/docs/table/25273), which includes MAXIFS. Those pages describe the functions, not our test.

## Example

Sample data is fictional: 15 made up customers and 40 orders, as of 10/30/2026 with a 90 day cutoff.

| Customer | Last order | Days since | Status | Lifetime spend |
|---|---|---|---|---|
| Alma Reyes | 10/27/2026 | 3 | Active | $317.80 |
| Hugo Park | 08/02/2026 | 89 | Active | $459.81 |
| Iris Wong | 08/01/2026 | 90 | Active | $250.23 |
| Jonah Bell | 07/31/2026 | 91 | Lapsed | $272.20 |
| Kira Stone | 07/30/2026 | 92 | Lapsed | $263.52 |
| Nate Dunn | 12/24/2025 | 310 | Lapsed | $226.43 |
| Opal Frost | No orders | | No orders | $0.00 |

The summary beside the list shows 15 customers: 9 Active, 5 Lapsed and 1 with no orders. The lapsed customers spent $1,387.91 in total, which gives your win-back list a size. The check cell says OK.

Move the as-of date to 10/31/2026 and Iris Wong becomes 91 days and Lapsed, so the count goes to 8 Active and 6 Lapsed. Lapsed rows are red and No orders rows are gray.

This labels customers by date. It says nothing about whether or how to contact anyone.

## Troubleshooting

### A customer shows No orders but has orders
The name does not match exactly. In the tests, a trailing space in the Customers list ("Alma Reyes ") found nothing. Capital letters do not matter: "BEN TRAN" matched "Ben Tran". Retype the name or remove the space. The summary cell for orders with an unlisted name counts the orders that lost their match, and the check cell shows CHECK.

### Every row says Check settings
The as-of date is text, or the days cell is not a whole number of 0 or more. **Settings B3** shows the reason. Retype the date as a real date and the days as a number such as 90.

### An order does not count toward the last order date
Its date is text, such as "March 3" or "2026-09-01" typed as text. MAXIFS ignores it. Its amount still counts in lifetime spend, so a customer whose only orders have text dates shows No orders next to a spend above $0.00. Retype the dates as real dates.

### A customer with a recent order shows a negative number of days
The order is dated after the as-of date. That customer stays Active, and the summary counts these orders in its own cell. Move the as-of date forward, or fix the order date.

### The check cell says CHECK
It shows CHECK for bad settings, a name listed twice, an order for a name that is not in the Customers list, or data past the order rows or customer slots. The summary cells beside it point to which one.

## Template

[Download the .xlsx](templates/tidy-tabs-lapsed-customer-flag.xlsx)

The file has four tabs. **Orders** holds 40 fictional orders, with yellow inputs in **A2:C200**. **Customers** lists 15 fictional names in **A5:A54** (50 slots) with last order, days since, status and lifetime spend in **B** to **E**, plus a summary in **G** and **H**. **Settings** has the yellow as-of date and cutoff and a settings check. **How to use** has short notes, including the MAXIFS version limit.

Limits: it holds 199 orders and 50 customers. More than that was not tested. Orders with a time part on the date were not tested either. The file was recalculated only in LibreOffice Calc 24.2.7.2. The Customers tab prints landscape and Orders prints portrait, both on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
