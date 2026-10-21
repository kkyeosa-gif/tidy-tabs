---
title: How to Calculate the Average Days Between Orders for Each Customer in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python date model, including 16 edited variants for order gaps (one order, same day orders, uneven gaps, shuffled rows, text dates, trailing spaces, a customer missing from Orders, 20 customers with 299 orders). No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Calculate the average days between a customer's orders with MINIFS and MAXIFS in Excel and Google Sheets. Free template with one row per customer.
image_prompts: A small metal card file box with blank index cards fanned out, a paper desk calendar with faint unreadable markings and a spool of ribbon on a gift shop counter, soft daylight, no screens, no readable text, no logos
image_alt: Index card file box, paper desk calendar and ribbon spool on a gift shop counter
threads: Three orders have two gaps between them, not three.\nSo average days between orders is (last date - first date)/(orders - 1).\nThe sample customer ordered on 01/05, 02/02 and 04/13/2026: 98 days, 2 gaps, 49.0 days.
---
Find a customer's average days between orders with `=IF(B2<2,"One order",(D2-C2)/(B2-1))`, where B2 is the number of orders, C2 the first order date and D2 the last. It divides the span by the number of gaps, which is one less than the number of orders.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. MAXIFS and MINIFS need Excel 2019 or later. Neither Excel nor Google Sheets was opened for this post.

![Customers sheet with order counts, first and last order dates and average days between orders](images/2026-10-22-team-04-average-days-between-orders-per-customer-maxifs-minifs-excel-google-sheets/days-between-orders-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

Put every order on one tab with one row per order, and give each customer one row on a second tab. The template uses an **Orders** tab and a **Customers** tab.

1. On **Orders**, type the customer name in column A and the order date in column B, starting in row 2. The template reads down to row 300.
2. On **Customers**, type each customer name once in column A, starting in A2. Spell it exactly as on Orders.
3. In **B2**, count the orders: `=COUNTIF(Orders!$A$2:$A$300,A2)`.
4. In **C2**, find the first order date: `=MINIFS(Orders!$B$2:$B$300,Orders!$A$2:$A$300,A2)`.
5. In **D2**, find the last order date: `=MAXIFS(Orders!$B$2:$B$300,Orders!$A$2:$A$300,A2)`.
6. In **E2**, type `=IF(B2<2,"One order",(D2-C2)/(B2-1))`.
7. Format C2 and D2 as MM/DD/YYYY and E2 as a number with 1 decimal. Fill the formulas down.

Why minus 1: three orders have two gaps between them, not three. Dividing by the number of orders would make every average too small.

The template's own cells wrap these formulas in extra tests for blank rows, a name that is not on Orders and dates stored as text, so they are longer than the ones above. The core is the same.

### In Excel

MAXIFS and MINIFS need Excel 2019 or later. Microsoft describes them on its [MAXIFS function](https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883) and [MINIFS function](https://support.microsoft.com/en-us/office/minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599) pages. The template stores them as `_xlfn.MINIFS` and `_xlfn.MAXIFS`, which is how Excel files record them.

### In Google Sheets

Google Sheets has MINIFS and MAXIFS, listed on Google's [function list](https://support.google.com/docs/table/25273). The same formulas should work as typed, but Sheets was not opened for this post.

## Example

Sample data is fictional: wholesale and subscription box customers of a small shop, 40 orders from 01/05/2026 to 09/22/2026.

| Customer | Orders | First order | Last order | Average days between |
|---|---|---|---|---|
| Ann Lopez | 3 | 01/05/2026 | 04/13/2026 | 49.0 |
| Maple Street Gift Shop | 6 | | | 28.0 |
| Harbor Coffee Cart | 5 | | | 61.3 |
| Dana Whitfield | 1 | | | One order |
| Noah Chen | 2 | | | 119.0 |

Ann Lopez ordered on 01/05, 02/02 and 04/13/2026. That is 98 days from first to last, divided by 2 gaps, so 49.0 days. The sheet computes this from the dates; it does not use the gaps one by one.

The Reorder window column, **F2**, is the last order date plus the average: `=IF(ISNUMBER(E2),D2+E2,"")`. For Ann Lopez it shows 06/01/2026. It is arithmetic on past orders, not a prediction, and it does not use today's date.

The summary cells show 40 orders on Orders, 40 counted across the customers and 9 customers with two or more orders. **I5** says OK when every order row belongs to a customer on the list and every date is a real date.

### What the average hides

The average is the span divided by the gaps, so it cannot see uneven gaps. A customer with orders on 01/01, 01/11, 01/21 and 04/11/2026 shows 33.3 days, although two of the three gaps are 10 days. Look at the dates before you rely on the average.

## Troubleshooting

### One customer's orders are undercounted and I5 says CHECK
A name on Orders probably has a trailing space, such as "Harbor Coffee Cart ". It counts as a different customer. In the template's test, Harbor dropped from 5 orders to 4 and from 61.3 days to 67.7. Retype the name or use **Data > Text to Columns** to clean the column.

### The cell shows One order instead of a number
With one order there are no gaps to average, so the formula returns "One order" rather than dividing by 0. The Reorder window stays blank for that customer.

### The average is 0.0
The customer has only same day orders. Two orders on 05/19/2026 give a span of 0 days, so the average is 0.0 with no error.

### The cell says Check dates
At least one order date for that customer is text, such as "02/09/2026" typed as text, or is blank. Retype the date or use **Data > Text to Columns** with the Date format. I5 turns to CHECK too.

### The cell says Not in Orders
The name on Customers does not match any name on Orders. Check the spelling and spaces. Upper and lower case do not matter.

### I have Excel 2016 or earlier
MINIFS and MAXIFS are not there. An array formula with MIN and IF finds the first date: `=MIN(IF(Orders!$A$2:$A$300=A2,Orders!$B$2:$B$300))`, and the same with MAX finds the last. Older Excel needs Ctrl+Shift+Enter for these, according to Microsoft's help pages. In LibreOffice Calc the two array formulas gave the same first and last date as MINIFS and MAXIFS on a four row probe (01/05/2026 and 04/13/2026). They are only in the template's How to use text, not in a cell, and were not run in Excel.

## Template

[Download the .xlsx](templates/tidy-tabs-days-between-orders.xlsx)

The file has three tabs. **Orders** holds the yellow inputs in **A2:B300**, with 40 fictional orders in rows 2 to 41. **Customers** lists up to 20 customers in **A2:A21** and calculates orders, first order, last order, average days and the reorder window in **B:F**, with a summary in **H1:I5**. **How to use** has short notes. All tabs print portrait on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
