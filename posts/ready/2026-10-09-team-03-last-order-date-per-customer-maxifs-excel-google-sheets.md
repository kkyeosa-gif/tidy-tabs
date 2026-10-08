---
title: How to Find the Last Order Date for Each Customer With MAXIFS in Excel and Google Sheets
labels: excel-formulas, customer-tracking, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. Excel and Google Sheets steps use the same functions; the MAXIFS, SUMIFS and import help links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Find each customer's last order date with =MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2), then pull that day's amount with SUMIFS. Free .xlsx template.
image_prompts: Wholesale delivery counter with a stack of kraft shipping boxes sealed with tape, a clipboard holding paper delivery slips, and a rolled-up spool of twine beside a brass bell, soft window light, all paper text blurred and unreadable, no screens, no logos
image_alt: Kraft shipping boxes stacked beside a clipboard of delivery slips and a brass bell
threads: Which customers haven't ordered in a month?\nOne formula gives each customer's most recent order date: =MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2).\nSubtract it from a report date and you get days since the last order. Needs Excel 2019 or Microsoft 365.
---
Use `=MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2)` to return the most recent order date for the customer in **A2**. Then pull that day's amount with `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2)`.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Summary sheet listing four wholesale customers with last order date, last order amount, days since and order count](images/2026-10-09-team-03-last-order-date-per-customer-maxifs-excel-google-sheets/last-order-date-per-customer-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Keep one row per order on an **Orders** sheet: a real date in column A, the customer name in column B, and the amount in column C. Build the report on a second sheet called **Summary**.

1. On **Orders**, type these headers in A to C: Order date, Customer, Amount.
2. Type each order below the headers. Use real dates such as 09/17/2026, and spell each customer name the same way every time.
3. On **Summary**, type these headers in A to E: Customer, Last order date, Last order amount, Days since, Orders.
4. In **A2:A5**, type each customer name once.
5. In **H1**, type Report date. In **H2**, type the date you are reporting on, for example 10/09/2026.
6. In **B2**, type `=MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2)`. It looks at every row where the customer matches A2 and returns the largest, meaning latest, date.
7. In **C2**, type `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2)`. It adds the amounts for that customer on that exact date.
8. In **D2**, type `=$H$2-B2` for the days since the last order.
9. In **E2**, type `=COUNTIFS(Orders!$B$2:$B$200,A2)` to count that customer's orders.
10. Format **B2:B5** as dates, **C2:C5** as currency with two decimals, and **D2:E5** as numbers. Select **B2:E2** and fill down.

Why two formulas for one answer: MAXIFS returns only the date, not the row it came from. SUMIFS then looks up the amount by matching both the customer and that date.

You can type `=TODAY()` in **H2** for a days-since column that updates itself. The sample uses a typed date so the numbers stay put.

### In Excel

The formulas work as typed. MAXIFS is available in Excel 2019 and Microsoft 365, not in Excel 2016 or earlier. Microsoft documents the arguments on its [MAXIFS function](https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883) and [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) pages. This post did not test Excel.

### In Google Sheets

The same formulas work as typed. Google documents SUMIFS on its [SUMIFS function](https://support.google.com/docs/answer/3238496?hl=en) page. MAXIFS is also a Google Sheets function, but this post did not test it there.

## Example

Sample data below is fictional. It is four wholesale customers and a report date of 10/09/2026.

| Customer | Last order date | Last order amount | Days since (as of 10/09/2026) | Orders |
|---|---|---|---|---|
| Harbor Coffee | 09/17/2026 | $865.00 | 22 | 3 |
| Maple Street Bakery | 09/30/2026 | $428.40 | 9 | 3 |
| Oak & Ember Candles | 09/09/2026 | $510.75 | 30 | 2 |
| Lakeview Florist | 09/24/2026 | $242.00 | 15 | 2 |

Read the Days since column and Oak & Ember Candles stands out at 30 days. That is the customer to email first.

Check one row by hand: 10/09/2026 minus 09/17/2026 is 13 days left in September plus 9 in October, so 22 days for Harbor Coffee. That matches D2.

## Troubleshooting

### Two orders on the same day show one big amount
If a customer places two orders on their last order date, SUMIFS adds both. The cell then shows the day's combined total, not a single order. The template notes this on its **How to use** sheet. If you need only one order, add an order number column and look that up instead.

### A customer with no orders shows 0 or a strange old date
MAXIFS returns 0 when nothing matches. In LibreOffice that is what the cell held. Formatted as a date, a zero shows up as a date from 1899 or 1900, depending on the app. How Excel or Sheets display it was not tested here. Wrap the formula in `IF(E2=0,"No orders",MAXIFS(...))` to show words instead. That wrapper was not tested either.

### The date comes back but an order is missing
Dates stored as text are ignored by MAXIFS. If the latest order is not the one you expect, check that its date aligns to the right like a real date, and retype it if it aligns to the left. `=COUNT(Orders!A2:A200)` against `=COUNTA(Orders!A2:A200)` should match when every date is real.

### The cell shows a number like 46282 instead of a date
MAXIFS returns the date as a serial number, and the cell is formatted as General. Format **B2:B5** as a date.

### #NAME? appears in the cell
Your Excel version may not have MAXIFS. It exists in Excel 2019 and Microsoft 365, not in Excel 2016 or earlier. That comes from Microsoft's help page and was not tested here. In an older version, an array formula with MAX and IF is the usual workaround, but it is not covered or tested in this post.

## Template

[Download the .xlsx](templates/tidy-tabs-last-order-date-maxifs.xlsx)

The file has three sheets. **Orders** holds the sample order rows with dates, customers and amounts. **Summary** holds the four customers with the MAXIFS, SUMIFS, days since and COUNTIFS formulas, plus the report date in **H2**. **How to use** has short fill-in notes, including the same-day order note.

The customers and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, where it stores MAXIFS as `_xlfn.MAXIFS` and calculated it correctly. It was not opened in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
