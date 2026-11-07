---
threads_url: https://www.threads.com/@tin_ylab/post/DfcQcXOm0PV
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-count-new-and-returning.html
title: How to Count New and Returning Customers Each Month in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 26 workbooks and 25,896 comparisons with 0 mismatches. Checked: the sample, no orders, one order, three orders from one customer on one day, an unsorted list, capitals in names, a trailing and a leading space in a name, orders on the first and last day of a month, orders in other years, the years 2027 and 2028 (leap day), a text date, a date with a time, a blank customer, a blank date, a text amount, a refund, one customer all year, all new customers, 300 orders, and six random lists of up to 300 orders and 60 customers. Excel and Google Sheets were not used. Not run: wildcard characters (* ? ~) in names, more than 300 orders, a non-US language setting.
search_description: Count new and returning customers by month in Excel and Google Sheets with MINIFS, COUNTIFS and SUMIFS. Free template splits orders and sales.
image_prompts: A small stack of blank kraft thank-you cards tied with twine, a wrapped bar of soap and a pencil on a wooden craft table, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Kraft thank-you cards tied with twine beside a wrapped soap bar and a pencil
threads: In the sample, 8 of October's 11 orders came from returning customers.\nMINIFS finds each customer's earliest order date, and an order is New when its date matches it.\nThe free template splits orders and sales by month. MINIFS needs Excel 2019 or newer, or Google Sheets.
---
Label an order New when its date equals the customer's earliest date, with `=IF(B6=E6,"New","Returning")`, where E6 is `=MINIFS($B$6:$B$305,$C$6:$C$305,C6)`. In the sample, October has 3 new customers and 8 of its 11 orders come from returning customers.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model (26 workbooks, 25,896 comparisons). MINIFS, COUNTIFS, SUMIFS and EDATE exist in Microsoft 365, Excel 2019 and Google Sheets, but neither Excel nor Google Sheets was opened for this post. MINIFS is not in Excel 2016 or earlier.

![Monthly table splitting orders and sales into new and returning customers for 2026](images/2026-11-06-team-05-new-vs-returning-customers-by-month-minifs-countifs-excel-google-sheets/new-vs-returning-customers-by-month-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. **Orders** has one row per order from row 6: order number in **A**, order date in **B** (dates only, no times), customer in **C**, amount in **D**. **Months** shows one row per month of the year typed in **B2**.

1. On **Orders**, log every order, one row each. Type the customer the same way every time.
2. In **E6**, type `=IF(OR(B6="",C6=""),"",IF(AND(ISNUMBER(B6),B6=INT(B6)),MINIFS($B$6:$B$305,$C$6:$C$305,C6),"Check"))`. This finds the customer's earliest order date and flags a date that has a time or is text.
3. In **F6**, type `=IF(E6="","",IF(E6="Check","Check",IF(B6=E6,"New","Returning")))`.
4. In **G6**, type `=IF(F6="New",IF(COUNTIFS($C$6:C6,C6,$B$6:B6,B6)=1,1,0),IF(F6="Returning",0,""))`. A customer who orders twice on their first day gets two New rows but only one 1 here, so the customer is counted once.
5. Fill **E6:G6** down to row 305.
6. On **Months**, type the year in **B2**, such as 2026. In **A5**, type `=DATE($B$2,1,1)`. In **A6** to **A16**, use `=DATE($B$2,2,1)` through `=DATE($B$2,12,1)`, and format column A as **mmmm**.
7. In **B5**, type `=COUNTIFS(Orders!$B$6:$B$305,">="&$A5,Orders!$B$6:$B$305,"<"&EDATE($A5,1))` for the orders in the month.
8. In **C5**, type `=SUMIFS(Orders!$G$6:$G$305,Orders!$B$6:$B$305,">="&$A5,Orders!$B$6:$B$305,"<"&EDATE($A5,1))` for new customers.
9. In **D5**, use the **B5** formula with `,Orders!$F$6:$F$305,"New"` before the last parenthesis. In **E5**, do the same with `"Returning"`.
10. In **F5**, type `=SUMIFS(Orders!$D$6:$D$305,Orders!$B$6:$B$305,">="&$A5,Orders!$B$6:$B$305,"<"&EDATE($A5,1),Orders!$F$6:$F$305,"New")`. In **G5**, use `"Returning"` instead.
11. In **H5**, type `=IF(B5=0,"",E5/B5)` and format it as 0.0%. Fill **B5:H5** down to row 16, then add `=SUM(B5:B16)` style totals in row 17.

The `>=` and `<` pair with `EDATE($A5,1)` picks every date from the first of the month up to, but not including, the first of the next month. Without the `IF(B5=0,"",...)` wrapper, an empty month would show a division error.

"New" means new to this list. Someone who ordered before your first row looks new in your first month. This sheet does not know about orders you did not log.

Help pages: Microsoft's [MINIFS function](https://support.microsoft.com/en-us/office/minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599) and [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 30 orders (1001 to 1030) from 12 customers, 08/03/2026 to 10/31/2026.

| Month | Orders | New customers | Orders on first day | Orders from returning | Sales on first day | Sales from returning | Returning share |
|---|---|---|---|---|---|---|---|
| August | 9 | 5 | 6 | 3 | $220.25 | $80.75 | 33.3% |
| September | 10 | 4 | 4 | 6 | $225.75 | $174.74 | 60.0% |
| October | 11 | 3 | 3 | 8 | $158.00 | $243.75 | 72.7% |
| Year total | 30 | 12 | 13 | 17 | $604.00 | $499.24 | 56.7% |

Order 1001 (Maria Alvarez, 08/03/2026, $42.00) is New and has a 1 in **G**. Her order 1004 on 08/12/2026 is Returning. Orders 1005 and 1006 (Tom Nguyen) are both on 08/14/2026, so both are New, but only 1005 has the 1. That is why August shows 6 first-day orders and 5 new customers.

The other months show 0 and a blank share.

## Troubleshooting

### Column E shows an error in an older Excel
MINIFS is in Excel 2019, Microsoft 365 and Google Sheets, not in Excel 2016 or earlier, so an older Excel cannot run column E. We did not open Excel to see what it shows. The template's **How to use** tab says the same.

### A row says Check
The date has a time of day or is text. Retype it as a plain date such as 11/03/2026. Fix every Check row first. In the test, a date with a time can still change a customer's earliest date, so other rows may be wrong until it is fixed.

### One customer is counted as two
A space before or after a name makes a different customer. In the test, a name with a trailing space and the same name without it were two customers, and so were a leading space and none. Capitals do not matter: the same name in a different case matched.

### The year total is below the number of orders
**Months** only counts orders in the year in **B2**. **K5** counts orders in other years, and **K8** says Check when it is above zero. Change **B2** to the year you want, or remove the other-year orders.

### A returning customer shows as New
A customer shows as New when the earlier order was typed under a slightly different name. The sheet compares dates for the same name only.

## Template

[Download the .xlsx](templates/tidy-tabs-new-vs-returning-customers.xlsx)

The file has three tabs. **Orders** has 300 rows from row 6 with 30 fictional sample orders and yellow input cells. **Months** has the year in **B2**, 12 month rows, a **Year total** row, and a summary in **K4:K8**: order rows, orders in other years, rows showing Check, customers seen, and a **CHECK** cell. **How to use** has notes on each formula.

Fix every Check before you read the table. The template has 300 order rows. A longer list was not tested.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
