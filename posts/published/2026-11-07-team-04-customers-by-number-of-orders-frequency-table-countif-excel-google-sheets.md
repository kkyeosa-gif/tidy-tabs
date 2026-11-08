---
threads_url: https://www.threads.com/@tin_ylab/post/DfdiCbcCRD1
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-count-customers-by-number-of.html
title: How to Count Customers by Number of Orders in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 19 workbooks and 17,765 comparisons with 0 mismatches, and no error value appeared in any recalculated workbook. Excel and Google Sheets were not used. Not run: more than 300 orders, a non-US language setting.
search_description: Count customers by number of orders in Excel and Google Sheets with SUMPRODUCT and COUNTIFS. Free template shows customers, orders and sales per group.
image_prompts: A stack of small cardboard shipping boxes sealed with plain tape, a roll of twine and a rubber stamp on a pale wooden packing bench, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Stacked cardboard shipping boxes with a roll of twine and a rubber stamp on a bench
threads: In the sample, 4 of 14 customers ordered exactly twice.\nOne formula gives every row its customer's order count, another marks each customer's first row so nobody is counted twice.\nThen COUNTIFS fills the 1, 2, 3, 4, 5 or more table. The template is free.
---
Give each customer's first row a 1 with `=IF(SUMPRODUCT(--($C$5:C5=C5))=1,1,0)`, count their orders with `=SUMPRODUCT(--($C$5:$C$304=C5))`, then count customers who ordered exactly twice with `=COUNTIFS($E$5:$E$304,1,$D$5:$D$304,2)`. In the sample, 4 of 14 customers ordered twice.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model (19 workbooks, 17,765 comparisons). SUMPRODUCT and COUNTIFS exist in Excel and Google Sheets, but neither Excel nor Google Sheets was opened for this post.

![Table of customers, orders and sales grouped by number of orders placed](images/2026-11-07-team-04-customers-by-number-of-orders-frequency-table-countif-excel-google-sheets/customers-by-number-of-orders-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. **Orders** has one row per order from row 5: customer in **A**, order amount in **B**. The list holds 300 orders (rows 5 to 304).

1. On **Orders**, log every order, one row each. Type each customer the same way every time.
2. In **C5**, type `=IF(TRIM(A5)="","",TRIM(A5))`. This removes extra spaces at the ends of the name.
3. In **D5**, type `=IF(C5="","",SUMPRODUCT(--($C$5:$C$304=C5)))`. It counts how many rows have this customer's name. Capitals do not matter.
4. In **E5**, type `=IF(C5="","",IF(SUMPRODUCT(--($C$5:C5=C5))=1,1,0))`. The range grows as you fill down, so only the first row of each customer gets a 1.
5. Fill **C5:E5** down to row 304.
6. In the table, type the group sizes 1, 2, 3 and 4 in **G6:G9** and "5 or more" in **G10**. In **H6**, type `=COUNTIFS($E$5:$E$304,1,$D$5:$D$304,2)` for customers who ordered exactly twice. Change the 2 for each row.
7. For "5 or more", use `=COUNTIFS($E$5:$E$304,1,$D$5:$D$304,">=5")`.
8. For orders in a group, use `=COUNTIFS($D$5:$D$304,2)`. For sales, use `=SUMIFS($B$5:$B$304,$D$5:$D$304,2)`.
9. For shares, use `=IF($H$10=0,"",H6/$H$10)` for customers and `=IF($K$10=0,"",K6/$K$10)` for sales, formatted as 0.0%.

The 1 in column **E** is what keeps a customer who ordered three times from being counted three times. COUNTIFS counts only rows where **E** is 1 and **D** matches the group.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34). They describe the functions, not our test.

## Example

Sample data is fictional: 40 orders from 14 customers.

| Orders placed | Customers | Share of customers | Orders | Sales | Share of sales |
|---|---|---|---|---|---|
| 1 | 3 | 21.4% | 3 | $130.00 | 9.2% |
| 2 | 4 | 28.6% | 8 | $200.00 | 14.1% |
| 3 | 3 | 21.4% | 9 | $355.75 | 25.1% |
| 4 | 2 | 14.3% | 8 | $303.50 | 21.4% |
| 5 or more | 2 | 14.3% | 12 | $429.25 | 30.3% |
| Total | 14 | | 40 | $1,418.50 | |

The summary shows 2.86 orders per customer on average. The most orders from one customer is 7, from Maria Lopez. The 5 or more group is 14.3% of customers and 30.3% of sales.

## Troubleshooting

### One person is counted as two customers
Two spellings, such as Jo Ann and JoAnn, are two customers. The formulas ignore capitals and trim spaces at the ends, so Maria Lopez and maria lopez with a trailing space are one customer. Clean the names first, then count.

### CHECK shows Check
A row has an amount but no customer name, or a name but an amount that is not a number. Look at **H16** and **H17** in the summary to see which. Fix those rows, then the total matches.

### The counts stop at row 304
The formulas cover 300 orders. Orders below row 304 are not counted. A longer list was not tested, and each row scans the whole list, so a much longer one may slow down.

### The table counts orders from all time
There is no date column. The table counts every order listed. To count one year, list only that year's orders.

### A name with * or ? gives odd counts
The comparison in `SUMPRODUCT` uses = , so * and ? in a name are plain characters, not wildcards. This was tested in LibreOffice only.

## Template

[Download the .xlsx](templates/tidy-tabs-customers-by-number-of-orders.xlsx)

The file has two tabs. **Orders** has 300 rows from row 5 with 40 fictional sample orders and yellow input cells, the frequency table in **G4:L10**, and a summary in **G12:H18** with a **CHECK** cell. **How to use** has notes on each formula.

Fix every Check before you read the table. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
