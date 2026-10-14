---
threads_url: https://www.threads.com/@tin_ylab/post/DedDSh1lMH5
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-repeat-customer-rate.html
title: How to Calculate Repeat Customer Rate in Excel and Google Sheets
labels: excel-formulas, small-business, countif
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math, including a typo case, an empty Orders tab and a new-customer case. Excel and Google Sheets were not opened; the two help links are from Microsoft Support and Google Docs Help.
search_description: Repeat customer rate: =COUNTIF(B2:B9,">=2")/COUNTIF(B2:B9,">0"). Sample shop: 3 of 7 customers ordered more than once, 42.9%. Free .xlsx.
image_prompts: Packing table with a few kraft shipping boxes, a roll of packing tape and a small paper order pad with a pen, soft window light, no screens, no readable text, no logos
image_alt: Kraft shipping boxes, a tape roll and a paper order pad on a packing table
threads: 3 of 7 customers ordered more than once. That is a 42.9% repeat rate.\nCount orders per customer with =COUNTIF(Orders!$B$2:$B$200,A2), then =COUNTIF(B2:B9,">=2")/COUNTIF(B2:B9,">0").\nOne misspelled name quietly lowers it, so the template has a check cell.
---
Count orders per customer with `=COUNTIF(Orders!$B$2:$B$200,A2)`, then divide repeat customers by customers who ordered with `=COUNTIF(B2:B9,">=2")/COUNTIF(B2:B9,">0")`. In the sample, 3 of 7 customers ordered more than once, a 42.9% repeat rate.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Customer list with order counts, repeat or one-time labels and a summary block showing the repeat customer rate](images/2026-10-13-team-03-repeat-customer-rate-countif-excel-google-sheets/repeat-customer-rate-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: list every order with the customer's name, count how many times each name appears, then compare the number of repeat customers to the number of customers who ordered.

1. On a tab named **Orders**, type these headers in row 1: Order, Customer, Order date, Total.
2. Add one row per order. Put the customer name in column **B**, the date as MM/DD/YYYY in column **C**, and the amount in column **D**.
3. On a second tab named **Customers**, type these headers in row 1: Customer, Orders, Type, Total spent.
4. In **A2:A9**, type each customer once.
5. In **B2**, type `=IF(A2="","",COUNTIF(Orders!$B$2:$B$200,A2))` and fill it down. This counts that customer's orders.
6. In **C2**, type `=IF(A2="","",IF(B2=0,"No orders",IF(B2=1,"One time","Repeat")))` and fill it down.
7. In **D2**, type `=IF(A2="","",SUMIF(Orders!$B$2:$B$200,A2,Orders!$D$2:$D$200))` and fill it down. This adds up what each customer spent.
8. In **G2**, type `=COUNTIF(B2:B31,">0")`. This is customers who ordered at all.
9. In **G3**, type `=COUNTIF(B2:B31,">=2")`. This is repeat customers.
10. In **G4**, type `=IF(G2=0,"",G3/G2)` and format it as a percent with one decimal.

Customers with 0 orders are left out of the bottom number on purpose. A name on your list who never ordered is not a lost repeat buyer.

Add a safety check. In **G5**, type `=SUM(B2:B31)`. In **G6**, type `=COUNTA(Orders!B2:B200)`. In **G7**, type `=IF(G5=G6,"OK","CHECK")`. If G7 says CHECK, some orders did not match a name on your list.

### In Excel

The formulas work as typed. Microsoft describes the function on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) page. Format **G4** as a percent from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [COUNTIF help page](https://support.google.com/docs/answer/3093480). Format **G4** with **Format > Number > Percent**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is 13 orders from a small online shop, 08/04/2026 to 10/09/2026.

| Customer | Orders | Type | Total spent |
|---|---|---|---|
| Ana Ruiz | 3 | Repeat | $94.50 |
| Ben Ortiz | 1 | One time | $65.50 |
| Cara Lee | 2 | Repeat | $72.00 |
| Dev Patel | 1 | One time | $54.25 |
| Eli Brooks | 4 | Repeat | $113.75 |
| Fay Nolan | 1 | One time | $12.00 |
| Gus Hall | 0 | No orders | $0.00 |
| Hana Cole | 1 | One time | $90.00 |

| Result | Cell | Value |
|---|---|---|
| Customers with an order | G2 | 7 |
| Repeat customers | G3 | 3 |
| Repeat customer rate | G4 | 42.9% |
| Orders counted / on Orders tab | G5 / G6 | 13 / 13 |
| Check | G7 | OK |
| Spent by repeat customers | G8 | $280.25 |
| Share of spending from repeat customers | G9 | 55.8% |

Gus Hall is on the list but has no orders, so he is not counted in the 7. The exact rate is 3 / 7 = 42.857%, shown as 42.9%.

The three repeat customers made 55.8% of the money ($280.25 of $502.00 total). G8 is `=SUMIF(C2:C31,"Repeat",D2:D31)` and G9 is `=IF(SUM(D2:D31)=0,"",G8/SUM(D2:D31))`.

## Troubleshooting

### The rate is higher than it should be and G7 says CHECK
A name on the Orders tab does not match the Customers tab. In the recalculated file, "Ana Ruis" in **B2** and "Ben Ortiz " with a trailing space in **B4** dropped Ana to 2 orders and Ben to 0. G2 fell to 6, the rate jumped to 50.0%, and G5 showed 11 against 13 in G6. Fix the spelling or delete the extra space.

### Capital letters do not matter
"eli brooks" in **B3** still counted for Eli Brooks (4 orders). COUNTIF ignores capitals but not spaces or different spellings.

### The rate is blank
No customer has an order, so G2 is 0 and the formula returns an empty cell on purpose. In the recalculated file, clearing the Orders tab gave 0 in G2 and G3, blank G4 and G9, and OK in G7.

### A new customer is not counted
The Customers list is typed by hand. Type the new name in the next empty cell in column A. In the recalculated file, adding a name in **A10** with one order gave 1 in B10, "One time" in C10, 8 in G2, and OK in G7.

### The rate changes when I paste more orders
It covers only the orders on the Orders tab. Use the same date range each time you compare rates. This sheet also says what share ordered again, not why.

## Template

[Download the .xlsx](templates/tidy-tabs-repeat-customers.xlsx)

The file has three tabs. **Orders** holds the 13 sample orders in A2:D14. **Customers** holds the 8 sample names in A2:A9, formulas ready for rows 2 to 31, and the summary in F1:G9. **How to use** has short fill-in notes. Formulas read Orders rows 2 to 200.

The names, orders and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
