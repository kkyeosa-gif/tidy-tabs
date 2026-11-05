---
threads_url: https://www.threads.com/@tin_ylab/post/DfXG3S8lqN4
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-count-orders-and-sales-by-state.html
title: How to Count Orders and Sales by State in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 22 workbooks and 1,408 comparisons with 0 mismatches. Checked: the sample, no orders, no states listed, a text total, a blank state, lower and mixed case, a trailing and a leading space, a misspelled state, a tie for the top state by sales and by orders, a state listed twice, gaps in the state list, all zero totals, a negative refund, a half-cent average, one state only, and six random lists of 1 to 300 orders over 50 states. Excel and Google Sheets were not used. Not run: 300 orders (more than the slots), TODAY, a non-US language setting.
search_description: Count orders and total sales by state in Excel and Google Sheets with COUNTIFS and SUMIFS. Free template shows share of sales and the top state.
image_prompts: A stack of three sealed cardboard shipping boxes with blank white address labels, a tape gun and a roll of kraft paper on a packing table in a small home workshop, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Sealed shipping boxes with blank labels beside a tape gun on a packing table
threads: Which state sends you the most orders, and which sends the most money?\nCOUNTIFS counts the orders per state and SUMIFS adds up the sales.\nIn the free template, anything not on your list lands in an All other row, so the table still adds up to the Orders tab.
---
Count the orders for a state with `=COUNTIFS(Orders!$C$5:$C$304,A5)` and add up its sales with `=SUMIFS(Orders!$D$5:$D$304,Orders!$C$5:$C$304,A5)`, where A5 holds the state code such as MA. In the sample, MA has 6 orders totaling $234.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIFS, SUMIFS and INDEX exist in current Excel and Google Sheets, but neither app was opened for this post.

![Spreadsheet table of orders, sales, share and average order for twelve states with a summary](images/2026-11-05-team-01-orders-by-state-countifs-sumifs-excel-google-sheets/orders-and-sales-by-state-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*


## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Orders** has one row per order from row 5: order number in **A**, date in **B**, two-letter ship-to state in **C**, order total in **D**. **States** lists the states you want to see from **A5**.

1. On **Orders**, log each order. Type the state as two letters, such as MA.
2. On **States**, type up to 12 states in **A5:A16**.
3. In **B5**, type `=IF(A5="","",COUNTIFS(Orders!$C$5:$C$304,A5))` for the number of orders.
4. In **C5**, type `=IF(A5="","",SUMIFS(Orders!$D$5:$D$304,Orders!$C$5:$C$304,A5))` for sales.
5. In **D5**, type `=IF(OR(A5="",$C$18=0),"",C5/$C$18)` and format it as a percent. This is the state's share of all sales.
6. In **E5**, type `=IF(OR(A5="",B5=0),"",ROUND(C5/B5,2))` for the average order.
7. Fill **B5:E5** down to row 16.
8. In **B17**, type `=COUNTA(Orders!A5:A304)-SUM(B5:B16)`. In **C17**, type `=ROUND(SUM(Orders!D5:D304)-SUM(C5:C16),2)`. This **All other** row catches every order whose state is not in your list.
9. In **B18** and **C18**, type `=SUM(B5:B17)` and `=SUM(C5:C17)` for the totals.

The dollar signs lock the Orders ranges so they stay put when you fill down. Without the `IF(A5="","",...)` wrapper, an empty state cell would show 0 instead of staying blank.

To name the top state by sales, **H6** is `=IF(MAX(C5:C16)=0,"",INDEX(A5:A16,MATCH(MAX(C5:C16),C5:C16,0)))`. Change both C ranges to B for the top state by number of orders.

This counts and adds up orders by the state you typed. It does not decide whether you owe sales tax in any state. Use that state's own website for that.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 30 orders from 10/01/2026 to 11/03/2026, totaling $1,529.00, an average of $50.97.

| State | Orders | Sales | Share of sales | Average order |
|---|---|---|---|---|
| MA | 6 | $234.00 | 15.3% | $39.00 |
| NY | 6 | $323.50 | 21.2% | $53.92 |
| CA | 5 | $327.50 | 21.4% | $65.50 |
| TX | 3 | $177.00 | 11.6% | $59.00 |
| FL | 3 | $130.50 | 8.5% | $43.50 |
| WA | 2 | $111.50 | 7.3% | $55.75 |
| IL | 2 | $96.00 | 6.3% | $48.00 |
| PA | 1 | $66.00 | 4.3% | $66.00 |
| OH, NC, CO, GA | 0 | $0.00 | 0.0% | blank |
| All other | 2 | $63.00 | 4.1% | $31.50 |

The two All other orders are one from VT and one from OR, which are not in the list. CA is the top state by sales at 21.4%. MA and NY tie at 6 orders, and **H7** names MA because it comes first in your list.

## Troubleshooting

### A state shows fewer orders than expected
The state on Orders may be spelled differently, such as `Mass` or `MAA`. Those orders are not lost. They land in the **All other** row. In the test, a misspelled state moved that order to All other and the table still added up to the Orders tab.

### Upper and lower case do not matter, spaces do
In the test, `ma` and `Ma` were counted with MA. A trailing space (`MA `) or a leading space was a different state and landed in All other. If All other looks too high, check for stray spaces in column C.

### The CHECK cell says Check
**H12** says OK when no row is missing a state, every total is a number, no state is listed twice and the table adds up to the Orders tab. In the test, a blank state, a text total and a state typed twice each changed it to Check. A state listed twice is counted twice, which pushes **All other** below zero.

### The top state is not the one you expected
With a tie, the name shown is the first one in your list. In the test, MA and NY both had 6 orders and MA was named. Sort your list or read the table if the tie matters to you.

### A state with no orders shows a blank average
There is nothing to divide, so **E** stays empty instead of showing an error. In the test, OH, NC, CO and GA showed 0 orders, $0.00 and a blank average.

## Template

[Download the .xlsx](templates/tidy-tabs-orders-by-state.xlsx)

The file has three tabs. **Orders** has 300 rows from row 5 with 30 fictional sample orders and yellow input cells. **States** has 12 state slots in **A5:A16**, an **All other** row, an **All orders** row, and a summary in **H5:H12**: order rows, the top state by sales, the top state by orders, orders in listed states, orders in All other, rows with a missing state, rows where the total is not a number, and a **CHECK** cell. **How to use** has notes on each formula.

Fix every Check before you compare states. The template has 300 order rows. A list longer than that was not tested.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
