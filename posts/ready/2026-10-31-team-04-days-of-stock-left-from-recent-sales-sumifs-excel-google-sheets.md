---
title: How to Calculate Days of Stock Left From Recent Sales in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model in 39 workbooks (8,268 cell comparisons, 0 mismatches). Checked: a sale exactly 30 days before the as-of date (left out) and 29 days before (counted), a sale on the as-of date and one after it, windows of 1, 7, 60 and 90 days, invalid settings, a product with zero stock and one with no sales, a SKU missing from Products, capital letters and a trailing space in a SKU, a text date and a text quantity in Sales, blank, text and negative on-hand cells, a repeated SKU, a full 40 products and 200 sales rows, and random data. TODAY() in the as-of cell was not run. Excel and Google Sheets were not used.
search_description: Calculate days of stock left in Excel and Google Sheets with SUMIFS over the last 30 days of sales. Free template with a run-out date and Order soon flag.
image_prompts: A shelf of folded cotton t-shirts and enamel mugs next to a small stack of cardboard shipping boxes with one shelf space empty, soft daylight, no screens, no readable text, no logos
image_alt: Shelf of folded t-shirts and enamel mugs with an empty gap beside stacked shipping boxes
threads: Divide what is on the shelf by what sold per day and you know how long it lasts.\nSUMIFS pulls just the last 30 days from your sales log, so a busy September does not fool the math.\nThe sheet adds a run-out date and flags anything under 21 days.
---
Divide the units on hand by the average units sold per day, using `=INT(C5/E5)` where E5 is the last 30 days of sales divided by 30. Pull the 30-day total from your sales log with `=SUMIFS(Sales!$C$5:$C$204,Sales!$B$5:$B$204,A5,Sales!$A$5:$A$204,">"&Settings!$B$1-Settings!$B$2,Sales!$A$5:$A$204,"<="&Settings!$B$1)`.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where every number below was recalculated and compared with a Python model. SUMIFS, INT and ROUND exist in Excel and Google Sheets, but neither app was opened for this post.

![Twelve products with on hand, sold in the last 30 days, average per day, days left, run-out date and status](images/2026-10-31-team-04-days-of-stock-left-from-recent-sales-sumifs-excel-google-sheets/days-of-stock-left-template.png) *PDF render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The layout matches the template, which adds blank-cell guards to these formulas. **Sales** has the date, SKU and units in columns A to C from row 5. **Products** has the SKU, name and units on hand in A to C from row 5. **Settings** holds the as-of date in **B1**, the look-back days in **B2** and the warning level in **B3**.

Type the as-of date instead of using `TODAY()`, so the numbers stay fixed until you change it. `TODAY()` was not run in the tests. The steps are the same in Excel and Google Sheets.

1. On **Settings**, type 10/31/2026 in **B1**, 30 in **B2** and 21 in **B3**.
2. On **Sales**, enter or paste one row per sale: a real date, the SKU and the units sold.
3. On **Products**, list each SKU once with its name and units on hand.
4. In **D5**, type the sold-in-window formula below.
5. In **E5**, type `=ROUND(D5/Settings!$B$2,2)`. This is the average units per day.
6. In **F5**, type `=IF(E5=0,"No recent sales",INT(C5/E5))`. This is the days left, rounded down.
7. In **G5**, type `=IF(ISNUMBER(F5),Settings!$B$1+F5,"")` and format the column as a date. This is the run-out date.
8. In **H5**, type `=IF(C5=0,"Out of stock",IF(F5="No recent sales","No recent sales",IF(F5<Settings!$B$3,"Order soon","OK")))`.
9. Fill **D5:H5** down to your last product.

The sold-in-window formula in **D5**:

`=SUMIFS(Sales!$C$5:$C$204,Sales!$B$5:$B$204,A5,Sales!$A$5:$A$204,">"&Settings!$B$1-Settings!$B$2,Sales!$A$5:$A$204,"<="&Settings!$B$1)`

It adds the units for one SKU dated after the as-of date minus the look-back days, up to and including the as-of date. With 30 days and 10/31/2026, the window runs 10/02/2026 to 10/31/2026. A sale dated 10/01/2026 is left out and one dated 10/02/2026 counts. Change **B2** to 7 or 60 and every row updates.

Days left is `INT(C5/E5)`, so 22.5 days shows as 22. It divides by the rounded average in column E, the number you see on the sheet. The template accepts 1 to 90 look-back days so that rounding stays small.

The status test is "below": exactly 21 days left stays OK and 20 is Order soon.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), and Google's [function list](https://support.google.com/docs/table/25273). Those pages describe the functions, not our test.

## Example

Sample data is fictional: 12 made up products and 60 sales rows, as of 10/31/2026 with a 30 day window and a warning level of 21 days.

| Product | On hand | Sold in window | Per day | Days left | Runs out about | Status |
|---|---|---|---|---|---|---|
| Logo tee, small | 16 | 24 | 0.80 | 20 | 11/20/2026 | Order soon |
| Logo tee, large | 0 | 12 | 0.40 | 0 | 10/31/2026 | Out of stock |
| Enamel mug | 24 | 45 | 1.50 | 16 | 11/16/2026 | Order soon |
| Sticker pack | 240 | 90 | 3.00 | 80 | 01/19/2027 | OK |
| Enamel pin | 42 | 60 | 2.00 | 21 | 11/21/2026 | OK |
| Soy candle, 8 oz | 12 | 0 | 0.00 | No recent sales | | No recent sales |

The summary beside the list shows 12 products: 4 Order soon, 1 Out of stock and 1 with No recent sales. The check cell says OK.

The soy candle sold 6 units before 10/02/2026 and none inside the window, so it shows No recent sales.

This is a simple average of recent sales. It ignores seasons, bursts and how long a restock takes. For lead time, see [How to Calculate a Reorder Point in Excel and Google Sheets](https://tidytabs.blogspot.com/2026/10/how-to-calculate-reorder-point-in-excel.html). The run-out date is arithmetic on your own numbers, not a forecast.

## Troubleshooting

### A product shows No recent sales but it has sales
The SKU in Sales does not match Products. In the tests, a trailing space ("CAP-01 ") dropped those units from the product. Capital letters do not matter. Retype the SKU without the space. The summary cell for unmatched units shows how many were lost, and the check cell shows CHECK.

### Every row says Check settings
The as-of date is text, the look-back is not a whole number from 1 to 90, or the warning level is below 0. **Settings B4** shows the reason.

### A sale does not count
Its date or quantity is text. SUMIFS skips it. The summary cell for sales rows to fix counts these rows and the check cell shows CHECK. Retype the value as a real date or number.

### Days left says Check on hand
The on hand cell is blank, text or negative. A blank is not treated as 0, so type a 0 for a product you ran out of.

## Template

[Download the .xlsx](templates/tidy-tabs-days-of-stock-left.xlsx)

The file has four tabs. **Products** lists 12 fictional products in **A5:C44** (40 slots) with the results in **D** to **H** and a summary in **J** and **K**. **Sales** holds 60 fictional sales in **A5:C204**. **Settings** has the yellow as-of date, look-back days, warning level and a settings check. **How to use** has short notes.

Limits: 200 sales rows and 40 products; more was not tested, and neither were look-backs above 90 days. Products prints landscape and the other tabs portrait, all on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
