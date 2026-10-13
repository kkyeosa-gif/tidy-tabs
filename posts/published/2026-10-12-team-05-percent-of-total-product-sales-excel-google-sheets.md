---
threads_url: https://www.threads.com/@tin_ylab/post/DecFf7wG09e
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-percent-of-total-sales.html
title: How to Calculate Percent of Total Sales by Product in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math, including an unsorted list and a blank-sales case. Excel and Google Sheets were not opened.
search_description: Get each product's share of sales with =B2/SUM($B$2:$B$9), then mark the top sellers that cover 80% of revenue. Free .xlsx template with sample data.
image_prompts: Wooden market table with four stacked bars of handmade soap in graduated heights, a small brass scoop and a kraft paper bag beside them, soft morning light, no screens, no readable text, no logos
image_alt: Graduated stacks of handmade soap bars next to a brass scoop and paper bag
threads: Which product actually pays the bills?\nDivide each sale by the total: =B2/SUM($B$2:$B$9). The dollar signs keep the total put when you fill down.\nIn the sample, 3 of 8 products cover 83% of $10,000.00.
---
Divide each product's sales by the total with `=B2/SUM($B$2:$B$9)`. The dollar signs keep the total fixed when you fill the formula down. Format the column as Percent and the shares add up to 100%.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions. Neither Excel nor Google Sheets was opened for this post.

![Sales by product table with percent of total, running percent and a Top sellers label](images/2026-10-12-team-05-percent-of-total-product-sales-excel-google-sheets/percent-of-total-sales-by-product-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

This works when each product appears once, with its total sales in one column.

1. Type the headers Product, Sales, Percent of total, Running percent and Group in **A1:E1**.
2. List your products in **A2:A9** and their sales in **B2:B9**. Sort the list largest first.
3. In **B10**, type `=SUM(B2:B9)`.
4. In **C2**, type `=B2/SUM($B$2:$B$9)` and fill it down to **C9**.
5. In **C10**, type `=SUM(C2:C9)`. It should show 100%.
6. Select **C2:C10** and format the cells as Percent.
7. In **D2**, type `=SUM($B$2:B2)/SUM($B$2:$B$9)` and fill it down to **D9**. This is the running percent: the share covered so far, from the top of the list.
8. Type 80% in **H2**. This is the share you want your top group to cover.
9. In **E2**, type `=IF(ROUND(D2-C2,6)<$H$2,"Top sellers","Long tail")` and fill it down to **E9**.
10. In **H4**, type `=COUNTIF(E2:E9,"Top sellers")`. In **H5**, type `=SUMIF(E2:E9,"Top sellers",C2:C9)`.

In step 9, `D2-C2` is the running percent before the current row. A product is a top seller if the products above it have not yet reached 80%. That is why the product that crosses 80% is still included.

The 80% line is a rule of thumb, not a law. Change **H2** to try 70% or 90%.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) page. Set the Percent format under **Home > Number Format** or press Ctrl+Shift+%.

### In Google Sheets

The same formulas should work as typed. Set the Percent format under **Format > Number > Percent**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small craft shop with eight products and $10,000.00 in sales.

| Product | Sales | Percent of total | Running percent | Group |
|---|---|---|---|---|
| Soy candles | $4,200.00 | 42.0% | 42.0% | Top sellers |
| Soap bars | $2,600.00 | 26.0% | 68.0% | Top sellers |
| Brass earrings | $1,500.00 | 15.0% | 83.0% | Top sellers |
| Gift boxes | $900.00 | 9.0% | 92.0% | Long tail |
| Stickers | $450.75 | 4.5075% | 96.5075% | Long tail |
| Gift tags | $199.50 | 1.995% | 98.5025% | Long tail |
| Greeting cards | $101.25 | 1.0125% | 99.515% | Long tail |
| Tote bags | $48.50 | 0.485% | 100.0% | Long tail |

**B10** shows $10,000.00 and **C10** shows 100.0%. Soy candles and soap bars reach 68.0%. Brass earrings cross 80% at 83.0%, so they are included. **H4** shows 3 top sellers and **H5** shows 83.0%.

So three of eight products bring in 83% of sales. The other five share 17%.

## Troubleshooting

### Every product says Top sellers
The list is not sorted. In the recalculated file, moving Tote bags ($4,200.00) to row 9 and Soy candles ($48.50) to row 2 made all 8 rows Top sellers, with **H4** at 8 and **H5** at 100%. The running percent only means something when the largest product is on top. Sort by Sales, largest first.

### The percent column shows #DIV/0!
The sales cells are all blank, so the total is zero. With every sale blank, **C2** and **C10** showed #DIV/0! and **B10** showed $0.00. Enter at least one sale, or wrap the formula in `IF` to show nothing when the total is zero.

### The percent column shows decimals like 0.42
The cells are formatted as Number, not Percent. The values are right. Format **C2:C10** and **D2:D9** as Percent.

### Adding a ninth product changes nothing
The ranges stop at row 9. Insert the new row inside the list, above row 9, so the ranges grow with it. Otherwise extend every range by hand and fill the formulas down.

### A big share does not mean a big profit
Percent of sales says nothing about cost or profit. A product with 42% of sales can still lose money after materials and fees. Use this sheet to see where revenue comes from, not what you keep.

## Template

[Download the .xlsx](templates/tidy-tabs-percent-of-total-sales.xlsx)

The file has two tabs. **Sales by product** holds the eight sample products, the percent and running percent columns, the Group column, the total row and the 80% input in **H2** with the results in **H4** and **H5**. The sheet is landscape and prints on one US Letter page. **How to use** has short fill-in notes.

The products and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
