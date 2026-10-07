---
threads_url: https://www.threads.com/@tin_ylab/post/DeNLDsEllNU
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-look-up-shipping-cost-by-weight.html
title: How to Look Up Shipping Cost by Weight in Excel and Google Sheets
labels: excel-formulas, shipping, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux) only. The Excel and Google Sheets steps follow Microsoft Support and Google Docs Editors Help and were not opened or tested here.
search_description: Look up a shipping charge by weight band with VLOOKUP approximate match (TRUE) or INDEX/MATCH, and why the table must be sorted. Free .xlsx.
image_prompts: A small digital postal scale holding a padded mailer, a roll of shipping labels, and a stack of cardboard boxes in a home packing corner, soft window light, all label text blurred and unreadable, no screens, no logos
image_alt: Postal scale weighing a padded mailer beside a label roll and cardboard boxes
threads: Weight bands in a table, rate in a cell: =VLOOKUP(B2,$G$2:$H$7,2,TRUE) does it.\nKeep the bands sorted smallest to largest. In LibreOffice an unsorted table returned $5.75 for 10 oz instead of $7.25, with no error.
---
Type `=VLOOKUP(B2,$G$2:$H$7,2,TRUE)` to get the shipping charge for a weight, with the weight bands in G:H sorted smallest to largest. In the sample table, a 4 oz package returns $5.75.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated. The Excel and Google Sheets formulas are the same and link to Microsoft and Google help pages, but neither app was opened for this post. The wrong answer from an unsorted table was seen in LibreOffice only.

![Orders sheet matching package weights to shipping rates beside the sorted band table](images/2026-10-07-team-02-shipping-cost-weight-tier-lookup-excel-google-sheets/shipping-weight-tier-lookup-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

The idea: each row of the rate table is the lightest weight in a band, plus the price for that band. A package gets the price of the last band it has reached. That is what the last argument, `TRUE`, asks for: the largest "From" value that is less than or equal to the weight.

1. In **G1:H1**, type the headers From (oz) and Rate.
2. In **G2:H7**, type the bands, lightest first: 1 and $4.50, 4 and $5.75, 8 and $7.25, 16 and $9.80, 32 and $13.40, 64 and $18.90.
3. In **B2**, type a package weight in ounces, such as 4.
4. In **C2**, type `=IFERROR(VLOOKUP(B2,$G$2:$H$7,2,TRUE),"Below minimum")`.
5. Press Enter, then fill the formula down for every order.
6. Format the rate cells as currency.

The `$` signs lock the rate table so it does not shift when you fill down. `IFERROR` is covered in Troubleshooting below.

You can get the same answer with INDEX and MATCH:

`=IFERROR(INDEX($H$2:$H$7,MATCH(B2,$G$2:$G$7,1)),"Below minimum")`

The `1` at the end of MATCH means "largest value less than or equal to". It needs the same ascending order. In the template both columns give identical results for all 12 orders.

### In Excel

Type the formulas as shown. Microsoft documents the arguments, including the approximate-match behavior, on its [VLOOKUP function page](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1). This post did not run them in Excel.

### In Google Sheets

Type the same formulas. Google describes the arguments on its [VLOOKUP help page](https://support.google.com/docs/answer/3093318). This post did not run them in Sheets.

For the exact-match version, which finds one SKU in a price list, see [How to Look Up a Price With VLOOKUP in Excel and Google Sheets](https://tidytabs.blogspot.com/2026/10/how-to-look-up-price-with-vlookup-in.html).

## Example

The rates below are made up for this sample and are not any carrier's prices. The orders are fictional too.

Rate table (G:H):

| From (oz) | Rate |
|---|---|
| 1 | $4.50 |
| 4 | $5.75 |
| 8 | $7.25 |
| 16 | $9.80 |
| 32 | $13.40 |
| 64 | $18.90 |

Twelve orders and the rate each one gets:

| Order # | Weight (oz) | Rate |
|---|---|---|
| 2101 | 2.50 | $4.50 |
| 2102 | 3.99 | $4.50 |
| 2103 | 4.00 | $5.75 |
| 2104 | 4.01 | $5.75 |
| 2105 | 7.90 | $5.75 |
| 2106 | 8.00 | $7.25 |
| 2107 | 12.50 | $7.25 |
| 2108 | 16.00 | $9.80 |
| 2109 | 20.00 | $9.80 |
| 2110 | 40.00 | $13.40 |
| 2111 | 64.00 | $18.90 |
| 2112 | 0.50 | Below minimum |

Look at the edges. Exactly 4 oz is in the 4 oz band, so it costs $5.75. At 3.99 oz the package has not reached that band yet, so it costs $4.50. At 4.01 oz it is past the edge and costs $5.75 again.

If you want a package to switch bands at 4.01 oz instead, put 4.01 in the From column.

## Troubleshooting

### The rate is wrong and there is no error
The most likely cause is a rate table that is not sorted from smallest to largest. Approximate match does not check every row, so on an unsorted table it can stop in the wrong place and return a real-looking price.

The template has an **Unsorted table demo** tab with the bands in the order 1, 8, 4, 16, 32, 64. In LibreOffice Calc 24.2.7, a 5 oz package happens to return the right $5.75, but a 10 oz package returns $5.75 when the correct rate is $7.25. No error appears. This was seen in LibreOffice only. Excel and Google Sheets were not tested, and they may return something different on an unsorted table. Sort the table and the problem goes away.

### The result is #N/A
The weight is lighter than the first band. The sample table starts at 1 oz, so a 0.50 oz package has no band and `VLOOKUP` returns `#N/A`. In the template, column E shows the bare formula and gets `#N/A` for order 2112.

Wrapping the formula in `IFERROR(...,"Below minimum")` turns that into readable text. Better, add a band that starts at 0 if you do want to price very light packages.

### IFERROR hides a real mistake
`IFERROR` replaces every error, not only the below-minimum case. A mistyped range or a deleted column also shows "Below minimum". Test the formula without `IFERROR` first, as column E of the template does, then add it.

### The weight is typed as text
A weight pasted from another system can be stored as text, and a text weight will not compare against numbers in the table. This was not reproduced for this post. If a rate looks off, check that the weight cell is right-aligned, which usually means it is a number, or use `=ISNUMBER(B2)`.

### Rates changed but the sheet did not
The formula reads the table, so edit G:H and every order updates. If a rate does not change, check that the lookup range still covers all rows after you add a band. The range `$G$2:$H$7` does not grow by itself, so extend it to `$G$2:$H$8` when you add a seventh band.

## Template

[Download the .xlsx](templates/tidy-tabs-shipping-weight-tier-lookup.xlsx)

The file has three tabs. **Orders** holds the rate table in G:H and 12 sample orders with the VLOOKUP formula, the INDEX/MATCH formula, and a VLOOKUP without `IFERROR`. **Unsorted table demo** shows the wrong-rate case. **How to use** has short notes on each formula.

The rates, bands and orders are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
