---
title: How to Round Prices to the Nearest $0.05 or .99 in Excel and Google Sheets
labels: excel-formulas, pricing, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same functions; the ROUND, ROUNDUP, MROUND and CEILING links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Round prices to the nearest nickel with =MROUND(B2,0.05) or end them in .99 with =ROUNDUP(B2,0)-0.01. Free .xlsx template with eight sample Etsy prices.
image_prompts: Handwritten-style price tags on a wooden craft fair table beside a small stack of ceramic mugs and a tin cash box with a few nickels beside it, soft daylight, all tag text blurred and unreadable, no screens, no logos
image_alt: Blank price tags next to stacked ceramic mugs and a small cash tin of nickels
threads: Markup math left you with $13.4167?\n=MROUND(B2,0.05) snaps it to the nearest nickel: $13.40.\n=ROUNDUP(B2,0)-0.01 gives $13.99. Careful, a price that is already $22.00 comes out as $21.99.
---
Use `=MROUND(B2,0.05)` to round a price to the nearest nickel, or `=ROUNDUP(B2,0)-0.01` to end it in .99. A price of $13.4167 becomes $13.40 with the first formula and $13.99 with the second.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Price sheet with eight items showing the raw price, nearest cent, nearest nickel, rounded up and .99 results](images/2026-10-08-team-03-round-prices-nearest-nickel-99-excel-google-sheets/round-prices-nearest-nickel-99-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Markup and discount math leaves prices like $13.4167. Put those raw prices in column B, starting in **B2**, and add one formula column for each rounding style.

1. In row 1, type these headers in A to F: Item, Price after markup math, Nearest cent, Nearest $0.05, Up to next $0.05, Ends in .99.
2. In **A2**, type the item name. In **B2**, type the unrounded price, such as 13.4167.
3. In **C2**, type `=ROUND(B2,2)`. This rounds to the nearest cent, so $13.4167 becomes $13.42.
4. In **D2**, type `=MROUND(B2,0.05)`. This rounds to the nearest nickel, so $13.4167 becomes $13.40.
5. In **E2**, type `=CEILING(B2,0.05)`. This always goes up to the next nickel, so $13.4167 becomes $13.45.
6. In **F2**, type `=ROUNDUP(B2,0)-0.01`. This goes up to the next whole dollar and takes off a cent, so $13.4167 becomes $13.99.
7. Select **C2:F2** and format the cells as currency with two decimals.
8. Fill **C2:F2** down for each item.

Pick one column as your price. The others are there so you can compare.

`=CEILING(B2,1)-0.01` gives the same .99 result as the ROUNDUP version.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [ROUND](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), [MROUND](https://support.microsoft.com/en-us/office/mround-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427) and [ROUNDUP](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) function pages.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [ROUND](https://support.google.com/docs/answer/3093440?hl=en), [ROUNDUP](https://support.google.com/docs/answer/3093443?hl=en) and [CEILING](https://support.google.com/docs/answer/3093471?hl=en) function pages. The MROUND page was not checked for Sheets here, so confirm `=MROUND(B2,0.05)` in your own sheet.

## Example

Sample data below is fictional. It is eight items for a small Etsy-style shop, with prices left over from markup math.

| Item | Price after markup math | Nearest cent | Nearest $0.05 | Up to next $0.05 | Ends in .99 |
|---|---|---|---|---|---|
| Beeswax candle, 4 oz | $13.4167 | $13.42 | $13.40 | $13.45 | $13.99 |
| Cotton tote bag | $8.2333 | $8.23 | $8.25 | $8.25 | $8.99 |
| Hand-poured soap set | $24.7083 | $24.71 | $24.70 | $24.75 | $24.99 |
| Ceramic mug | $17.9625 | $17.96 | $17.95 | $18.00 | $17.99 |
| Greeting card, single | $6.1300 | $6.13 | $6.15 | $6.15 | $6.99 |
| Linen napkins, set of 4 | $31.4467 | $31.45 | $31.45 | $31.45 | $31.99 |
| Sticker sheet | $12.9833 | $12.98 | $13.00 | $13.00 | $12.99 |
| Embroidered patch | $22.0000 | $22.00 | $22.00 | $22.00 | $21.99 |

Check the first row by hand. $13.4167 divided by 0.05 is 268.33, which rounds to 268, and 268 times 0.05 is $13.40. CEILING rounds that 268.33 up to 269, which is $13.45.

MROUND can go down and CEILING cannot. The mug is $17.9625, and the nearest nickel is $17.95, but CEILING gives $18.00. The sticker sheet is $12.9833: the nearest nickel is $13.00 and the .99 ending is $12.99.

The last row is a catch. The patch is already $22.00, so `=ROUNDUP(B2,0)-0.01` gives $21.99, one cent below the price you started with.

## Troubleshooting

### The .99 formula lowers a whole-dollar price
ROUNDUP leaves a whole number as it is, then the formula subtracts $0.01. A price of $22.00 becomes $21.99. If you do not want that, use the price as is for whole dollars, or round to the nearest cent first and check those rows by hand.

### The result still shows more than two decimals
The cell has General or a decimals format. Format the result column as currency with two decimals. ROUND, MROUND and CEILING change the value itself. Formatting only changes how it looks.

### MROUND and CEILING give different answers
MROUND picks the nearest nickel, up or down. CEILING always goes up. In the example, the mug is $17.95 with MROUND and $18.00 with CEILING.

### The formula shows an error
Check that the price in column B is a number and not text. A price pasted with a dollar sign as text will not round. Text-formatted prices were not tested in this template.

### Negative prices or a negative step
Negative values were not checked in the template. In a scratch test, LibreOffice returned $13.40 for `=MROUND(13.4167,-0.05)`. Excel may behave differently, so keep both the price and the step positive.

## Template

[Download the .xlsx](templates/tidy-tabs-round-prices-nearest-nickel-99.xlsx)

The file has two tabs. **Prices** holds the eight sample items with the raw price in column B and the four formulas in columns C to F. **How to use** has short fill-in notes and the $13.4167 example.

The items and prices are fictional. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608?hl=en).
