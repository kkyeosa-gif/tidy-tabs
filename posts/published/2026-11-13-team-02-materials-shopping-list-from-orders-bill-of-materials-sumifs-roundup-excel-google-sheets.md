---
threads_url: https://www.threads.com/@tin_ylab/post/DfsOvwVGuIM
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-make-materials-shopping-list.html
title: How to Make a Materials Shopping List From Your Orders in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 12 workbooks and 16,872 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Total the supplies your orders need with SUMIFS, subtract what is on the shelf and round up to whole packs to buy. Free Excel template.
image_prompts: A candle maker's workbench with an open bag of soy wax flakes, a row of empty amber glass jars and a bundle of cotton wicks tied with string, a small kitchen scale at the edge, no labels readable.
image_alt: Soy wax flakes, empty amber jars and cotton wicks on a workbench
threads: 595 oz of soy wax for one made-up week of candle orders, and 400 oz on the shelf.\nSUMIFS totals each material from a Recipe tab, MAX(0,F5-C5) gives the shortfall, and ROUNDUP turns 1.22 packs into 2.\nUnits have to agree: type a 10 lb bag as 160 oz.
---
List what goes into one item on a Recipe tab, then total each material over all your orders with `=SUMIFS(Recipe!$E$4:$E$203,Recipe!$B$4:$B$203,A5)` and round the shortfall up to whole packs with `=ROUNDUP(ROUND(G5/D5,9),0)`. In the made-up sample, the week's candle and soap orders need 595 oz of soy wax against 400 oz on hand, so the sheet says buy 2 packs of 160 oz, and six supplies come to $307.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with a Python model (12 workbooks, 16,872 comparisons, 0 mismatches). SUMIFS, ROUNDUP and the other functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Materials tab with amount needed, shortfall, packs to buy and cost for candle supplies](images/2026-11-13-team-02-materials-shopping-list-from-orders-bill-of-materials-sumifs-roundup-excel-google-sheets/materials-shopping-list-template.png) *Ten candle and soap supplies totaled from the week's orders: six to buy for $307.00. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has three working tabs: **Materials**, **Orders** and **Recipe**.

1. On **Orders**, starting at row 4, type one line per product per order or event: who it is for in **A**, the product in **B**, the quantity in **C**.
2. On **Recipe**, type one line per material per product: the product in **A**, the material in **B**, the amount for one item in **C**.
3. In **Recipe D4**, count how many of that product were ordered: `=SUMIFS(Orders!$C$4:$C$203,Orders!$B$4:$B$203,A4)`.
4. In **Recipe E4**, type `=C4*D4` for the material that product needs.
5. On **Materials**, list each supply in **A** with its unit in **B**, what you have on hand in **C**, the pack size in **D** and the pack price in **E**.
6. In **Materials F5**, total the need over every product: `=SUMIFS(Recipe!$E$4:$E$203,Recipe!$B$4:$B$203,A5)`.
7. In **G5**, find the shortfall: `=MAX(0,F5-C5)`.
8. In **H5**, round up to whole packs: `=ROUNDUP(ROUND(G5/D5,9),0)`. You cannot buy 1.22 packs.
9. In **I5**, multiply packs by price for the cost to buy. Fill each formula down.

Units must agree: Amount per item uses the unit on the Materials tab. The file does not convert pounds, so type a 10 lb bag as 160 oz.

The inner `ROUND(...,9)` in step 8 stops a decimal leftover such as 3.0000000001 from adding a pack. It is a precaution: in LibreOffice, `ROUNDUP(0.9/0.3,0)` already gave 3 without it.

The template wraps these formulas so a name such as "Candle*" is not read as a wildcard, and Status columns mark text, negative or blank numbers as Check. Lines marked Check or No recipe are left out of the totals, not guessed.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7), and Google's [SUMIFS](https://support.google.com/docs/answer/3238496). They describe the functions, not our test.

## Example

Sample data is fictional: a candle and soap maker's week of gift shop, farmers market and Etsy orders, totaling 40 lavender candles, 24 cedar candles, 42 mini candles and 50 soap bars.

| Material | On hand | Pack size | Pack price | Needed | Short | Packs | Cost | Status |
|---|---|---|---|---|---|---|---|---|
| Soy wax (oz) | 400 | 160 | $32.00 | 595 | 195 | 2 | $64.00 | Buy |
| Lavender fragrance oil (oz) | 20 | 16 | $28.50 | 36.6 | 16.6 | 2 | $57.00 | Buy |
| Cedar fragrance oil (oz) | 12 | 16 | $26.00 | 14.4 | 2.4 | 1 | $26.00 | Buy |
| Cotton wick (each) | 150 | 100 | $9.00 | 106 | 0 | 0 | $0.00 | Enough |
| 8 oz amber jar (each) | 36 | 24 | $30.00 | 64 | 28 | 2 | $60.00 | Buy |
| 4 oz tin (each) | 50 | 24 | $14.40 | 42 | 0 | 0 | $0.00 | Enough |
| Jar label (each) | 300 | 250 | $18.00 | 106 | 0 | 0 | $0.00 | Enough |
| Soap base (oz) | 64 | 32 | $15.50 | 225 | 161 | 6 | $93.00 | Buy |
| Colloidal oatmeal (oz) | 8 | 16 | $9.25 | 5 | 0 | 0 | $0.00 | Enough |
| Soap band (each) | 40 | 100 | $7.00 | 50 | 10 | 1 | $7.00 | Buy |

Soy wax: 8 oz candles take 7 oz and minis 3.5 oz, so 40 x 7 + 24 x 7 + 42 x 3.5 = 595 oz. That is 195 short of 400 on hand; 195 / 160 is 1.22, rounded up to 2 packs, leaving 125 oz.

The summary shows 6 materials to buy for $307.00. CHECK says Check on purpose: the 10 Holiday gift sets have no Recipe lines, so that order shows No recipe and is left out. Add the gift set's lines to Recipe (or delete the order line) and CHECK turns to OK.

Waste, spills and test pours are not included; raise Amount per item if you always lose some. This is a shopping list, not inventory accounting or cost of goods for taxes.

## Troubleshooting

### An order is not counted
Orders shows No recipe: the product has no Recipe lines, or the name differs. Case does not matter, but "Lavender soy candle 8 oz " with a trailing space showed No recipe.

### A material shows 0 needed or Not used
The Recipe name differs from the Materials tab, and Recipe shows Not on Materials tab. Copy the name across so they match.

### Packs to buy shows Check
You are short and the pack size is 0, blank or negative, or On hand is blank or text. Type a number.

### Cost to buy is blank
Pack price is empty, so the cost stays blank on purpose.

### CHECK says Check
An order line is not counted or a recipe line has a problem (in the sample, 1 order line). Fix those before you trust the totals.

## Template

[Download the .xlsx](templates/tidy-tabs-materials-from-orders.xlsx)

**Materials** holds up to 100 supplies (rows 5 to 104) with the formula columns and a summary with a CHECK cell in **M4:N9**; it prints Landscape on US Letter. **Orders** and **Recipe** each hold 200 lines with a Status column, and **How to use** explains each step.

Yellow cells are inputs and gray cells are formulas. Status shows Buy in orange, Enough in green and Check in red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
