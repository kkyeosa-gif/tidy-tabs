---
threads_url: https://www.threads.com/@tin_ylab/post/DebxtZHoKnG
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-cost-per-unit-for.html
title: How to Calculate Cost Per Unit for a Handmade Batch in Excel and Google Sheets
labels: excel-formulas, handmade-business, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math, including a zero-sellable case, a zero units-per-package case and a text-quantity case. Excel and Google Sheets were not opened; the one help link is from Microsoft Support.
search_description: Price each material with =ROUND(B2/C2*D2,2), add the rows, then divide by sellable items. Sample candle batch: $151.26, or $6.88 each. Free .xlsx.
image_prompts: Wooden workbench with a block of soy wax flakes in a measuring pitcher, a row of empty tin candle jars, and a coil of cotton wick, soft window light, no screens, no readable text, no logos
image_alt: Soy wax flakes in a pitcher beside empty tin jars and a coil of cotton wick
threads: Pricing a candle batch? Add each material with =ROUND(B2/C2*D2,2), then divide by what you can actually sell.\nThe sample batch costs $151.26 in materials.\n24 made, 2 lost, so it is $6.88 per candle, not $6.30.
---
Price each material with `=ROUND(B2/C2*D2,2)` (package price divided by units per package, times units used), add the rows, and divide the total by the items you can sell with `=ROUND(E8/H4,2)`. The sample candle batch costs $151.26 in materials, or $6.88 per sellable candle.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Candle batch materials priced line by line with a total and cost per sellable candle](images/2026-10-12-team-04-batch-cost-per-unit-excel-google-sheets/batch-cost-per-unit-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: turn each package price into a price per unit, multiply by the units you used, and add it all up. Then spread the total over the items you can actually sell, not the items you made.

1. In row 1, type these headers in A to E: Material, Package price, Units per package, Units used, Cost.
2. In **A2:A7**, type one material per row.
3. In **B2:B7**, type what you paid for one package, such as $60.00.
4. In **C2:C7**, type how many units are in that package, such as 12 jars in a case.
5. In **D2:D7**, type how many units this batch used, such as 24 jars.
6. In **E2**, type `=ROUND(B2/C2*D2,2)` and fill it down to **E7**. This is the cost of that material in this batch, rounded to cents.
7. In **E8**, type `=SUM(E2:E7)`. This is the materials total.
8. In **G2**, type Items made and put the number in **H2**. In **G3**, type Items lost or unsellable and put the number in **H3**.
9. In **G4**, type Sellable items. In **H4**, type `=H2-H3`.
10. In **G5**, type Cost per item made. In **H5**, type `=IF(H2<=0,"",ROUND(E8/H2,2))`.
11. In **G6**, type Cost per sellable item. In **H6**, type `=IF(H4<=0,"",ROUND(E8/H4,2))`.

Keep the units the same inside a row. If the wax is priced per pound, enter the pounds used, not ounces.

H6 is the number to use when you think about what a batch cost you. Lost or cracked items still cost money, so their cost lands on the ones you can sell.

### In Excel

The formulas work as typed. Microsoft documents the rounding on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) page. Format **E2:E8** and **H5:H6** as Currency from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Format the cells with **Format > Number > Currency**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a batch of 24 soy candles from a small shop.

| Material | Cost in this batch |
|---|---|
| Soy wax | $18.90 |
| Fragrance oil | $4.20 |
| Cotton wicks | $2.88 |
| Tin jars | $120.00 |
| Labels | $4.56 |
| Dye chips | $0.72 |
| **Total (E8)** | **$151.26** |

The tin jars come in cases of 12 at $60.00 a case, and the batch used 24. So the formula is $60.00 / 12 x 24, which is $120.00.

| Result | Cell | Value |
|---|---|---|
| Items made | H2 | 24 |
| Lost or unsellable | H3 | 2 |
| Sellable items | H4 | 22 |
| Cost per item made | H5 | $6.30 |
| Cost per sellable item | H6 | $6.88 |

Two candles were lost, so each of the 22 you can sell carries $0.58 more than the plain $6.30 average. The results matched a Python check that rounds each line to cents first.

## Troubleshooting

### Cost per sellable item is blank
Items lost equals items made, so H4 is 0 and the formula returns an empty cell on purpose. In the recalculated file, setting H3 to 24 gave a blank H6 and no divide by zero error. Fix the counts in **H2:H3**.

### #DIV/0! in the cost column and the total
A units-per-package cell is 0 or empty. In the recalculated file, a 0 in **C2** made E2, E8 and H5 show #DIV/0!. Enter the real number of units in the package.

### #VALUE! in the cost column and the total
A quantity cell has text in it. In the recalculated file, typing "2.4 oz" in **D3** made E3 and E8 show #VALUE!. Type only the number, such as 2.4, and keep the unit in the header or the material name.

### The total is a cent or two off a hand calculation
Each line is rounded to cents before it is added. Adding unrounded lines and rounding once at the end can differ by a cent or two. Pick one method and keep it for every batch.

### The cost per item looks too low
This sheet covers materials only. Labor, packaging bought separately, fees, shipping, overhead and taxes are not included, and it does not tell you what to charge. Add those costs yourself before you compare the number to a price.

## Template

[Download the .xlsx](templates/tidy-tabs-batch-cost-per-unit.xlsx)

The file has two tabs. **Batch cost** is set to print landscape on one US Letter page. It holds the six sample materials in A2:D7, the line costs in E2:E7, the total in E8, and the yellow inputs and results in H2:H6. **How to use** has short fill-in notes.

The materials and prices are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
