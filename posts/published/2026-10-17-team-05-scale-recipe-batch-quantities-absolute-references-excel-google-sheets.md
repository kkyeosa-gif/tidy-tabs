---
threads_url: https://www.threads.com/@tin_ylab/post/Deo9eJ-m6Hc
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-scale-recipe-up-or-down-in-excel.html
title: How to Scale a Recipe Up or Down in Excel and Google Sheets
labels: excel-formulas, small-business, food-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 8 sample ingredients were checked against a Python Decimal calculation at factors 0.5, 1, 2.5, 5 and 10, plus a 1.1 factor, a blank or zero target, a zero base yield and a quantity typed as text. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Scale a recipe up or down with =ROUND(B6*$C$3,2) in Excel and Google Sheets. Absolute reference ($C$3) keeps the scale factor fixed. Free template.
image_prompts: Row of empty glass jam jars with metal lids beside a kitchen scale and a bowl of fresh strawberries on a wooden counter, soft morning window light, no screens, no readable text, no logos
image_alt: Empty glass jars lined up beside a kitchen scale and a bowl of strawberries
threads: Turning a 12 jar jam batch into 30 jars means one multiplier, not 8 retyped amounts.\n=ROUND(B6*$C$3,2) with the $ signs locks the factor in C3 while you fill down.\n3 lemons at 2.5x is 7.5, so the sheet also shows 8 to buy.
---
Divide the new yield by the base yield to get a scale factor, then multiply each ingredient by it with the factor locked in place: `=ROUND(B6*$C$3,2)`. The `$` signs keep the reference on C3 when you fill the formula down the ingredient list.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Recipe table with a 2.50x scale factor, scaled quantities and a buy whole units column](images/2026-10-17-team-05-scale-recipe-batch-quantities-absolute-references-excel-google-sheets/scale-recipe-batch-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The yields sit at the top, with labels in **A** and values in **C**. Below them is one ingredient per row: ingredient in **A**, base quantity in **B**, unit in **C**, scaled quantity in **D** and whole units to buy in **E**.

1. On a tab named **Recipe**, type the labels Base batch yield (jars), Target yield (jars), Scale factor and Check in **A1:A4**. Type the base yield in **C1** (for example 12) and the target yield in **C2** (for example 30).
2. In **C3**, type `=IF(AND(ISNUMBER(C1),C1>0),IF(ISNUMBER(C2),C2/C1,0),"")`. This divides the target by the base and avoids a #DIV/0! error when the base is 0 or blank. Give it the custom number format `0.00"x"` to show 2.50x.
3. In row 5, type the headers: Ingredient, Base quantity, Unit, Scaled quantity, Buy whole units.
4. From row 6 down, type each ingredient, its quantity for the base batch and its unit (the template uses a dropdown of oz, cup, tsp and each). Type fractions as decimals: 0.33 for 1/3 cup.
5. In **D6**, type `=IF(OR(C$3="",NOT(ISNUMBER(B6))),"",ROUND(B6*$C$3,2))`. The part that does the work is `ROUND(B6*$C$3,2)`; the rest leaves the cell blank when the factor or the quantity is missing.
6. In **E6**, type `=IF(D6="","",IF(C6="each",ROUNDUP(ROUND(B6*$C$3,6),0),ROUND(B6*$C$3,2)))`. Items counted in whole pieces, such as lemons, round up. Everything else rounds to 2 decimals.
7. Select **D6:E6** and drag the fill handle down to row 30. The template shows both columns with two decimals, so 8 appears as 8.00.
8. Change **C2** to any new yield and every row updates. The template also has a status formula in **C4** that shows OK or a Check message; you can skip it.

### Why the dollar signs matter

`B6` is relative: filling down turns it into B7, B8 and so on, so each row uses its own quantity. `$C$3` is absolute, so every row multiplies by the same factor. Without the `$` signs, row 7 would point at C4, which holds the status message instead of a number.

Press **F4** while typing a reference in Excel to add the `$` signs. Microsoft explains this on its page about [relative, absolute and mixed references](https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9).

### In Excel

The formulas work as typed. Microsoft describes the rounding functions on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) pages.

### In Google Sheets

The same formulas should work as typed, and you can fill down by dragging the small square at the corner of the selected cell. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small jam kitchen scaling a 12 jar batch to 30 jars, so **C3** shows 2.50x.

| Ingredient | Base quantity | Unit | Scaled quantity | Buy whole units |
|---|---|---|---|---|
| Strawberries | 64 | oz | 160 | 160 |
| Granulated sugar | 6 | cup | 15 | 15 |
| Lemon juice | 0.25 | cup | 0.63 | 0.63 |
| Powdered pectin | 4 | tsp | 10 | 10 |
| Lemons | 3 | each | 7.5 | 8 |
| Vanilla extract | 1.5 | tsp | 3.75 | 3.75 |
| Salt | 0.25 | tsp | 0.63 | 0.63 |
| Jar lids | 12 | each | 30 | 30 |

Lemon juice is 0.625 before rounding, so it shows 0.63. You can't buy 7.5 lemons, so the last column shows 8. Only the unit `each` rounds up. Factors of 0.5, 1, 5 and 10 also matched Python for all 8 rows.

## Troubleshooting

### Every quantity shows 0
**C2** is blank, 0 or text such as thirty. The factor becomes 0 and **C4** says Check target yield. Type a number in **C2**.

### The factor and the quantities are blank
**C1** is 0 or blank, so nothing can be divided. **C3**, **D** and **E** stay empty and **C4** says Check base yield, with no #DIV/0! error.

### One ingredient has no scaled quantity
A quantity was typed as text, such as the word six for sugar. That row stays blank, the others are right, and **C4** says Check quantities.

### Whole units round up one too many
Some decimals carry tiny errors: 3 x 1.1 is held as 3.3000000000000003, which `ROUNDUP` could push to 4. The extra `ROUND(...,6)` in **E6** removes that noise first. In the recalculated copy, a 1.1 factor (target 11 on a base of 10) gave 3.3 lemons and 4 to buy, and 13.2 lids and 14 to buy.

### Every row uses the wrong factor
You probably wrote `C3` instead of `$C$3`. Click into the formula in **D6**, add the `$` signs, then fill down again. This case was not run.

## Limits

This is straight multiplication. Cooking times, pan sizes, leavening, jam set points and food safety at larger sizes can change when you scale, and the sheet does not cover them. The formulas run from row 6 to row 30.

## Template

[Download the .xlsx](templates/tidy-tabs-scale-recipe-batch.xlsx)

The file has two tabs. **Recipe** holds the yields in yellow input cells **C1:C2**, the scale factor and status in **C3:C4**, and 8 sample ingredients in rows 6 to 13. Formulas in **D** and **E** run to row 30. **How to use** has short notes, and both tabs print on US Letter.

The ingredients are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
