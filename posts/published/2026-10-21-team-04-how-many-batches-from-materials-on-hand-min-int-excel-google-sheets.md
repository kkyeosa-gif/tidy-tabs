---
threads_url: https://www.threads.com/@tin_ylab/post/Dey83XJmxoY
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-how-many-batches-you-can.html
title: How to Find How Many Batches You Can Make From the Materials You Have in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including edited copies (material at exactly one batch, short by a cent, zero on hand, ties, blank and zero needed amounts, text and negative inputs, decimals). Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Find how many batches you can make from materials on hand with INT and MIN in Excel and Google Sheets. Free template that names the limiting material.
image_prompts: A kitchen scale holding a block of soy wax beside a row of empty glass candle jars and a spool of cotton wick on a workshop table, soft natural window light, no screens, no readable text, no logos
image_alt: Scale with a block of wax beside empty glass jars and a wick spool
threads: Wax says 9 batches. Jars say 7.\nSo it is 7: =INT(B2/C2) per material, then =MIN(D2:D7). INDEX and MATCH name the jars as the one that runs out first.\nThat is 84 candles.
---
Divide what you have by what one batch needs for each material with `=INT(B2/C2)`, then take the smallest result with `=MIN(D2:D7)`. That number is how many whole batches you can make, and the material behind it is the one that runs out first.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Materials sheet listing six supplies with whole batches possible and the limiting material shaded](images/2026-10-21-team-04-how-many-batches-from-materials-on-hand-min-int-excel-google-sheets/batches-from-materials-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

Lay out one row per material: the name in column A, what you have on hand in B, and what one batch needs in C.

1. In **A2:A7**, type the material names, for example Soy wax (lb).
2. In **B2:B7**, type the amount on hand, for example 42.5.
3. In **C2:C7**, type the amount one batch needs, for example 4.5.
4. In **D2**, type `=INT(B2/C2)` and fill it down to **D7**. This is the whole batches that material allows.
5. In **B10**, type `=MIN(D2:D7)`. This is how many whole batches you can make.
6. In **B11**, type `=INDEX(A2:A7,MATCH(B10,D2:D7,0))`. This names the material that runs out first.

`INT` drops the fraction, because 7.25 jars' worth is 7 batches, not 7.25. `MIN` picks the scarcest material, since every batch needs all of them.

If two materials tie for the smallest number, MATCH returns the first one listed, so that is the one named.

### Guards in the template

The template's D2 formula is longer than `=INT(B2/C2)`:

`=IF(AND(B2="",C2=""),"",IF(OR(NOT(ISNUMBER(B2)),AND(C2<>"",NOT(ISNUMBER(C2)))),"Check",IF(OR(B2<0,C2<0),"Check",IF(OR(C2="",C2=0),"",INT(ROUND(B2/C2,10))))))`

It does four things:

- A blank or 0 in the needed column leaves the row empty, so you never see #DIV/0!. The sheet ignores that material on purpose instead of guessing.
- Text or a negative number shows Check.
- A blank on-hand cell with a number needed also shows Check.
- `ROUND(...,10)` inside INT guards against tiny binary rounding errors in a division. In LibreOffice a bare `INT(B2/C2)` gave the same answers on the cases we tried (0.3/0.1, 0.7/0.1, 1.1/0.1, 4.35/0.05), so the guard was not shown to be needed there. Excel was not run.

B10 is wrapped too: it shows Check inputs if any row says Check, and No materials if every row is blank.

### Leftovers and what to buy

- **E2** `=IF(AND(ISNUMBER(D2),ISNUMBER($B$10)),ROUND(B2-C2*$B$10,6),"")` is what is left after the batches you can make.
- **F2** `=IF(AND(ISNUMBER(D2),ISNUMBER(Batch!$B$2)),ROUND(MAX(0,C2*Batch!$B$2-B2),6),"")` is how much to buy to make N batches in total. N is typed on the Batch tab in **B2**. It is the total needed minus what you have, not N more batches on top of the ones you can already make.

The Batch tab also holds the yield in **B1**, how many candles one batch makes. **B12** multiplies it by the batches you can make.

### In Excel

The formulas work as typed. Microsoft describes the arguments on its [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) pages. Those are the help pages, not our test.

### In Google Sheets

The same formulas should work as typed. Google lists INT, MIN, INDEX and MATCH on its [function list](https://support.google.com/docs/table/25273). Sheets was not opened for this post, so compare the result with the table below.

## Example

Sample data is fictional: a soy candle batch that makes 12 candles, with six materials on hand.

| Material | On hand | Needed per batch | Whole batches |
|---|---|---|---|
| Soy wax (lb) | 42.5 | 4.5 | 9 |
| Fragrance oil (oz) | 30 | 3.5 | 8 |
| Cotton wicks | 130 | 12 | 10 |
| 8 oz jars | 87 | 12 | 7 |
| Labels | 150 | 12 | 12 |
| Shipping boxes | 100 | 12 | 8 |

The smallest number is 7, from the jars, so you can make 7 batches. That is 84 candles. After 7 batches the jars have 3 left, which is less than the 12 one batch needs.

Planning for 10 batches, the "to buy" column shows what is short in total:

| Material | To buy for 10 batches |
|---|---|
| Soy wax (lb) | 2.5 |
| Fragrance oil (oz) | 5 |
| Cotton wicks | 0 |
| 8 oz jars | 33 |
| Labels | 0 |
| Shipping boxes | 20 |

B14 counts the materials to buy, which is 4 here. **B15** shows OK when no leftover is negative and the limiting material has less than one batch left over.

The sheet shades the limiting row. This is arithmetic only. It makes no claim about yield, shelf life or food safety.

## Troubleshooting

### Two materials tie, which one is named
The first one listed. With jars and boxes both at 7 batches, B11 shows 8 oz jars because it sits higher in the list. The batch count is the same either way.

### A row is empty and there is no #DIV/0!
A blank or 0 in Needed per batch is ignored on purpose. If you meant to include that material, type its amount. With jars left blank, fragrance oil becomes the limit at 8 batches.

### B10 says Check inputs
At least one row has text, a negative number, or a blank on-hand cell next to a needed amount. Entries like "thirty" or "3.5 oz" count as text. Retype them as plain numbers, and keep the unit in the material name.

### The answer is 0 batches because of a few cents
One batch needs exactly 4.5 lb of wax. With 4.49 lb on hand, the sheet says 0 batches and names Soy wax (lb). To make 10 batches you would need 40.51 lb more. The sheet does not round up for you.

### To buy shows 0 for some materials
That material is already enough for N batches. Cotton wicks and labels show 0 in the 10 batch example.

One thing the sheet cannot catch: mixed units. If wax is in pounds on hand but ounces per batch, the numbers will be wrong with no warning. That case was not tested.

## Template

[Download the .xlsx](templates/tidy-tabs-batches-from-materials.xlsx)

The file has three tabs. **Batch** holds the yellow settings: the yield in **B1** and the number of batches to plan for in **B2**. **Materials** holds the yellow inputs in **A2:C7**, the per material results in **D2:F7**, and the summary in **B10:B15**, with a CHECK cell in **B15**. **How to use** has short notes. All tabs print portrait on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
