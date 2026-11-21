---
title: How to Find a Value in a Grid and Return Its Row and Column Label in Excel and Google Sheets
labels: excel-formulas, inventory, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 24 workbooks and 7,512 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Find a value anywhere in a grid with SUMPRODUCT and ROW or COLUMN, then use INDEX on the labels to return its shelf and bin, such as F-Bin 8.
image_prompts: Rows of small plastic parts bins on a steel shelving unit in a home shop storeroom, a few bins holding bubble-wrapped ceramic mugs and candle tins, with a clipboard pick list and a wire hand basket resting on the lower shelf, soft overhead light, realistic photo, no screens, no readable text, no logos
image_alt: Steel shelving with plastic bins of mugs and tins, a clipboard and wire basket
threads: SUMPRODUCT can find a SKU anywhere on a shelf map.\n=SUMPRODUCT((grid=x)*(ROW(grid)-ROW(first_cell)+1)) gives the row, COLUMN gives the column, INDEX turns them into F-Bin 8.\nCount matches first. In LibreOffice, a SKU in 2 bins added both positions and INDEX returned an error.
---
Find the row of a value anywhere in a grid with `=SUMPRODUCT((grid=x)*(ROW(grid)-ROW(first_cell)+1))`, do the same with COLUMN for the column, and use both numbers in INDEX on your row and column labels. In the made-up shelf map, typing SKU-147 returns "F-Bin 8" and MUG-BLU-12 returns "A-Bin 6"; because this only works when the value appears once, the download first counts matches with `=SUMPRODUCT(--(grid=x))` and shows "Not on map" or "In 2 bins" instead of a wrong cell or an error.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (24 workbooks, 7,512 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six-by-eight shelf map of SKUs above a lookup list returning each SKU's bin](images/2026-11-21-team-04-find-value-in-grid-return-row-and-column-label-sumproduct-index-excel-google-sheets/shelf-bin-location-finder-template.png) *Map tab with 47 fictional SKUs on shelves A to F and bins 1 to 8 above the Find tab, where SKU-147 is in F-Bin 8, MUG-BLU-12 is in A-Bin 6, CAN-LAV-08 is in 2 bins (orange on the map) and SKU-200 is Not on map. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. The map must be one block with one SKU per cell.

1. On the **Map** tab, type shelf labels down **A5:A10** and bin labels across **B4:I4**.
2. Type one SKU per bin in **B5:I10** as text, with no spaces before or after. Leave empty bins blank.
3. On the **Find** tab, paste the SKUs to find in **A5:A54**, such as today's pick list.
4. In **B5**, count the bins that hold the SKU: `=IF(TRIM(A5)="","",SUMPRODUCT(--(Map!$B$5:$I$10=A5)))`. The blank check in front matters, because an empty lookup cell matches an empty bin.
5. In **C5**, find the row number: `=IF(B5<>1,"",SUMPRODUCT((Map!$B$5:$I$10=A5)*(ROW(Map!$B$5:$I$10)-ROW(Map!$B$5)+1)))`. The test is 1 in the matching bin and 0 elsewhere, so only that bin's row offset survives.
6. In **D5**, type the same formula with COLUMN in place of ROW for the column number.
7. In **E5**, turn the two numbers into a label: `=IF(B5="","",IF(B5=1,INDEX(Map!$A$5:$A$10,C5)&"-"&INDEX(Map!$B$4:$I$4,D5),IF(B5=0,"Not on map","In "&B5&" bins")))`.
8. In **F5**, type `=IF(B5=1,C5*100+D5,"")` for a walk order: shelf A Bin 1 is 101 and shelf F Bin 8 is 608.
9. Fill **B5:F5** down to row 54, then select **A4:G54** and use **Data > Sort** by Walk order, ascending, to pick in one pass.

Row and column numbers stay blank unless the SKU is in exactly one bin, because two matches add their offsets together.

Microsoft's [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) and [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) pages and Google's [SUMPRODUCT help page](https://support.google.com/docs/answer/3094294) cover each app. Case was checked only in LibreOffice, where the = test ignores it: mug-blu-12 finds MUG-BLU-12.

## Example

The SKUs are fictional. The map holds SKU-100 to SKU-147 across and then down, with MUG-BLU-12 at A-Bin 6, CAN-LAV-08 at both B-Bin 2 and E-Bin 5, the number 1001 at C-Bin 3, and D-Bin 4 empty.

| SKU typed | Times on map | Row | Column | Location | Walk order | Status |
|---|---|---|---|---|---|---|
| SKU-100 | 1 | 1 | 1 | A-Bin 1 | 101 | OK |
| SKU-147 | 1 | 6 | 8 | F-Bin 8 | 608 | OK |
| MUG-BLU-12 | 1 | 1 | 6 | A-Bin 6 | 106 | OK |
| SKU-123 | 1 | 3 | 8 | C-Bin 8 | 308 | OK |
| SKU-131 | 1 | 4 | 8 | D-Bin 8 | 408 | OK |
| mug-blu-12 | 1 | 1 | 6 | A-Bin 6 | 106 | OK |
| "MUG-BLU-12 " (space at end) | 0 | | | Not on map | | Check: space, or number vs text |
| CAN-LAV-08 | 2 | | | In 2 bins | | Check: in 2 bins |
| SKU-200 | 0 | | | Not on map | | Check: not on map |
| 1001 (number) | 1 | 3 | 3 | C-Bin 3 | 303 | OK |
| "1001" (text) | 0 | | | Not on map | | Check: space, or number vs text |

The Find summary shows 11 SKUs listed, 7 found in one bin, 3 not on map, 1 in more than one bin and 4 rows showing Check. The Map has 47 filled bins, 1 empty and 2 holding a repeated SKU, so both CHECK cells show Check on purpose.

## Troubleshooting

### Location shows In 2 bins
The SKU sits in two bins, like CAN-LAV-08. Without the count guard, LibreOffice added both positions to row 7 and column 7 on a 6 x 8 map, and INDEX returned #REF!. Move one or give it its own SKU; the repeated bins turn orange.

### Not on map, but the SKU is right there
The Status says Check: space, or number vs text: a space sits before or after the SKU on one side, or one side holds the number 1001 and the other the text "1001". Retype it and format the map as text (see [keeping leading zeros in ZIP codes](https://tidytabs.blogspot.com/2026/09/how-to-keep-leading-zeros-in-excel-zip.html)). For a SKU truly missing, such as SKU-200, an unguarded INDEX gave #VALUE! in LibreOffice.

### A blank row or the number 0 points at the empty bin
Without the TRIM check, LibreOffice matched a blank row to D-Bin 4. The guard fixes blanks, but the number 0 still matched the empty bin (Times 1, D-Bin 4, OK), since an empty cell equals 0 there. The text "0" showed Not on map, so keep SKUs as text.

### Map CHECK says Check for a SKU with * or ?
In LibreOffice, SKUs such as SKU-1* and SKU-14? were found exactly by the lookup (A-Bin 1 and F-Bin 8), because SUMPRODUCT compares whole values. The repeat check in Map **L6** uses COUNTIF, which read * and ? as wildcards and counted both bins as repeated. If no bin truly repeats, ignore that check or rename the SKUs.

### CHECK says Check with nothing wrong in the list
A SKU was typed below row 54 (**J9** counts it). Only rows 5 to 54 are read, so move it up.

## Template

[Download the .xlsx](templates/tidy-tabs-shelf-bin-location-finder.xlsx)

The file has three tabs. **Map** holds 48 bins in **B5:I10** with checks for filled, empty, repeated and space-padded bins in **K4:L8**. **Find** reads SKUs in rows 5 to 54, returns Times on map, Row number, Column number, Location, Walk order and Status in **B:G**, and sums them up in **I4:J11**. **How to use** explains each formula. Map and Find print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Repeated SKUs turn orange, and every Check turns red. The sheet finds SKUs on your map; it does not track stock counts. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
