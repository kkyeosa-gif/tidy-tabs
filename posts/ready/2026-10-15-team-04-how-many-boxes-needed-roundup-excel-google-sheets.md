---
title: How to Calculate How Many Boxes You Need in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math.ceil, including an exact multiple, zero or blank units per box, units typed as text, a decimal quantity, a zero-unit order and an empty sheet. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Boxes needed: =ROUNDUP(B2/C2,0). Shipping 50 candles at 12 per box needs 5 boxes, not 4. Free .xlsx with box cost and empty slots.
image_prompts: Three small brown cardboard shipping boxes, one open with crumpled kraft paper and a few soy candle jars inside, beside a roll of packing tape on a wooden packing table, soft window light, no screens, no readable text, no logos
image_alt: Open cardboard shipping box with candle jars and kraft paper next to a tape roll
threads: Your box math is probably one short.\n50 candles at 12 per box is 4.17, so you need 5 boxes.\n=ROUNDUP(B2/C2,0) rounds up every time, and the sample sheet shows the 10 empty slots in the last box.\nA quantity typed as text gets skipped, so the post covers that.
---
Get the box count with `=ROUNDUP(B2/C2,0)`, where B2 is units to ship and C2 is units per box. Shipping 50 candles at 12 per box needs 5 boxes, not 4, because 50 divided by 12 is 4.17.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Orders table with boxes needed, last box units, empty slots and box cost](images/2026-10-15-team-04-how-many-boxes-needed-roundup-excel-google-sheets/boxes-needed-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

ROUNDUP always rounds away from zero, so 4.17 becomes 5. A normal ROUND would give you 4 boxes and 2 candles with nothing to put them in. The 0 at the end means zero decimal places.

1. On a tab named **Orders**, type these headers in row 1: Item, Units to ship, Units per box, Boxes needed, Units in last box, Empty slots, Box cost.
2. Type the item name in column **A**, the number of units to ship in **B**, and how many fit in one box in **C**.
3. In **J2**, type your cost per box, for example 1.45.
4. In **D2**, type `=IF(AND(ISNUMBER(B2),ISNUMBER(C2),C2>0),ROUNDUP(B2/C2,0),"")`. The checks leave the cell blank instead of showing #DIV/0! when units per box is 0 or empty.
5. In **E2**, type `=IF(OR(D2="",B2=0),"",B2-(D2-1)*C2)`. This is how many units land in the last box.
6. In **F2**, type `=IF(D2="","",D2*C2-B2)`. This is the empty slots across all boxes.
7. In **G2**, type `=IF(D2="","",D2*$J$2)`. This is the box cost for the order.
8. Copy D2:G2 down as far as you have orders. The template fills to row 30.
9. For totals, type `=SUM(D2:D30)` for boxes, `=SUM(G2:G30)` for box cost and `=SUM(F2:F30)` for empty slots.

### In Excel

The formulas work as typed. Microsoft describes the functions on its [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) and [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) pages. Copy the formulas down by dragging the fill handle at the bottom right corner of D2:G2.

### In Google Sheets

The same formulas should work as typed. Drag the blue square at the bottom right corner of the selected cells to fill down. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small shop's shipping queue, with a cost of $1.45 per box.

| Item | Units to ship | Units per box | Boxes | Last box | Empty slots | Box cost |
|---|---|---|---|---|---|---|
| Soy candle | 50 | 12 | 5 | 2 | 10 | $7.25 |
| Soap | 30 | 8 | 4 | 6 | 2 | $5.80 |
| Earring cards | 17 | 20 | 1 | 17 | 3 | $1.45 |
| Mugs | 26 | 6 | 5 | 2 | 4 | $7.25 |
| Tote bags | 48 | 12 | 4 | 12 | 0 | $5.80 |
| Birthday cards | 130 | 25 | 6 | 5 | 20 | $8.70 |
| Gift tag sets | 0 | 40 | 0 | (blank) | 0 | $0.00 |

The totals come to 25 boxes, $36.25 in box cost and 39 empty slots. All 25 box counts matched Python math.ceil.

Tote bags show the exact-multiple case. 48 divided by 12 is exactly 4, so the last box is full and nothing is wasted. The empty slots column is useful when you are deciding whether a smaller box, or a bigger order, would waste less.

## Troubleshooting

### The box count is blank
The row is skipped when units per box is 0 or empty, or when the units are not a number. In the recalculated copy, a units per box of 0 left D to G blank and the box total fell to 20. Type a number above 0.

### A row is missing from the totals
The units to ship was probably stored as text. In the recalculated copy, "30" typed as text left the soap row blank and the box total at 21. Nothing warns you, so blank rows in column D are worth a look. Text usually lines up on the left of the cell. Retype it as a plain number.

### The count looks too low by one
If you used ROUND or plain division, the last partial box gets dropped. Use ROUNDUP, which counts any leftover units as one more box.

### A decimal quantity gives a half-empty box
Decimals are accepted. In the recalculated copy, 10.5 units at 5 per box gave 3 boxes, 0.5 units in the last box and 4.5 empty slots. Real units should be whole numbers, so check for a typo.

### All results show 0
With every row cleared, column D stays blank and the three totals read 0. That is expected for an empty sheet.

## Template

[Download the .xlsx](templates/tidy-tabs-boxes-needed.xlsx)

The file has two tabs. **Orders** holds the 7 sample rows in yellow input cells, formulas filled to row 30, the cost per box in J2 and the three totals in J4 to J6. **How to use** has short fill-in notes.

The template handles one box size per row. It does not check weight, box dimensions or carrier rules, so confirm those with your carrier. Text input drops a row from the totals with no warning.

The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
