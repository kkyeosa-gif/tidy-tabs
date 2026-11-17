---
title: How to Pick the Smallest Shipping Box That Fits Each Item in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 28 workbooks and 14,364 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Find which box size fits each item: sort sides into longest, middle and shortest, add padding, and INDEX/MATCH returns the first fit. Free template.
image_prompts: A packing table in a home shipping corner with three empty brown corrugated boxes of different sizes standing in a row, a ceramic mug wrapped in kraft paper sitting inside the smallest box with barely any room, and a long cardboard poster tube lying across the tops too long for any of them, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Three cardboard boxes on a packing table beside a wrapped mug and poster tube
threads: Sort every side into longest, middle and shortest before you match an item to a box.\nThen =INDEX(boxes,MATCH(1,INDEX((long>=a)*(mid>=b)*(short>=c),0),0)) returns the first box that fits.\nWith 1 in of padding, a 3 x 3 x 4 candle jar gets the 6x4x4 and a 24 in poster tube gets No box fits.
---
Sort each box and each item into longest, middle and shortest side with `MAX`, `SUM-MAX-MIN` and `MIN`, then return the first box on your list where all three sides fit with `=INDEX(boxes,MATCH(1,INDEX((long>=a)*(mid>=b)*(short>=c),0),0))`, adding your padding to each item side. In the sample box list with 1 in of padding, a 3 x 3 x 4 candle jar gets the 6x4x4 box ($0.85), a 4.5 x 3.5 x 4 mug needs the 9x6x6 ($1.35), and a 24 x 3 x 3 poster tube shows No box fits, while the download's Status column flags 0, negative or text sizes as Check.

Sorting the sides lets an item fit turned any way. The formula checks straight fits only, not diagonal packing.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (28 workbooks, 14,364 cell comparisons, 0 mismatches), plus the edge cases below. The functions used also exist in Excel and Google Sheets, but this post was not tested in Excel or Google Sheets.

![Seven items with length, width, height, padding, chosen box size and box cost](images/2026-11-17-team-01-pick-smallest-shipping-box-that-fits-item-dimensions-match-index-excel-google-sheets/smallest-shipping-box-per-item-template.png) *Seven fictional items with 1 in of padding: the candle jar gets the 6x4x4 box at $0.85, the mug needs the 9x6x6 at $1.35, and the 24 in poster tube shows No box fits (the helper and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The formulas are typed the same way in Excel and Google Sheets. Measure the inside of each box, not the outside size on a supplier listing.

1. On the **Boxes** tab, type one box per row from row 5: name in **A**, inside sizes in inches in **B:D**, cost in **E** (optional).
2. List the boxes in the order you want them tried, smallest or cheapest first. The first box that fits in list order wins.
3. In **F5**, type `=MAX(B5:D5)` for the longest side. In **G5**, type `=SUM(B5:D5)-MAX(B5:D5)-MIN(B5:D5)` for the middle side. In **H5**, type `=MIN(B5:D5)` for the shortest side. Fill all three down.
4. On the **Orders** tab, type the padding you want on each side in **B1**. How much padding protects an item is your call.
5. Type each item in **A5** and its three sides in **B5:D5**, in any order.
6. In **F5**, type the lookup: `=INDEX(Boxes!$A$5:$A$24,MATCH(1,INDEX((Boxes!$F$5:$F$24>=MAX(B5:D5)+$B$1)*(Boxes!$G$5:$G$24>=SUM(B5:D5)-MAX(B5:D5)-MIN(B5:D5)+$B$1)*(Boxes!$H$5:$H$24>=MIN(B5:D5)+$B$1),0),0))`. Multiplying the three tests gives 1 only where all sides fit, and `MATCH(1,...,0)` finds the first 1.
7. Wrap it in `IFERROR(...,"No box fits")` so an item too big for every box says so. Fill it down.

The `MATCH(1,INDEX(...,0),0)` trick is the same one used to [auto-categorize expenses from bank descriptions](https://tidytabs.blogspot.com/2026/10/how-to-auto-categorize-expenses-from.html). Keep the inner `INDEX(...,0)`: in LibreOffice, the MATCH typed without it gave 0 instead of No box fits on the poster tube.

The template adds guards: a Status in **Orders E5** (the lookup runs only on OK rows), a Box row helper in **H** so the name and cost come from the same row, and sides rounded to 6 decimals so an exact fit cannot fail on a binary leftover such as 10.000000000000002.

Help pages: Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd).

## Example

The boxes, costs and orders are fictional. The box list, in the order tried: 6x4x4 $0.85, 8x6x4 $1.05, 10x8x4 $1.30, 9x6x6 $1.35, 12x9x6 $1.75, 12x12x8 $2.40, 16x12x8 $2.95, 18x14x10 $3.80. Padding is 1 in.

| Order or item | L x W x H | Status | Box to use | Box cost |
|---|---|---|---|---|
| Candle jar | 3 x 3 x 4 | OK | 6x4x4 | $0.85 |
| Mug | 4.5 x 3.5 x 4 | OK | 9x6x6 | $1.35 |
| Flat print | 11 x 8.5 x 0.25 | OK | 12x12x8 | $2.40 |
| Tray (typed on its side) | 5 x 9 x 3 | OK | 10x8x4 | $1.30 |
| Big quilt | 15 x 11 x 7 | OK | 16x12x8 | $2.95 |
| Poster tube | 24 x 3 x 3 | OK | No box fits | |
| Soap bar set | 5 x 3 x 3 | OK | 6x4x4 | $0.85 |

The mug padded to 5.5 x 5 x 4.5 is too tall for the three 4 in boxes. The flat print's padded middle side of 9.5 in fails the 12x9x6. The soap bar set padded to exactly 6 x 4 x 4 gets the 6x4x4, because `>=` keeps an exact fit.

The summary in **K4:K11** shows 7 items, 6 with a box, 1 No box fits, $9.70 total box cost and CHECK OK. With padding 0, the mug fits the 6x4x4; with padding 2, the candle jar moves up to the 9x6x6.

## Troubleshooting

### The box is bigger than needed
The sides were compared as typed, length to length. With padding 1, that sent the candle jar to the 9x6x6 and the tray to the 12x12x8. Sort both the box and the item into longest, middle and shortest first, and they get the 6x4x4 and the 10x8x4.

### Every item gets the same big box
The box list is not in the order you want them tried. With the 18x14x10 typed first, every item that fits went to it, and the boxes below showed Note: smaller than above in orange. Move the smallest or cheapest box to the top.

### A row shows Check: size
A side is 0, negative, blank or text such as "8 in". Type a plain number of inches.

### Every row shows Check: padding
The padding in **B1** is text such as "1 in" or a negative number. Type a plain number, or leave it blank to use 0.

### A box is named but the cost is blank
That box's cost is typed as text, such as "$2.95". The Boxes tab marks it Check: cost and CHECK shows Check; the box is still used. Retype the cost as a number.

## Template

[Download the .xlsx](templates/tidy-tabs-smallest-box-that-fits.xlsx)

The file has three tabs. **Boxes** holds rows 5 to 24: Box name, Inside length, Inside width, Inside height (in) and Cost each (inputs in **A:E**), then Longest side, Middle side, Shortest side, Volume and Status (formulas in **F:J**). A summary in **L4:M8** counts boxes and flags. **Orders** has the padding in **B1** and rows 5 to 104: Order or item, Length, Width and Height (inputs in **A:D**), then Status, Box to use, Box cost and the Box row helper (formulas in **E:H**), with the summary in **J4:K11**. **How to use** explains each column.

For carrier size limits and rates, see [dimensional weight and billable weight](https://tidytabs.blogspot.com/2026/10/how-to-calculate-dimensional-weight-and.html) and [shipping cost by weight](https://tidytabs.blogspot.com/2026/10/how-to-look-up-shipping-cost-by-weight.html).

Yellow cells are inputs and gray cells are formulas. Check turns red and No box fits turns orange. Both data tabs print Landscape on US Letter with the header row repeated. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
