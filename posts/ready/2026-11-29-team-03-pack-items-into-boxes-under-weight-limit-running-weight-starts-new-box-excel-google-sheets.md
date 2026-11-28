---
title: How to Split Items Into Boxes Under a Weight Limit in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and every formula cell was compared with an independent Python model, 645 workbooks (45 edge cases plus 600 random lists) and 2,011,755 cell comparisons, 0 mismatches, 0 error values. The heaviest-first result (5 boxes to 4) and the same boxes without ROUND were seen in LibreOffice Calc only, on scratch copies. Not tested in Excel or Google Sheets.
search_description: Give each item a box number with a running weight that starts a new box when the next item would pass your weight limit, plus each box's packed weight.
image_prompts: A shipping table in a small home goods shop with a postal scale holding a half-packed cardboard box of tissue-wrapped jars and mugs, a heavy cast iron doorstop sitting alone beside the scale, and a roll of packing tape on a dispenser, soft overhead daylight, realistic photo, no screens, no readable text, no logos
image_alt: Half-packed carton on a postal scale beside a cast iron doorstop and tape dispenser
threads: Start a new box when the next item would pass the limit: =IF(ROUND(D6+B7,4)>$B$3,C6+1,C6)\nThe test is >, not >=, so a box that lands exactly on the limit keeps the item.\nIn the sample, 10 items pack into 5 boxes, and box 1 weighs exactly 20.00 lb packed.
---
Give the first item box 1, then for each next item use `=IF(ROUND(D6+B7,4)>$B$3,C6+1,C6)`: if the running weight plus this item would pass the usable weight in B3, it starts a new box, otherwise it joins the current one. With a 20 lb limit and 1.5 lb of box and fill, the made-up December wholesale order of 10 items packs into 5 boxes, and the first one weighs exactly 20.00 lb packed.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (645 workbooks, 2,011,755 cell comparisons, 0 mismatches, 0 error values), including the edge cases below. The heaviest-first result and the run without ROUND were seen in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Five packed boxes with weights, room left, an at-the-limit box and one overweight doorstop box](images/2026-11-29-team-03-pack-items-into-boxes-under-weight-limit-running-weight-starts-new-box-excel-google-sheets/pack-boxes-by-weight-template.png) *Boxes tab of a made-up December wholesale order of 10 items with a 20 lb limit and 1.5 lb of box and fill: per box the item count, item and packed weight, room left, status (box 1 at the limit in green, the cast iron doorstop box over the limit in red) and first and last item, with the summary of 5 boxes, 62.10 lb of items and a weight-only minimum of 4. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

This packs items in the order you list them (next-fit): it fills a box until the next item would go over, then opens the next one. It does not look for the fewest possible boxes, and the limit is your own. Carrier limits, rates, box sizes and fragile items are not covered; for billable weight, see [dimensional weight](https://tidytabs.blogspot.com/2026/10/how-to-calculate-dimensional-weight-and.html).

The steps are the same in Excel and Google Sheets.

1. On the **Packing** tab, type the most one packed box may weigh in **B1** (for example 20) and the weight of the empty box plus tape and fill in **B2** (for example 1.5).
2. **B3** works out the weight left for items with `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),ROUND(B1-B2,4),"")`. In the sample that is 18.50 lb.
3. From row 6, type one item per row in packing order: the name in **A** and the weight in pounds as a decimal in **B** (8 oz is 0.5). For pounds and ounces, convert first with [ounces to pounds](https://tidytabs.blogspot.com/2026/10/how-to-convert-ounces-to-pounds-and.html).
4. Column **C** is the box number. **C6** is the number 1. **C7** and down use `=IF(LEFT(E7,6)="Packed",IF(AND(D6>0,ROUND(D6+B7,4)>$B$3),C6+1,C6),C6)`. The extra parts skip rows that are not packed and keep a too-heavy first item from leaving an empty box.
5. Column **D** is the weight in the box so far, a running total that starts over in each new box. **D6** is `=IF(LEFT(E6,6)="Packed",B6,0)` and **D7** is `=IF(LEFT(E7,6)="Packed",IF(C7=C6,ROUND(D6+B7,4),B7),D6)`: same box, add the item; new box, start at the item's weight.
6. Check column **E** (Status). It shows Packed, "Packed alone: too heavy for one box" or a Check message. Check rows are not packed.
7. Open the **Boxes** tab. Each box gets one row: item count, item weight (SUMIFS of the Packed rows with that box number), packed weight (items plus **B2**), room left, a Status of OK, At the limit or Over the limit, and the first and last item.
8. Read the summary in **J4:K13**. **Weight-only minimum** is the total item weight divided by B3, rounded up: the count if items could be split. **CHECK** in **K13** is OK only when the Packing CHECK is OK, no more than 200 boxes are needed and the box totals match the item total.

A box that comes out at exactly B3 keeps the item, because the test is "more than", not "at least". For the functions, see the ROUND help from [Microsoft](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) or [Google](https://support.google.com/docs/answer/3093440), and the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).

Related: [how many boxes you need](https://tidytabs.blogspot.com/2026/10/how-to-calculate-how-many-boxes-you.html), [split a list into batches of 10](https://tidytabs.blogspot.com/2026/10/how-to-split-list-into-batches-of-10.html) and a [running total that restarts each month](https://tidytabs.blogspot.com/2026/11/how-to-make-running-total-that-restarts.html).

## Example

The order is made up: a maker's December wholesale order of 10 pieces, 62.10 lb in all, with a 20.00 lb limit per box and 1.50 lb of box and packing, so 18.50 lb is usable for items.

| Item | Weight (lb) | Box | Weight in box so far (lb) | Status |
|---|---|---|---|---|
| Mug set (4) | 6.40 | 1 | 6.40 | Packed |
| Cast iron trivet | 9.80 | 1 | 16.20 | Packed |
| Candle 3 pack | 2.30 | 1 | 18.50 | Packed |
| Soap crate (12 bars) | 7.40 | 2 | 7.40 | Packed |
| Wool throw | 3.20 | 2 | 10.60 | Packed |
| Ceramic vase | 5.90 | 2 | 16.50 | Packed |
| Jam gift box | 4.40 | 3 | 4.40 | Packed |
| Cast iron doorstop | 19.00 | 4 | 19.00 | Packed alone: too heavy for one box |
| Card bundle (25) | 1.10 | 5 | 1.10 | Packed |
| Tea tin set | 2.60 | 5 | 3.70 | Packed |

The Boxes tab turns that into a packing list:

| Box | Items | Item weight (lb) | Packed weight (lb) | Status |
|---|---|---|---|---|
| 1 | 3 | 18.50 | 20.00 | At the limit |
| 2 | 3 | 16.50 | 18.00 | OK |
| 3 | 1 | 4.40 | 5.90 | OK |
| 4 | 1 | 19.00 | 20.50 | Over the limit: one item too heavy |
| 5 | 2 | 3.70 | 5.20 | OK |

Box 1 lands at exactly 20.00 lb, so the candles stay in. The soap crate would take box 1 to 25.90 lb of items, so it starts box 2; the jam gift box would take box 2 to 20.90 lb, so it starts box 3. The 19.00 lb doorstop is over 18.50 lb on its own, so it gets box 4 and a flag. Total packed weight is 69.60 lb.

Dividing 62.10 lb by 18.50 lb gives a weight-only minimum of 4 boxes, but in this list order the items need 5. In a LibreOffice scratch copy with the same rows typed heaviest first, the order packed into 4 boxes (68.10 lb packed): the doorstop alone, then 9.80 + 7.40, then 6.40 + 5.90 + 4.40, then the four lightest. Heaviest first often helps, but it does not promise the fewest boxes.

## Troubleshooting

### A row says "Check: blank row inside the list (skipped)"
An empty row sits between items. Packing carries on past it (its Box and running weight repeat the row above in light gray), but **CHECK** shows Check until you delete it.

### A row says "Check: weight blank or text"
The weight is blank or typed with a unit, such as "3.2 lb". Type the number only (3.2). Weights of 0 or below show "Check: weight 0 or below".

### Every row says "Check: settings (B1:B2)" and there are 0 boxes
**B1** is not above 0, **B2** is below 0 or at least **B1**, or one of them is text or blank. Fix the settings and the rows pack again.

### One box is over the limit
It holds one item heavier than **B3** ("Packed alone: too heavy for one box"). Raise your limit or ship that item on its own.

### Does the running weight need ROUND?
In LibreOffice Calc, removing ROUND from the box and running-weight formulas gave the same boxes in all 45 edge cases, including 0.1 + 0.2 against 0.3 lb usable and 25 items of 0.1 lb against 1.0 lb. That is a LibreOffice result only; the template keeps `ROUND(...,4)` as cheap insurance.

## Template

[Download the .xlsx](templates/tidy-tabs-pack-boxes-by-weight.xlsx)

**Packing** holds the limit and box weight in **B1:B2**, the usable weight in **B3**, up to 500 items (rows 6 to 505) with Box, running weight and Status, and its own CHECK in **G5:H12**, which also catches anything typed below row 505. **Boxes** lists up to 200 boxes (rows 5 to 204) with weights, room left, status and first and last item, plus the summary in **J4:K13**. **How to use** explains each formula. Packing and Boxes print Landscape on US Letter with the header row repeated.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
