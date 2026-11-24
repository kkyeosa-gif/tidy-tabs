---
title: How to Find the Cheapest Mix of Pack Sizes to Buy in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 96 workbooks and 1,923,648 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: List every mix of large and medium packs with INT and MOD, fill the rest with ROUNDUP, and take the lowest cost with MIN. No Solver needed.
image_prompts: Two plain brown shipping cases on a home candle workshop floor, one open and showing rows of empty glass candle jars in cardboard dividers, with a few loose single jars lined up on the workbench beside them, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Open shipping case of empty glass candle jars beside loose single jars
threads: 95 jars, and buying the biggest case first costs $116.80. Two 48-jar cases cost $104.00.\nList every mix of large and medium packs with INT and MOD, let ROUNDUP fill the rest with singles, and take =MIN(costs).\nNo Solver needed.
---
List every count of large packs times every count of medium packs with `=INT(mix/medium_options)` and `=MOD(mix,medium_options)`, let the smallest pack cover the rest with `=ROUNDUP((need-large_count*large-medium_count*medium)/small,0)`, and take the lowest total with `=MIN(costs)`. For 95 made-up candle jars sold as singles at $1.80, 12-cases at $15.00 and 48-cases at $52.00, the cheapest buy is two 48-cases for $104.00 with 1 jar extra, while filling with the biggest case first costs $116.80.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (96 workbooks, 1,923,648 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Jar pack prices beside 19 case mixes, the $104.00 two-case mix highlighted](images/2026-11-24-team-02-cheapest-mix-of-pack-sizes-to-buy-at-least-a-quantity-int-mod-roundup-min-excel-google-sheets/cheapest-pack-mix-template.png) *Order tab for 95 fictional jars beside the Mixes tab: two 48-jar cases for $104.00 (96 jars, 1 extra) is mix 18 and the cheapest, while buying the biggest case first costs $116.80, $12.80 more. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The sheet handles up to three pack sizes: it counts the large and medium packs, and the smallest fills the rest. No Solver, no macros.

The steps are the same in Excel and Google Sheets.

1. On the **Order** tab, type how many units you need in **B1**, such as 95.
2. Type up to three packs in rows 5 to 7: Pack name in **A**, Units per pack in **B**, Price per pack in **C**. Any order works. For only two packs, leave the third row blank.
3. The sheet sorts the packs into Large, Medium and Smallest in rows 11 to 13 with LARGE and SMALL.
4. Count the options for each pack in **D11** and **D12** with `=ROUNDUP($B$1/B11,0)+1`: for 95 jars, 3 counts of 48-cases (0 to 2) and 9 of 12-cases (0 to 8). With two packs, D12 is 1.
5. On the **Mixes** tab, number the mixes in **A6** from 0 down to D11 times D12 minus 1. For 95 jars that is 27 mixes, 0 to 26.
6. Turn each mix number into counts: Large packs in **B6** with `=INT($A6/Order!$D$12)` and Medium packs in **C6** with `=MOD($A6,Order!$D$12)`. Together they step through every pair, like an odometer whose last wheel has D12 positions. The columns show counts (0, 1, 2 and up), not Yes/No. See the [ROUNDUP](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) and [MOD](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) help from Microsoft.
7. Fill the rest in **D6** with `=MAX(0,ROUNDUP((Order!$B$1-B6*Order!$B$11-C6*Order!$B$12)/Order!$B$13,0))`. MAX keeps the count at 0 when the cases already cover the need.
8. Price the mix in **E6** with `=ROUND(B6*Order!$C$11+C6*Order!$C$12+D6*Order!$C$13,2)`. **F6** adds up the units bought and **G6** the packs in all.
9. Take the lowest cost in **E4** with `=MIN(E6:E2005)`. **Order H11** shows it, and **H12:H15** read the counts of the chosen mix back with INDEX.

**Ties:** when two mixes cost the same, the sheet picks the one with the fewest units bought (fewest extra), then the fewest packs, then the first row (fewer large packs). Columns **H** to **J** on Mixes apply it, and **J** marks the chosen mix with 1.

**Biggest-first:** **E11:E13** on Order do what most people do by hand: as many large packs as fit with `=INT($B$1/B11)`, then as many medium packs as fit, then the smallest pack for the rest. **H18** prices that order and **H19** shows how much more it costs.

## Example

The supplier, packs and prices are made up. Packs: 48-jar case $52.00, 12-jar case $15.00, Single jar $1.80, which works out to $1.08, $1.25 and $1.80 per jar.

For 95 jars, the Order result shows a cheapest total of $104.00 (2 large, 0 medium, 0 smallest packs), 96 units bought, 1 extra, $1.09 per jar needed, and mix 18 chosen. Biggest-first totals $116.80, so it costs $12.80 more.

A few rows from the Mixes tab for the same order:

| Mix # | Large | Medium | Smallest | Cost | Units |
|---|---|---|---|---|---|
| 0 | 0 | 0 | 95 | $171.00 | 95 |
| 8 | 0 | 8 | 0 | $120.00 | 96 |
| 12 | 1 | 3 | 11 | $116.80 | 95 |
| 13 | 1 | 4 | 0 | $112.00 | 96 |
| 18 | 2 | 0 | 0 | $104.00 | 96 |

Mix 12 is the biggest-first order; mix 18 is chosen.

Other quantities with the same packs:

| Jars needed | Cheapest mix | Cost | Extra | Biggest-first | Biggest-first cost |
|---|---|---|---|---|---|
| 95 | 2 x 48 | $104.00 | 1 | 1 x 48 + 3 x 12 + 11 singles | $116.80 |
| 130 | 2 x 48 + 3 x 12 | $149.00 | 2 | 2 x 48 + 2 x 12 + 10 singles | $152.00 |
| 40 | 1 x 48 | $52.00 | 8 | 3 x 12 + 4 singles | $52.20 |
| 10 | 1 x 12 | $15.00 | 2 | 10 singles | $18.00 |
| 7 | 7 singles | $12.60 | 0 | same | $12.60 |

For 10 jars, one 12-case beats 10 singles even with 2 jars left over. For 7 jars, biggest-first is already the cheapest.

## Troubleshooting

### The order I worked out by hand costs more
Biggest-first is not always cheapest: $116.80 against $104.00 for 95 jars. Compare **H18** with **H11** before you order.

### No result, and CHECK says Check: too many mixes
The need is large compared with the large and medium packs. 3,000 jars with these packs would take 16,064 mixes, past the 2,000 limit, so the sheet leaves the result blank instead of giving a wrong answer. Try larger packs in rows 5 to 7, or work out part of the order at a time.

### No result, and a pack row says Check
Read the Status in column **E**. "Same units as another pack" means two rows share a size; "units must be a whole number" catches entries like 1.5; "units or price blank or text" catches a price typed as text. Fix the yellow cell.

### Two mixes cost the same and the sheet picked the one with fewer packs
That is the tie rule. In a test with 24 needed and packs of 24 at $24.00, 12 at $12.00 and 1 at $1.00, every exact mix cost $24.00 and the sheet chose one 24-case. A plain MATCH on the lowest cost would take the first row instead: 24 singles.

## Template

[Download the .xlsx](templates/tidy-tabs-cheapest-pack-mix.xlsx)

The file has three tabs. **Order** holds the units needed in **B1**, up to three packs in **A5:C7**, the sorted packs and biggest-first counts in rows 11 to 13, a summary with a CHECK cell in **G4:H9** and the result in **G10:H20**. **Mixes** lists every mix in rows 6 to 2005 and prints its first 27 mixes on one page. **How to use** explains each formula. Order and Mixes print Landscape on US Letter.

Yellow cells are inputs and gray cells are formulas. Every status starting with Check turns red, and the chosen mix and cheapest total turn light green.

Minimum orders, shipping and whole-order discounts are not part of this sheet. For a case-size minimum, see [How to Round an Order Quantity Up to a Case Size and Minimum Order](https://tidytabs.blogspot.com/2026/11/how-to-round-order-quantity-up-to-case.html).

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
