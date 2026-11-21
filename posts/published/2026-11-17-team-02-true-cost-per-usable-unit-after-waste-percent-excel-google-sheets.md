---
threads_url: https://www.threads.com/@tin_ylab/post/Df_AYkzEVxy
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-find-true-cost-per-usable-unit.html
title: How to Find the True Cost per Usable Unit After Waste in Excel and Google Sheets
labels: excel-formulas, small-business, pricing, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 15 workbooks and 9,885 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Cost per unit with a waste percentage: =pack_cost/(units_per_pack*(1-waste)). Why adding waste on top comes out low, plus a waste log. Free template.
image_prompts: A wooden florist workbench in a small flower studio with a galvanized bucket of fresh garden roses, a pile of trimmed stems, stripped leaves and two broken rose heads swept to one side, and a pair of pruning shears resting on the bench, soft morning light, realistic photo, no screens, no readable text, no logos
image_alt: Bucket of garden roses beside a pile of trimmed stems and pruning shears
threads: 25 roses for $30.00 and 20% of them get tossed: $1.50 per usable stem.\nAdding 20% to the $1.20 bought cost says $1.44, 6 cents low on every stem.\nDivide instead: =pack_cost/(units_per_pack*(1-waste))
---
Divide the pack cost by the units you can actually use with `=pack_cost/(units_per_pack*(1-waste))`, where waste is your own percent such as 20%. In the made-up sample, $30.00 for 25 garden roses at 20% waste is $1.50 per usable stem, while the common shortcut of adding 20% to the $1.20 bought cost gives $1.44, 6 cents low on every stem; the download shows Check if waste is typed as 20 instead of 20%.

The waste percent comes from your own counts: stems tossed, fabric lost to cutting, board ends, roast loss. The template has a log that measures it.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (15 workbooks, 9,885 cell comparisons, 0 mismatches), plus the edge cases below. Not tested in Excel or Google Sheets: neither app was opened for this post.

![Five materials with pack cost, waste percent, bought cost, usable cost and shortcut cost](images/2026-11-17-team-02-true-cost-per-usable-unit-after-waste-percent-excel-google-sheets/cost-per-usable-unit-after-waste-template.png) *Five fictional materials: 25 garden roses for $30.00 at 20% waste cost $1.20 each as bought, $1.50 per usable stem, and $1.44 with the add 20% shortcut (the unit, helper and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Materials** tab, one material per row starting in row 5.

1. Type the material in **A5** and its unit (stems, yd, ft, lb) in **B5**.
2. Type the pack cost in **C5** ($30.00) and the units in one pack in **D5** (25).
3. Type your waste percent in **E5** as a percent: 20%, not 20.
4. In **F5**, type `=C5/D5` for the cost per bought unit ($1.20).
5. In **G5**, type `=D5*(1-E5)` for the usable units per pack (20 stems).
6. In **H5**, type `=C5/(D5*(1-E5))` for the cost per usable unit ($1.50). You pay for the whole pack but use only what is left.
7. Optional: in **J5**, type the usable units you need (200), and in **K5**, `=ROUNDUP(J5/(D5*(1-E5)),0)` for whole packs to buy (10). For case sizes and minimum orders, see [how to round an order quantity up to a case size](https://tidytabs.blogspot.com/2026/11/how-to-round-order-quantity-up-to-case.html).
8. Fill the row down for each material.

To measure the waste instead of guessing it, use the **Waste log** tab:

1. Log each pack or batch: date in **A** (MM/DD/YYYY, such as 11/02/2026), material in **B**, units bought in **C**, usable units in **D**.
2. Type each material name once in **G5:G24**.
3. **K5** works like `=1-SUMIF(B:B,G5,D:D)/SUMIF(B:B,G5,C:C)`: 1 minus total usable over total bought. The template uses SUMIFS to count only OK rows.
4. On the Materials tab, link the waste cell to the log, for example `='Waste log'!K5`, as in sample row 9.

The template wraps these in checks. **L5** shows Check: cost or units (blank, text, 0 or negative) or Check: waste (blank, below 0%, 100% or more, or 20 instead of 20%). Costs stay blank on a Check row, so 100% waste never divides by zero.

Help pages: Microsoft's [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7).

## Example

The materials, prices and waste percents are fictional, not typical figures.

| Material | Pack cost | Units per pack | Waste % | Cost per bought unit | Cost per usable unit | Shortcut | Usable units needed | Packs to buy |
|---|---|---|---|---|---|---|---|---|
| Garden roses (stems) | $30.00 | 25 | 20.0% | $1.20 | $1.50 | $1.44 | 200 | 10 |
| Quilting cotton (yd) | $9.50 | 1 | 15.0% | $9.50 | $11.1765 | $10.925 | 30 | 36 |
| Pine board (ft) | $14.40 | 8 | 12.5% | $1.80 | $2.0571 | $2.025 | 60 | 9 |
| Green coffee (lb) | $6.75 | 1 | 18.0% | $6.75 | $8.2317 | $7.965 | 40 | 49 |
| Garden roses (waste from log) | $30.00 | 25 | 18.0% | $1.20 | $1.4634 | $1.416 | 200 | 10 |

The cost columns show up to four decimals. The shortcut is low on every row. Column M shows what the packs give: 30.6 usable yd of cotton, 63 ft of pine, 40.18 lb of roasted coffee.

The Waste log holds three rose bunches (25 to 21, 25 to 19, 50 to 42): 100 bought, 82 usable, 18.00% measured waste. Averaging each bunch's own percent (16%, 24%, 16%) gives 18.67% instead, because it weights a 25-stem bunch the same as a 50-stem one. Two cotton rows measure 14.06%; pine board has no rows, so it stays blank.

## Troubleshooting

### The cost per usable unit looks too low
You used the shortcut, cost x (1 + waste). For the roses it gives $1.44 instead of $1.50. Divide by the usable units with `=C5/(D5*(1-E5))` instead.

### Status says Check: waste
The waste was typed as 20 instead of 20%, left blank, is negative, or is 100% or more. A typed 20 would mean 25 x (1 - 20) = -475 usable stems. Type the percent with the % sign, or link the cell to the Waste log.

### The measured waste ignores some log rows
The name in column B does not match the name in G. In LibreOffice, a row typed "garden roses " with a trailing space was not counted, while "GARDEN ROSES" in capitals was. **H27** counts log rows whose name is not in G. [Cleaning up a list with TRIM and PROPER](https://tidytabs.blogspot.com/2026/10/how-to-clean-up-customer-list-with-trim.html) fixes stray spaces. A row with more usable than bought shows Check: usable vs bought and is left out too.

### Packs to buy shows Check
The usable units needed in J is text (such as "200" typed as text) or negative. Type a plain number, or leave J blank; the costs still show.

## Template

[Download the .xlsx](templates/tidy-tabs-cost-per-usable-unit-after-waste.xlsx)

The file has three tabs. **Materials** holds rows 5 to 54 with the inputs above, the cost columns, the Shortcut for comparison, Packs to buy, Status and Usable units from those packs. The summary in **O4:P8** counts materials, Check rows and entries below row 54, with a CHECK.

**Waste log** holds rows 5 to 204 (Date, Material, Bought, Usable, Status) and a table in **G4:L24** with rows logged, bought, usable and measured waste % per material. **G26:H29** counts Check rows, names not in G and entries below row 204, with a CHECK. **How to use** explains each column and Status word.

The sheet finds your material cost, not a price; for pricing, see [markup and profit margin](https://tidytabs.blogspot.com/2026/10/how-to-calculate-markup-and-profit.html).

Yellow cells are inputs and gray cells are formulas. A Status starting with Check turns red, and so does a CHECK cell that is not OK. The data tabs print Landscape on US Letter with the header row repeated. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
