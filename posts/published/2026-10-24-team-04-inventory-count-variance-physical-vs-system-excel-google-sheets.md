---
threads_url: https://www.threads.com/@tin_ylab/post/De6rNUDG4A6
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-compare-physical-inventory-count.html
title: How to Compare a Physical Inventory Count With Expected Stock in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 25 edited variants (all items counted, none counted, a typed 0 against a blank, expected 0 with a count, negative or text counts, text or blank cost, fractional counts, half cent products, numbers in a row with no item name, a row 16 added, rows 100 and 101). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Compare a physical inventory count with expected stock in Excel and Google Sheets. Free template that flags over and short items and values the gap at cost.
image_prompts: A wooden shelf of ceramic mugs and lidded soy candle jars next to an open cardboard stock box and a clipboard holding a blank sheet, soft overhead shop light, no screens, no readable text, no logos
image_alt: Shelf of ceramic mugs and candle jars beside an open stock box and clipboard
threads: Counted the shelf and it does not match the books?\nIn E2 use =IF(C2="","",C2-B2). Negative is short, positive is over, and a blank count stays blank.\nIn the sample, 13 of 14 items are counted and the gap is worth -$82.15 at cost. The shop is made up.
---
Subtract the expected quantity from the counted quantity with `=IF(C2="","",C2-B2)`, where B2 is the expected count and C2 is what you counted. A negative result means that item is short, a positive result means it is over, and a blank C2 stays blank.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Count sheet of 14 gift shop items with red short and blue over rows](images/2026-10-24-team-04-inventory-count-variance-physical-vs-system-excel-google-sheets/inventory-count-variance-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: one row per item. Type what the books say you should have, type what you counted, and let the sheet subtract.

1. In **A1:D1**, type the headings Item, Expected quantity, Counted quantity and Unit cost.
2. In **A2:A15**, type one item per row.
3. In **B2:B15**, type the expected quantity, the number your records say is on the shelf.
4. In **D2:D15**, type what each unit cost you.
5. Count the shelf. In **C2:C15**, type the count for each item. Leave C blank for an item you have not counted yet.
6. In **E2**, type `=IF(C2="","",C2-B2)` and fill it down.
7. In **F2**, type `=ROUND(E2*D2,2)` and fill it down. This values the difference at cost.
8. In **G2**, type a formula that returns Short when E2 is below 0, Match when it is 0, Over when it is above 0 and Not counted when C2 is blank, and fill it down.

How the difference formula reads:

- `C2=""` is true only when the counted cell is empty, so an uncounted item shows nothing instead of looking short.
- `C2-B2` is counted minus expected. Short is negative and over is positive.
- A 0 you type is a count of zero, not a blank. A missing mug that you counted as 0 is short by the full expected amount.

Both apps document COUNTIF, which the summary uses: [COUNTIF in Excel](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [COUNTIF in Google Sheets](https://support.google.com/docs/answer/3093480). Those pages describe the function, not our test.

### What the template adds

The built formulas are longer than the bare ones. **E2** also checks that both B2 and C2 are numbers and shows "Check count" if either is text. **F2** shows "Check cost" when the unit cost is blank or text. The core of each is the formula above.

**G2** reads the sign of E2 and returns Short, Match, Over, Not counted or Check count. **H2** is `=ABS(E2)`, the units off for that row, so the summary can add them without an array formula.

The summary sits in **J2:K9**:

- **K2** counts the items you counted: `=COUNT(C2:C100)`.
- **K3** counts the Not counted rows.
- **K4** is the number of items off: `=COUNTIF(E2:E100,">0")+COUNTIF(E2:E100,"<0")`.
- **K5** is units off: `=SUM(H2:H100)`.
- **K6** is the net value, **K7** the shortages and **K8** the overages, all at cost.
- **K9** says OK or CHECK.

### In Excel

The formulas work as typed. IF, ROUND, COUNT, COUNTIF, SUM, SUMIF and ABS are old functions, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed.

## Example

Sample data is fictional: a small gift shop with 14 items and made-up costs.

| Item | Expected | Counted | Unit cost | Difference | Value | Status |
|---|---|---|---|---|---|---|
| Ceramic mug | 48 | 45 | $6.50 | -3 | -$19.50 | Short |
| Soy candle 8 oz | 30 | 30 | $4.25 | 0 | $0.00 | Match |
| Canvas tote bag | 12 | 14 | $5.25 | 2 | $10.50 | Over |
| Sticker sheet | 200 | 190 | $0.20 | -10 | -$2.00 | Short |
| Ceramic trinket dish | 6 | 0 | $8.75 | -6 | -$52.50 | Short |
| Postcard set | 100 | | $0.30 | | | Not counted |

Ceramic mug: 45 - 48 = -3, and -3 x $6.50 = -$19.50. The Postcard set has no count yet, so E, F and H stay blank and the item is not in any total.

For all 14 items, the summary shows 13 counted, 1 not counted, 10 items off and 34 units off. Shortages come to -$102.30, overages to $20.15 and the net to -$82.15. K9 says OK.

This is arithmetic on the numbers you type. It does not tell you why a count is off or what to do about it, and it is not an accounting or tax treatment.

## Troubleshooting

### An item I have not counted shows as short
It should not. A blank counted cell stays blank and the status says Not counted. If you typed 0, the sheet treats that as a real count of zero and shows the full expected quantity as short. Clear the cell instead.

### The difference says Check count
Either the expected or the counted cell holds text, such as "n/a", or the expected cell is blank while a count is typed. Retype both as plain numbers. K9 shows CHECK until you do.

### The value says Check cost
The unit cost in column D is blank or text on a row that has a count. Type the cost as a number. A cost of 0 is allowed and gives $0.00.

### K9 says CHECK
The check fails when a quantity or cost is text, a number is negative, a named item has no expected quantity or cost, or numbers sit in a row with no item name. The template reads rows 2 to 100, so keep items inside that range.

## Template

[Download the .xlsx](templates/tidy-tabs-inventory-count-variance.xlsx)

The file has two tabs. **Count** holds the yellow inputs: item, expected quantity, counted quantity and unit cost in **A2:D100**. Gray cells are formulas: difference in **E**, value in **F**, status in **G**, units off in **H**, and the summary with the check in **J2:K9**. Short rows turn red, Match rows green and Over rows blue, and K9 turns red when it is not OK. **How to use** has short notes.

The items and costs are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, and the row colors were not seen in a spreadsheet window. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
