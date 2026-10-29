---
threads_url: https://www.threads.com/@tin_ylab/post/DfFSOr7EbVH
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-fix-total-that-is-off-by-one.html
title: How to Fix a Total That Is Off by One Cent in Excel and Google Sheets
labels: excel-formulas, small-business, rounding
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model that uses Decimal with half-up rounding. Checked: the 8-line sample, all whole cents, half-cent prices with odd quantities, 20 lines of 1 x $0.125, a large quantity, negative quantities, a zero quantity, one line only, random 3 and 4 decimal prices, a blank row, and a price typed as text. Excel and Google Sheets were not used.
search_description: Fix a total that is off by one cent in Excel and Google Sheets with ROUND. Free invoice template with unrounded and rounded totals side by side, the difference and a check.
image_prompts: A glass jar of soy wax flakes, a spool of cotton wick and two small round tin lids on a worn wooden craft table, soft window light, no screens, no readable text, no logos
image_alt: Jar of wax flakes, a wick spool and two tin lids on a wooden table
threads: Printed invoice lines add up to one cent less than the total?\nThe cells hold more decimals than they show. Round each line first: =ROUND(B5*C5,2), then SUM the rounded column.\nNow the printed lines match the printed total, and the template shows both totals side by side.
---
Round each line to cents with `=ROUND(B5*C5,2)`, where B5 is the quantity and C5 the unit price, then add the rounded lines with `=SUM(E5:E12)`. The printed lines then add up to the printed total.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. ROUND and SUM are old functions that Excel and Google Sheets also have, but neither app was opened for this post.

![Invoice sheet listing eight craft supply lines with unrounded and rounded amounts and two totals](images/2026-10-29-team-02-totals-off-by-one-cent-round-each-line-excel-google-sheets/round-each-line-invoice-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The cause is hidden decimals. A cell formatted to 2 decimals still holds the full value, so 7 x $1.145 is stored as 8.015 and shows as $8.02, but a plain SUM adds 8.015. Rounding each line makes the cell hold the cents you see.

The layout below matches the template: item names in column A, quantity in B, unit price in C, starting in row 5.

### In Excel

1. In **A5**, **B5** and **C5**, type an item, a quantity and a unit price, such as Cotton wick, 7 and 1.145. Add more lines below.
2. Select **C5:C12**, press **Ctrl+1** and set **Number** to 4 decimal places so fractions of a cent stay visible.
3. In **D5**, type `=B5*C5`. This is the line before rounding.
4. In **E5**, type `=ROUND(D5,2)`. Or skip D and type `=ROUND(B5*C5,2)` in E5.
5. Select **D5:E5** and drag the fill handle down to your last line.
6. Under the lines, type `=SUM(E5:E12)` for the total. Adjust the range to your last row.

Microsoft's page: [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c). It describes the function, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 6 above, typed exactly as written.
2. To show 4 decimals in column C, select it and use **Format > Number > Custom number format**, then type `0.0000`.

Google's help page: [ROUND function](https://support.google.com/docs/answer/3093440). That page describes the function, not our test.

### The version in the template

The template keeps both versions side by side so you can see the gap. Row 5 uses:

- **D5** `=IF(AND(ISNUMBER(B5),ISNUMBER(C5)),B5*C5,"")`, the line before rounding.
- **E5** `=IF(D5="","",ROUND(D5,2))`, the line rounded to cents.
- **F5** `=IF(D5="","",ROUND(E5-D5,4))`, how much rounding moved the line.

The summary sits in **J5:J9**. **J5** is `=SUM(D5:D24)`, **J6** is `=SUM(E5:E24)`, and **J7** is `=ROUND(ROUND(J5,2)-J6,2)`. **J9** says OK when the printed lines add up to the printed total, or CHECK.

## Example

Sample data is fictional: eight made up candle supply lines with unit prices that carry fractions of a cent.

| Item | Qty | Unit price | Unrounded | Rounded |
|---|---|---|---|---|
| Soy wax flakes (oz) | 3 | $0.3350 | 1.005 | $1.01 |
| Cotton wick | 7 | $1.1450 | 8.015 | $8.02 |
| Paper label | 12 | $0.0425 | 0.51 | $0.51 |
| Tin lid | 5 | $2.3490 | 11.745 | $11.75 |
| Fragrance oil (ml) | 9 | $0.8750 | 7.875 | $7.88 |
| Gift tag | 24 | $0.1675 | 4.02 | $4.02 |
| Kraft sleeve | 6 | $3.4575 | 20.745 | $20.75 |
| Twine (ft) | 15 | $0.0850 | 1.275 | $1.28 |

**J5**, the plain total of the unrounded lines, is $55.19. **J6**, the total of the rounded lines, is $55.22. **J7**, the difference, is -$0.03.

The printed lines add to $55.22, so a total of $55.19 would look 3 cents short. Six of the eight lines sit exactly on a half cent. In LibreOffice all six rounded up, and they matched the Python half-up model. That is the result of our check, not a statement about Excel or Google Sheets.

Other cases the team checked in LibreOffice:

- All whole cents: both totals matched and the difference was $0.00.
- 20 lines of 1 x $0.125: $2.50 unrounded, $2.60 rounded, difference -$0.10.
- 1,000,000 x $0.335 and similar large lines: the totals differed by 0.0025, which shows as $0.00 in J7 because J7 is shown to cents.
- One line of 3 x $0.335: $1.005 shows as $1.01 too, so both totals looked equal.
- A return typed as a negative quantity: -3 x $0.335 rounded to -$1.01.

This post does not say which method a customer or authority requires. It only shows how to make the printed page add up.

## Troubleshooting

### The rounded total still looks off by a cent
Check whether the line amounts are rounded or only formatted. Click a line cell. If the formula bar shows `=B5*C5` with no ROUND, the cell holds hidden decimals. Wrap it in `=ROUND(...,2)` and sum those cells.

### J9 says CHECK
A unit price or quantity is probably text, so that row has no line. Text left-aligns. Retype it as a number. A quantity with no price also gives CHECK.

### Two totals are the same but the lines still hide fractions
With one line, or with whole-cent prices, the unrounded and rounded totals can match, so J7 is $0.00. That is expected. Rounding only changes a total when a line has a fraction of a cent.

### Someone suggests Precision as displayed
Excel has a **Set precision as displayed** option that changes stored values to what is shown. We did not test it, and the template's How to use sheet lists it as a risky option. Rounding each line with ROUND changes only the cells you wrap, so it is easier to check.

## Template

[Download the .xlsx](templates/tidy-tabs-totals-off-by-a-cent.xlsx)

The file has two tabs. **Invoice** holds yellow inputs in **A5:C24** (item, quantity, unit price) for up to 20 lines, with 8 fictional rows filled in. Gray cells are formulas in **D** to **G** and **J5:J9**. **How to use** has short notes.

The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Invoice tab. It gives no tax, invoice or accounting advice. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
