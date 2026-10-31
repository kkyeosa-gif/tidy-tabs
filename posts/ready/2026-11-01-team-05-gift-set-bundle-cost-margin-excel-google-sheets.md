---
title: How to Work Out the Cost and Margin of a Gift Set in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (Decimal, round half up) of every line cost, set cost, profit, margin and the check cell. Checked: a misspelled part name, a quantity of 0, a quantity typed as text, a part listed twice on the Parts tab, a typo in a set name and a blank price. The PDF export pages are US Letter. Excel and Google Sheets were not used.
search_description: Work out the cost and margin of a gift set from its parts with INDEX, MATCH and SUMIF in Excel and Google Sheets. Free bundle cost template with a check cell.
image_prompts: A kraft gift box open on a wooden table with a bar of soap, a small candle and a spool of ribbon beside it, a pencil resting on a blank paper notepad, soft window light, no screens, no readable text, no logos
image_alt: Open kraft gift box with a soap bar, small candle and ribbon spool next to a blank notepad
threads: Still costing out gift sets by hand?\nGive each set one row per part and pull each cost from one price list with INDEX and MATCH.\nSUMIF totals the set, then =(H5-G5)/H5 gives the margin.
---
Keep one price list of parts, then give each set one row per part. Look up each part's cost with `=ROUND(C5*INDEX(Parts!$B$5:$B$30,MATCH(B5,Parts!$A$5:$A$30,0)),2)`, total a set with `=SUMIF($A$5:$A$54,F5,$D$5:$D$54)`, and get the margin with `=(H5-G5)/H5`.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. INDEX, MATCH, SUMIF and IFERROR are in every current version of Excel and Google Sheets, but neither app was opened for this post.

![Gift set sheet with part line costs beside each set's total cost, price and margin](images/2026-11-01-team-05-gift-set-bundle-cost-margin-excel-google-sheets/gift-set-cost-margin-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Parts** lists each part once, with the name in column A and the unit cost in column B, starting in row 5. **Sets** has one line per part in a set (set name in A, part in B, quantity in C), and a short list of sets on the right (name in F, price in H).

1. On **Parts**, type each part once in **A5:B30**, such as Gift box and $0.95.
2. On **Sets**, type one line per part in **A5:C54**: the set name, the part spelled exactly as on **Parts**, and the quantity.
3. In **D5**, type the line cost formula below, then drag the fill handle down to row 54.
4. In **F5:F10**, type each set name once. In **H5:H10**, type its selling price.
5. In **G5**, type `=SUMIF($A$5:$A$54,F5,$D$5:$D$54)` and fill down to row 10. This adds every line cost that has the same set name.
6. In **I5**, type `=(H5-G5)/H5`. Format the column as a percent with one decimal.
7. In **J5**, type `=H5-G5` for the profit in dollars. Fill both columns down.

The line cost in **D5**:

`=IF(B5="","",IF(NOT(ISNUMBER(C5)),"Check qty",IFERROR(ROUND(C5*INDEX(Parts!$B$5:$B$30,MATCH(B5,Parts!$A$5:$A$30,0)),2),"Missing part")))`

`MATCH` finds the row of the part on **Parts**, `INDEX` returns that row's unit cost, and `ROUND` keeps the line to whole cents. The `0` at the end of `MATCH` means exact match.

A part name that is not on **Parts** does not show #N/A here. The `IFERROR` turns it into the words Missing part, so you can see which line to fix. A quantity that is text shows Check qty.

Margin here is profit as a share of the price, so $14.77 profit on a $24.00 set is 61.5%. It is not the same as markup, which divides by cost.

Help pages: Microsoft's [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The SUMIF and Google links were not link-checked.

## Example

Sample data is fictional: made up parts, costs and prices for a small gift shop.

The Spa Set is built from these lines:

| Part | Qty | Unit cost | Line cost |
|---|---|---|---|
| Soap bar, oatmeal | 2 | $1.85 | $3.70 |
| Soy candle, 8 oz | 1 | $4.10 | $4.10 |
| Gift box | 1 | $0.95 | $0.95 |
| Tissue paper sheet | 2 | $0.18 | $0.36 |
| Ribbon, 1 yd | 1 | $0.12 | $0.12 |

The lines add up to $9.23. Here are all four sets:

| Set | Cost | Price | Profit | Margin |
|---|---|---|---|---|
| Spa Set | $9.23 | $24.00 | $14.77 | 61.5% |
| Cozy Mug Set | $7.83 | $22.00 | $14.17 | 64.4% |
| Candle Duo | $9.69 | $20.00 | $10.31 | 51.6% |
| Starter Trio | $11.28 | $32.00 | $20.72 | 64.8% |

Change the soy candle to $4.50 on **Parts** and every set that uses it updates on its own. That is the reason to keep one price list instead of typing costs into each set.

Only the parts you list count. If labor, a shipping box or a marketplace fee belongs in the cost, add it as a part. This is arithmetic on your own numbers, not pricing advice.

## Troubleshooting

### A line says Missing part
The part name on **Sets** does not match **Parts** exactly. In the tests, "Soy candle 8 oz" without the comma did not match "Soy candle, 8 oz", and the Spa Set cost fell from $9.23 to $5.13. The margin jumped to 78.6%, which looks great and is wrong. Retype the name or copy it from **Parts**.

### A line says Check qty
The quantity is text, such as the word two, so it cannot be multiplied. Retype it as a number. In the test, that line was left out and the status for the set read Fix lines.

### A set's cost is lower than it should be, with no error
A typo in the set name, such as Spa Sett, drops that line out of the set total. The line still shows a cost, so look for it with the check cell below. In the test, one such line took the Spa Set from $9.23 to $9.11.

### A part is listed twice on Parts
`MATCH` uses the first one it finds and ignores the second. In the test, Gift box was listed at $0.95 and again at $2.50, and the costs stayed at the $0.95 price. Delete the duplicate.

### The margin is blank
The price in column H is empty or not a number. In the test, the cost still showed but margin and profit stayed blank.

## Template

[Download the .xlsx](templates/tidy-tabs-gift-set-cost.xlsx)

The file has three tabs. **Parts** has 7 fictional parts in yellow cells **A5:B30** (26 slots). **Sets** has 19 lines for 4 fictional sets in yellow cells **A5:C54** (50 slots), a set list with prices in **F5:H10**, and cost, margin, profit and status columns. **How to use** has notes on each formula.

The **CHECK** cell, **G12**, says OK, or CHECK if any line says Missing part or Check qty, a part is listed twice, a line has no set name, a set has no price, or a line's set name is not in the set list. Fix every CHECK before you trust a margin. In the tests, each case in Troubleshooting turned it to CHECK.

Limits: 6 sets, 26 parts and 50 lines. More was not tested. **Sets** prints landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
