---
title: How to Calculate Weighted Average Cost With SUMPRODUCT in Excel and Google Sheets
labels: excel-formulas, inventory, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets use the same SUMPRODUCT, SUMIFS and AVERAGEIFS functions; the help links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Weighted average cost with =SUMPRODUCT((Purchases!B2:B40=A2)*qty*price)/total qty. Why it beats a plain average of prices. Free .xlsx template.
image_prompts: Candle-making workbench with a paper sack of soy wax flakes beside a stainless steel pouring pot, and a spool of cotton wick next to a row of empty glass jars, soft window light, no screens, no readable text, no logos
image_alt: Soy wax flakes, a pouring pot, wick spool and glass jars on a workbench
threads: Bought wax at three different prices and need one cost per pound?\nA plain average of the prices gets it wrong. Weight each price by quantity: =SUMPRODUCT((Purchases!B2:B40=A2)*Purchases!C2:C40*Purchases!D2:D40)/B2.\nMy fake soy wax lands at $3.6667 per lb, not $3.7667.
---
Divide the total money spent on an item by the total quantity bought: `=SUMPRODUCT((Purchases!$B$2:$B$40=A2)*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B2`, where B2 is that item's total quantity. This weighted average counts a 40 lb purchase more than a 10 lb one, unlike a plain average of the prices.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Summary table comparing weighted and simple average costs for soy wax, wicks and jars](images/2026-10-09-team-02-weighted-average-cost-sumproduct-excel-google-sheets/weighted-average-cost-summary-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Log every purchase on a **Purchases** tab: one row per buy, with the item name in column B, quantity in C and unit price in D. Then build a **Summary** tab with one row per item.

1. On **Purchases**, type these headers in row 1: Date, Item, Quantity, Unit price, Line total.
2. Enter each buy below the headers. Spell each item name the same way every time.
3. In **E2**, type `=C2*D2` and fill it down. This is the money spent on that purchase.
4. On **Summary**, list each item once in column A, starting at **A2**.
5. In **B2**, type `=SUMIFS(Purchases!$C$2:$C$40,Purchases!$B$2:$B$40,A2)` for the total quantity.
6. In **C2**, type `=SUMIFS(Purchases!$E$2:$E$40,Purchases!$B$2:$B$40,A2)` for the total spent.
7. In **D2**, type `=C2/B2`. This is the weighted average cost.
8. In **E2**, type the SUMPRODUCT formula from the top of this post. It should match **D2** exactly.
9. In **F2**, type `=AVERAGEIFS(Purchases!$D$2:$D$40,Purchases!$B$2:$B$40,A2)`. This is the plain average of the prices, shown only for comparison.
10. Fill row 2 down for every item. Format money cells to four decimals if your unit costs are a few cents.

SUMPRODUCT multiplies matching rows and adds the results. The part `(Purchases!$B$2:$B$40=A2)` turns each row into 1 if the item matches and 0 if not, so only that item's quantity times price is added. Dividing by total quantity gives dollars per unit.

A quick hand check with the first two soy wax buys: 10 lb x $4.00 + 40 lb x $3.50 = $180.00 over 50 lb is $3.60 per lb. The plain average of $4.00 and $3.50 is $3.75, which ignores that you bought four times as much at the lower price.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) and [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) pages.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [SUMPRODUCT](https://support.google.com/docs/answer/3094294) and [SUMIFS](https://support.google.com/docs/answer/3238496) pages.

## Example

Sample data below is fictional. It is a small candle maker's purchases from 07/06/2026 to 09/07/2026: soy wax 10 lb at $4.00 (07/06/2026), 40 lb at $3.50 (08/03/2026) and 25 lb at $3.80 (09/07/2026); cotton wicks 200 at $0.12, 500 at $0.09 and 300 at $0.10; glass jars 48 at $1.85, 96 at $1.60 and 24 at $2.10.

| Item | Total quantity | Total spent | Weighted average | Simple average |
|---|---|---|---|---|
| Soy wax (lb) | 75 | $275.00 | $3.6667 | $3.7667 |
| Cotton wicks | 1,000 | $99.00 | $0.0990 | $0.1033 |
| Glass jars | 168 | $292.80 | $1.7429 | $1.8500 |

The SUMPRODUCT formula and the SUMIFS division returned identical values for all three items in LibreOffice Calc. The simple average is higher every time here because the cheaper buys were the larger ones. Four decimals matter for wicks: at two decimals both averages round to about $0.10 and the difference disappears.

This is an average purchase cost for planning prices. It is not an inventory valuation method and not tax advice, so ask your accountant what your books require.

## Troubleshooting

### An item shows #DIV/0!
The total quantity in **B2** is 0, usually because the item name in **A2** does not match the spelling on **Purchases**. Check for a typo or a trailing space. The SUMIFS and SUMPRODUCT formulas both depend on an exact match.

### The weighted average is wrong after adding new purchases
The formulas read rows 2 to 40. If your log is longer, change 40 to a larger number in every formula. Keep the ranges the same size, because SUMPRODUCT returns an error when ranges have different row counts.

### The SUMPRODUCT result is 0 or an error
A quantity or price stored as text breaks the multiplication. Retype one cell and see if the result changes, or convert the column to real numbers first. This was not reproduced in a test here.

### The weighted average equals the simple average
That happens when every purchase for the item has the same quantity, or the same price. It is not a mistake.

### Does this work with returns or negative quantities?
Not tested. A negative quantity would reduce both the quantity and the money spent, but check one item by hand before you trust the result.

## Template

[Download the .xlsx](templates/tidy-tabs-weighted-average-cost.xlsx)

The file has three tabs. **Purchases** holds the nine sample buys with the line total in column E. **Summary** has the SUMIFS totals, the weighted average two ways and the simple average for comparison. **How to use** has short fill-in notes.

The items and prices are fictional. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
