---
title: How to Set a Wholesale Price From the Retail Price and Check Your Margin in Excel and Google Sheets
labels: excel-formulas, pricing, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 27 workbooks and 13,176 comparisons with 0 mismatches. Also exported to PDF on US Letter with no error values. Excel and Google Sheets were not used.
search_description: Set a wholesale price from the retail price and check your margin in Excel and Google Sheets with ROUND. Free template flags products below target.
image_prompts: A brass kitchen scale holding a small kraft box, a roll of twine and a blank hang tag on a plain wooden shelf, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Kraft gift box on a brass scale beside twine and a blank hang tag
threads: Wholesale price = retail times a percent you pick: =ROUND(C7*$B$2,2).\nIn the sample, 50% of an $18.00 retail price is $9.00, which leaves a 54.4% margin on a $4.10 cost.\nThe free template flags products below your target margin or selling at a loss.
---
Get the wholesale price with `=ROUND(C7*$B$2,2)` and your margin at that price with `=E7/D7`, where E7 is `=ROUND(D7-B7,2)`. At 50% of an $18.00 retail price the wholesale price is $9.00, and a $4.10 unit cost leaves a 54.4% margin (made-up numbers; the percent is yours to choose).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was built there and every number below was recalculated and compared with a Python model (27 workbooks, 13,176 comparisons). ROUND, ROUNDUP and IF exist in Excel and Google Sheets, but neither was opened for this post.

![Price table with wholesale price, profit and margin columns for 11 sample products](images/2026-11-07-team-05-wholesale-price-from-retail-price-margin-check-excel-google-sheets/wholesale-price-from-retail-price-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. Sample data is fictional. On **Prices**, the wholesale percent goes in **B2** and your target margin in **B3**. Products start in row 7: product in **A**, your unit cost in **B**, retail price in **C**.

1. In **B2**, type the percent of retail you want to charge wholesale buyers, such as 50%. Type it with the percent sign, not as 50.
2. In **B3**, type the margin you want to keep on wholesale sales, such as 40%.
3. In **B4**, type `=IF(AND(ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B2>0,B2<=1,B3>=0,B3<1),"OK","Check"),"Check")`. It shows Check if a percent was typed as 50 or left as text.
4. In rows 7 and below, enter each product's cost in **B** and retail price in **C**.
5. In **D7**, type `=ROUND(C7*$B$2,2)`. This is the wholesale price, rounded to the cent. The template wraps it in more checks, so a retail price of 0, a negative cost or text shows Check.
6. In **E7**, type `=ROUND(D7-B7,2)` for profit per unit at wholesale.
7. In **F7**, type `=E7/D7` for margin at wholesale. In **G7**, type `=(C7-B7)/C7` for margin at retail.
8. In **I7**, type `=IF(ISNUMBER(F7),IF(E7<0,"Loss",IF(ROUND(F7,6)<ROUND($B$3,6),"Below target","OK")),F7)`.
9. In **J7**, type `=ROUNDUP(ROUND(B7/(1-$B$3),6),2)` for the lowest wholesale price that reaches your target. In **K7**, type `=ROUNDUP(ROUND(J7/$B$2,6),2)` for the retail price that gets you there.
10. Format **B:E**, **J** and **K** as currency and **F:H** as 0.0%. Fill the row down.

Margin is profit divided by the price. Markup is profit divided by cost. If those two get mixed up, see our [markup vs margin post](posts/published/2026-09-30-team-03-markup-vs-margin-calculator-excel-google-sheets.md).

The 50% and 40% are example inputs. This sheet does not tell you what percent or margin to use. It also ignores shipping, fees, minimum orders and taxes: it is arithmetic on the costs you type.

Help pages: Microsoft's [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7). They describe the functions, not our test.

## Example

Sample data is fictional: 11 products, wholesale at 50% of retail, target margin 40%. Three of the rows are shown.

| Product | Unit cost | Retail | Wholesale | Profit per unit | Margin at wholesale | Status | Wholesale needed for target | Retail needed for target |
|---|---|---|---|---|---|---|---|---|
| Soy candle, Lavender, 8 oz | $4.10 | $18.00 | $9.00 | $4.90 | 54.4% | OK | $6.84 | $13.68 |
| Enamel mug, Blue, 12 oz | $7.80 | $16.00 | $8.00 | $0.20 | 2.5% | Below target | $13.00 | $26.00 |
| Bath bomb gift box | $11.00 | $19.99 | $10.00 | -$1.00 | -10.0% | Loss | $18.34 | $36.68 |

The candle's margin at retail is 77.2%, and its wholesale price is 50.0% below retail. The summary shows 11 products listed, 6 at or above target, 4 below target, 1 selling at a loss, 0 rows showing Check, and a lowest margin at wholesale of -10.0%.

The wholesale price is rounded to the cent before the margin is computed. The status compares margins rounded to 6 decimals, so a margin exactly on target counts as OK.

## Troubleshooting

### A row says Check
The cost or retail price is text, the cost is negative, or the retail price is 0. Retype the numbers as plain numbers. If every row says Check, look at **B4** first.

### B4 says Check
A percent was typed as 50 instead of 50% or 0.5, or the cell is text. The wholesale percent must be above 0 and at most 100%. The target margin must be at least 0 and below 100%.

### The margin looks too high
You may be reading markup. Profit divided by cost gives a bigger number than profit divided by price. Column **F** divides by the wholesale price.

### Wholesale needed for target is a cent above what I expected
**J** and **K** are the lowest prices in whole cents that reach the target after rounding, so they round up, never down.

## Template

[Download the .xlsx](templates/tidy-tabs-wholesale-from-retail-price.xlsx)

The file has two tabs. **Prices** has 60 product rows from row 7 with 11 fictional samples, yellow input cells in **B2**, **B3** and the cost and retail columns, and gray formula cells. The summary in **N6:N12** counts products listed, at or above target, below target, selling at a loss and rows showing Check, shows the lowest margin at wholesale, and ends with a **CHECK** cell. **How to use** has notes on each formula. Prices is set to landscape.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
