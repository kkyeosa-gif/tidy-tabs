---
threads_url: https://www.threads.com/@tin_ylab/post/DfkgW8BlKoL
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-round-order-quantity-up-to-case.html
title: How to Round an Order Quantity Up to a Case Size and Minimum Order in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with independent Python models over 39 workbooks and 7,866 cell comparisons. Checked: the sample, 24 edge workbooks (need 0, need below, equal to and above the minimum, a minimum that is not a multiple of the case, case size 1, decimals, negatives, text, blanks, case size 0 and 1.5, only a name, no name, 1,000,000 units at several unit costs) and 14 random lists of 10 to 40 rows. Order cost matched exact Decimal half-up rounding in every row. Excel and Google Sheets were not used. Not run: a non-US language setting, more than 40 rows.
search_description: Round an order quantity up to whole cases and a supplier minimum with ROUNDUP and MAX in Excel and Google Sheets. Free template shows extra units.
image_prompts: Stacked cardboard cases of kraft mailers beside a box cutter and a roll of packing tape on a shelf in a small home workshop, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Stacked cardboard cases of mailers beside a box cutter and packing tape on a shelf
threads: Need 130 mailers, they come 25 to a case, and the supplier has a minimum of 100?\n=ROUNDUP(MAX(B4,D4)/C4,0) gives 6 cases, so 150 units.\nThat is 20 extra, $93.00 at $0.62 each (made-up terms).
---
Get the number of cases with `=ROUNDUP(MAX(B4,D4)/C4,0)`, where B4 is the units you need, D4 is the supplier's minimum order and C4 is the case size. Then get the units to order with `=E4*C4`. Needing 130 mailers with a case of 25 and a minimum of 100 gives 6 cases, or 150 units (made-up supplier terms).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with Python models (39 workbooks, 7,866 comparisons). The formulas use only IF, AND, OR, ISNUMBER, MAX, ROUNDUP and ROUND, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Table of eight supplies with need, case size, minimum, cases, units to order and order cost](images/2026-11-10-team-02-round-order-quantity-up-to-case-size-minimum-order-roundup-excel-google-sheets/order-quantity-case-size-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. Put one supply per row from row 4, with headers in row 3. Columns: **A** item, **B** units needed, **C** case size, **D** minimum order in units, **H** unit cost.

1. In **E4**, type `=ROUNDUP(MAX(B4,D4)/C4,0)`. MAX picks the larger of what you need and the supplier's minimum. Dividing by the case size and rounding up gives whole cases.
2. In **F4**, type `=E4*C4`. This is the number of units you actually order.
3. In **G4**, type `=F4-B4`. This is the extra units you end up holding beyond what you needed.
4. In **I4**, type `=ROUND(F4*H4,2)` for the order cost in dollars.
5. Fill **E4:I4** down as far as your list goes.

The ROUND in step 4 keeps each cost at whole cents, so the column adds up to what a supplier invoice would show.

The template wraps these formulas in a check. A row shows Check, instead of a wrong number, when the need or minimum is not a whole number of 0 or more, or the case size is not a whole number of 1 or more. A need of 0 orders nothing, even when the supplier has a minimum.

The case size and minimum come from your supplier. This post does not suggest what to buy.

Help pages: Microsoft's [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) and [MAX function](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098). They describe the functions, not our test.

## Example

Sample data is fictional: 8 supplies with made-up case sizes, minimums and prices.

| Item | Need | Case | Minimum | Cases | Units to order | Extra | Unit cost | Order cost |
|---|---|---|---|---|---|---|---|---|
| Kraft mailer 6x6x4 | 130 | 25 | 100 | 6 | 150 | 20 | $0.62 | $93.00 |
| Tissue paper packs | 3 | 2 | 0 | 2 | 4 | 1 | $18.50 | $74.00 |
| 8 oz candle tin | 340 | 24 | 240 | 15 | 360 | 20 | $1.15 | $414.00 |
| Soy wax 10 lb box | 0 | 1 | 0 | 0 | 0 | 0 | $54.00 | $0.00 |
| Lavender oil 4 oz | 5 | 6 | 12 | 2 | 12 | 7 | $14.25 | $171.00 |
| Thank-you cards | 40 | 50 | 0 | 1 | 50 | 10 | $0.30 | $15.00 |
| Shipping label roll (500) | 900 | 500 | 1000 | 2 | 1000 | 100 | $0.03 | $30.00 |
| Glass jar 8 oz | 77 | 12 | 0 | 7 | 84 | 7 | $1.40 | $117.60 |
| Total | 1,495 | | | | 1,660 | 165 | | $914.60 |

For the mailers, MAX(130,100) is 130. Divided by 25 that is 5.2, which rounds up to 6 cases. Six cases of 25 is 150 units, so 20 are extra, and 150 times $0.62 is $93.00.

For the lavender oil, the need of 5 is below the minimum of 12, so the minimum wins: 12 divided by 6 is exactly 2 cases.

The labels show both rules at once. MAX(900,1000) is 1000, which is exactly 2 cases of 500. The 100 extra rolls come from the minimum, not the case size.

## Troubleshooting

### A row orders units even though I need 0
The plain formula `=ROUNDUP(MAX(B4,D4)/C4,0)` does that, because the minimum is larger than 0. The template tests for a need of 0 first and returns 0 cases. If that is not how your supplier works, remove the `IF(B4=0,0,` part of the formula.

### The units come out higher than my minimum
The minimum is not a multiple of the case, so it is rounded up to whole cases. A minimum of 100 with a case of 24 gives 5 cases, which is 120 units. Most suppliers only ship whole cases, but check yours.

### The row shows Check
Something in the row is not usable: a decimal such as 12.5, a negative number, text such as "a dozen", a case size of 0 or 1.5, or a minimum typed as text. Fix the cell. Checked rows are left out of the totals, and a counter at the right says how many there are.

### #DIV/0! in the cases column
The case size is 0 or blank. If you build the formula without the template's check, enter a case size of 1 for items sold by the single unit.

### The cost is a cent off the invoice
Your supplier may round differently, or your unit cost has more than 2 decimals. The template rounds each row to cents. Compare the unit cost on the invoice with the one in column **H**.

## Template

[Download the .xlsx](templates/tidy-tabs-order-quantity-case-size.xlsx)

The file has two tabs. **Order** has 40 slots from row 4 to row 43 and 8 fictional sample items. Columns **A** to **J** hold the item, units needed, case size, minimum order, Cases, Units to order, Extra units, Unit cost, Order cost and Row check. A summary block at the right counts rows with an order, units needed, units to order, extra units, order cost, rows showing Check, and an overall CHECK. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
