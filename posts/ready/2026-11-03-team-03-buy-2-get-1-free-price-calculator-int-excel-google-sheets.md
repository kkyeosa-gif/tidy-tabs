---
title: How to Price Buy 2 Get 1 Free Deals in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (Decimal) over 24 workbooks and 1,276 comparisons with 0 mismatches. Checked: prices from $0.00 to $100.01, buy counts 1 to 5, free counts 0 to 3, carts from 0 to 100 items, buy of 0, a buy count of 2.5, a negative free count, and text in a yellow cell. Excel and Google Sheets were not used.
search_description: Price a buy 2 get 1 free deal with INT in Excel and Google Sheets. Free calculator for any buy X get Y free offer with a cart table.
image_prompts: Three tied bars of handmade soap stacked on a market table next to a small chalkboard sign with no writing and a brown paper bag, soft daylight, no screens, no readable text, no logos
image_alt: Stack of three soap bars beside a blank chalkboard sign and a paper bag
threads: Buy 2 get 1 free, cart of 7. How many are free?\n=INT(7/3)*1. Two full deals of 3, so 2 free items.\nAt $18.00 each you pay for 5 and the cart is $90.00 instead of $126.00. The sheet works for buy 3 get 2 too, and flags a buy of 0.
---
For buy 2 get 1 free, count the free items with `=INT(B9/B6)*B4`, where B9 is the number of items in the cart, B6 is the size of one full deal (paid plus free, so 3) and B4 is the number free per deal. Charge for the rest: a cart of 7 at $18.00 has 2 free items and costs $90.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. INT and ROUND exist in current Excel and Google Sheets, but neither app was opened for this post.

![Deal calculator with cart total and a table of cart sizes showing free items and totals](images/2026-11-03-team-03-buy-2-get-1-free-price-calculator-int-excel-google-sheets/buy-2-get-1-free-price-calculator.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The prices here are examples, so type your own.

1. In **B2**, type the price of one item, such as $18.00.
2. In **B3**, type how many items are paid in each deal (2). In **B4**, type how many are free (1).
3. In **B6**, type `=B3+B4`. This is the size of one full deal.
4. In **B9**, type the number of items in the cart, such as 7.
5. In **B10**, type `=INT(B9/B6)*B4` for the free items.
6. In **B11**, type `=B9-B10` for the items you pay for. In **B12**, type `=ROUND(B11*B2,2)` for the cart total.

`INT` drops the part after the decimal point. Seven items divided by 3 is 2.33, so INT gives 2 full deals. Two deals times 1 free item is 2 free items. The seventh item has no partner, so it is paid.

To compare, **B13** is the cost without the deal, `=ROUND(B9*B2,2)`, and **B14** is the saving, `=ROUND(B13-B12,2)`.

This assumes the free items cost the same as the paid ones and that every full group earns its free items. It is arithmetic, not a policy for which items go free.

Help pages: Microsoft's [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/02/2026; the pages were not read for this post.

## Example

Sample data is fictional: soap bars at $18.00 each, buy 2 get 1 free.

| Items in cart | Items you pay for | Free items | Cart total |
|---|---|---|---|
| 1 | 1 | 0 | $18.00 |
| 2 | 2 | 0 | $36.00 |
| 3 | 2 | 1 | $36.00 |
| 6 | 4 | 2 | $72.00 |
| 7 | 5 | 2 | $90.00 |
| 12 | 8 | 4 | $144.00 |

A cart of 7 costs $90.00 and would be $126.00 without the deal, a saving of $36.00. For buy 3 get 2 free, set **B3** to 3 and **B4** to 2. The deal size becomes 5.

## Troubleshooting

### B5 shows a message and the totals are blank
An input is out of range. Buy must be 1 or more, free must be 0 or more, both whole numbers, and the price must be 0 or more. In the test, a buy of 0, a buy of 2.5, a negative free count and text in a yellow cell each blanked the totals with a message and no error value.

### The deal starts too early
A cart of 2 gets no free item for buy 2 get 1. The customer needs the third item for the free one. If you want the free item to come from the cart's cheapest items, this sheet does not do that.

### The average per item looks odd
Column E is the cart total divided by the items in the cart, rounded to a cent. A cart of 7 is $12.86, not $12.00. Only full groups of 3 reach $12.00.

### A cart size is blank in the table
The cell in column A is blank or text, and the row is skipped. A decimal such as 2.5 also leaves the row blank, and in the test the **CHECK** cell then said Check.

### The price has half cents
Money is rounded to cents at the cart total. A price such as $100.01 can still produce a half-cent average that rounds up. That is how ROUND works.

## Template

[Download the .xlsx](templates/tidy-tabs-buy-x-get-y-free.xlsx)

The file has two tabs. **Deal** has three yellow settings (**B2:B4**), an input check in **B5**, a one-cart calculation in **B9:B14**, a **CHECK** cell in **B15** and a table of 12 cart sizes in **A19:E30**. **How to use** has notes on each formula.

**CHECK** says OK when the inputs are valid and the cart total and table rows are numbers. Fix every Check before you quote a total.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
