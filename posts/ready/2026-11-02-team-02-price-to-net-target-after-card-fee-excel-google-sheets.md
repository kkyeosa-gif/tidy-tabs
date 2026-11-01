---
title: How to Set a Price That Still Nets the Amount You Want After a Card Fee in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (Decimal) across 46 workbooks and 1,380 comparisons with 0 mismatches. Checked: fee 0%, flat fee $0.00, fee 99%, fee 100% and 150%, a percent typed as 2.9, a negative flat fee, a blank target, text in the yellow cells and a $1,000,000.00 target. The ROUND result was checked on a sample of 238 prices, not proven for every input. Excel and Google Sheets were not used.
search_description: Find the price to charge so you keep a target amount after a percent card fee plus a flat fee, in Excel and Google Sheets. Free calculator.
image_prompts: A small handmade ceramic mug on a wooden market counter with a blank kraft paper price tag tied to the handle by string and a pencil beside it, soft daylight, no screens, no readable text, no logos
image_alt: Ceramic mug with a blank kraft price tag and a pencil on a market counter
threads: Added 2.9% to your price and still came up short?\nThe fee is taken from the fee too. Divide instead: =ROUND((B2+B4)/(1-B3),2).\nOn a $100.00 goal that is $103.30. The multiply shortcut gives $103.20 and you keep $99.91.
---
To end up with a target amount after a percent fee plus a flat fee, divide instead of multiply: `=ROUND((B2+B4)/(1-B3),2)`, where B2 is the amount you want to keep, B3 the fee percent and B4 the flat fee. Multiplying the target by one plus the percent comes up short because the fee is charged on the fee too.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. ROUND and the other functions used are in every current version of Excel and Google Sheets, but neither app was opened for this post.

![Calculator listing price to charge, fee taken and amount kept for five target amounts](images/2026-11-02-team-02-price-to-net-target-after-card-fee-excel-google-sheets/price-to-net-target-card-fee-calculator.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The rates here are examples. Type the percent and flat fee from your own payment processor's current price list. This post makes no claim about any provider's rates.

1. In **B2**, type the amount you want to keep, such as $100.00.
2. In **B3**, type the percent fee as a percent, such as 2.9%. Format the cell as a percent with one decimal.
3. In **B4**, type the flat fee per sale, such as $0.30.
4. In **B6**, type `=ROUND((B2+B4)/(1-B3),2)`. This is the price to charge.
5. In **B7**, type `=ROUND(B6*B3+B4,2)` for the fee taken.
6. In **B8**, type `=ROUND(B6-B7,2)` for what you keep, and in **B9** type `=ROUND(B8-B2,2)` for the difference from your target.

The fee is a share of the price, not of the target. Say the price is P. You keep P minus (P times the percent) minus the flat fee. Set that equal to the target and solve for P, and you get the target plus the flat fee, divided by one minus the percent.

`ROUND` keeps the price to whole cents. In the tests, it landed on the target to the cent. See the note on rounding in the example below.

If you type a percent as a plain number, such as 2.9 instead of 2.9%, the sheet reads it as 290%. The input check in **B5** catches that and shows a message instead of a price.

Help pages: Microsoft's [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Neither link was checked.

## Example

Sample data is fictional: example rates and target amounts for a small shop.

You want to keep $100.00, and the fee is 2.9% plus $0.30.

| Method | Price | Fee taken | You keep |
|---|---|---|---|
| Divide (B6) | $103.30 | $3.30 | $100.00 |
| Multiply shortcut (B10) | $103.20 | $3.29 | $99.91 |

The shortcut is `=ROUND(B2*(1+B3)+B4,2)`. It is $0.09 short here, and the gap grows with the amount. At a $1,000,000.00 target, the shortcut came up $841.01 short.

The same formulas fill the table in **A16:E21**:

| You want to keep | Price to charge | Fee taken |
|---|---|---|
| $25.00 | $26.06 | $1.06 |
| $48.00 | $49.74 | $1.74 |
| $100.00 | $103.30 | $3.30 |
| $250.00 | $257.78 | $7.78 |
| $1,250.00 | $1,287.64 | $37.64 |

Every difference column reads $0.00.

A note on rounding. In the tests, ROUND hit the target exactly for all 238 prices checked. Rounding up always would have overshot by a cent in 45 of the same 238. That is a sample, not a proof for every possible input.

This is arithmetic on your own numbers. It does not say whether you may pass a fee on to customers.

## Troubleshooting

### The price cell is blank and B5 shows a message
An input is out of range. The percent must be 0% up to under 100%, and the amounts must be 0 or more. In the tests, a fee of 100% or 150%, a negative flat fee, a blank target or text in a yellow cell left the price and the table blank, with status Check inputs and no #DIV/0!.

### The percent shows as 290%
You typed 2.9 into a cell that is not formatted as a percent. Type 2.9% or format the cell first. The sheet's message covers the wrong value. How Excel itself treats a percent typed into an unformatted cell was not tested.

### The status says 1 cent short from rounding
The price was rounded to a cent and the result landed a cent under the target. In the tests, this did not happen on the valid inputs checked. If you see it, add a cent to the price.

### A table row is empty
The target in column A is blank or text, so the row is skipped. A negative target still computes, such as $-4.84 for a $-5.00 goal, so look at your entries.

### The fee is 0%
The price is the target plus the flat fee, such as $100.30. This is correct. With a flat fee of $0.00 and 2.9%, the price for $100.00 is $102.99.

## Template

[Download the .xlsx](templates/tidy-tabs-price-after-card-fee.xlsx)

The file has two tabs. **Calculator** has three yellow input cells (**B2:B4**) with fictional example values, the price, fee, kept amount and difference in **B6:B9**, the wrong shortcut and its shortfall in **B10:B12**, a status in **B13** and a five-row table of targets in **A17:A21**. **How to use** has notes on each formula.

The **CHECK** cell, **B14**, says OK, or Check if an input is out of range, a price is blank, any table row is short by more than a cent, or a target is missing a result. Fix every Check before you trust a price.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
