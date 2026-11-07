---
title: How to Combine All Items in an Order Into One Cell in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 16 workbooks and 648 comparisons with 0 mismatches. Checked: the sample, no lines, one line, one order of 3 lines, one line per order, orders not on the Lines tab, text order numbers, a number and text mixed, unsorted lines, a 528 character result, 100 lines (32 orders), and 5 random sets. Excel and Google Sheets were not used. Not run: MATCH type 1 with empty trailing cells in Excel or Sheets, more than 100 lines, a non-US language setting.
search_description: Combine every item in an order into one cell in Excel and Google Sheets with a running join and INDEX/MATCH. No TEXTJOIN. Free template included.
image_prompts: An open gift box holding a candle jar and a lip balm tin on kraft tissue paper, with a roll of twine beside it, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Open gift box holding a candle jar and lip balm tin on tissue paper
threads: "3 x Enamel pin, 2 x Sticker pack, 1 x Soy candle Cedar, 4 x Lip balm" in one cell, no TEXTJOIN or array formula.\nSort the lines by order number, then =IF(A2=A1,D1&", "&C2&" x "&B2,C2&" x "&B2) builds the text row by row.\nINDEX/MATCH fetches the last row of each order.
---
On each line of an order-sorted list, build the text with `=IF(A2=A1,D1&", "&C2&" x "&B2,C2&" x "&B2)`, then fetch the finished text for an order with `=INDEX(Lines!$D$2:$D$101,MATCH(A2,Lines!$A$2:$A$101,1))` on the Orders tab. Order 2003 comes out as "3 x Enamel pin, 2 x Sticker pack, 1 x Soy candle Cedar, 4 x Lip balm" with no array formula and no TEXTJOIN.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every result below was recalculated and compared with a Python model (16 workbooks, 648 comparisons). The formulas use only IF, INDEX, MATCH, COUNTIF and the & operator, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Five order numbers each paired with all their items joined in one cell](images/2026-11-08-team-03-combine-order-items-into-one-cell-running-join-excel-google-sheets/order-items-in-one-cell-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional.*

## Steps

The formulas are the same in Excel and Google Sheets. Only the sort command differs. The **Lines** tab has one row per item: order number in **A**, item in **B**, quantity in **C**. Row 1 is a header, and that matters below.

### Sort the lines first

The running join only works when every order's rows sit together and the orders are in ascending order. Sort before you add the formulas.

- In Excel, select your data and choose **Data > Sort**, then sort by the Order column, smallest to largest.
- In Google Sheets, select your data and choose **Data > Sort range > Advanced range sorting options**, tick "Data has header row", and sort column A from A to Z.

Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) page explains why: the match type 1 used later expects ascending data.

### Build the running text

1. In **D2**, type `=IF(A2="","",IF(A2=A1,D1&", "&C2&" x "&B2,C2&" x "&B2))`.
2. Fill **D2** down as far as your data goes. The template fills it to row 101.

When the order number matches the row above, the formula adds this line to the text above it. When it does not, it starts fresh text. Row 2 always starts fresh because **A1** is the header, which is why you need one.

The last row of each order now holds the full text. The rows above it hold partial text.

### Fetch one order's text

1. On the **Orders** tab, type an order number in **A2**, such as 2003. Keep these ascending too.
2. In **B2**, type `=INDEX(Lines!$D$2:$D$101,MATCH(A2,Lines!$A$2:$A$101,1))`.
3. Fill **B2** down.

`MATCH` with type 1 returns the position of the last row whose order number is less than or equal to the one you typed. On sorted data that is the last row of the order, which holds the complete text. Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) page covers the other half.
Google's [MATCH help page](https://support.google.com/docs/answer/3093378) describes the same match type 1 behavior.

The template adds one check: **Orders!B2** also tests that the row found really belongs to your order, so a missing order shows Not found instead of borrowing the previous order's text.

## Example

Sample data is fictional: 12 lines in 5 orders.

| Order | Item | Qty |
|---|---|---|
| 2001 | Soy candle Lavender | 2 |
| 2001 | Lip balm | 1 |
| 2001 | Gift box | 1 |
| 2002 | Tote bag | 1 |
| 2003 | Enamel pin | 3 |
| 2003 | Sticker pack | 2 |
| 2003 | Soy candle Cedar | 1 |
| 2003 | Lip balm | 4 |
| 2004 | Tote bag | 2 |
| 2004 | Gift box | 1 |
| 2005 | Sticker pack | 5 |
| 2005 | Enamel pin | 1 |

The Orders tab then shows:

| Order | Items | Lines |
|---|---|---|
| 2001 | 2 x Soy candle Lavender, 1 x Lip balm, 1 x Gift box | 3 |
| 2002 | 1 x Tote bag | 1 |
| 2003 | 3 x Enamel pin, 2 x Sticker pack, 1 x Soy candle Cedar, 4 x Lip balm | 4 |
| 2004 | 2 x Tote bag, 1 x Gift box | 2 |
| 2005 | 5 x Sticker pack, 1 x Enamel pin | 2 |

The summary reads 5 orders listed, 0 not found, 12 lines found, 12 lines on the Lines tab, 0 order numbers split apart, and CHECK says OK.

### What breaks without the sort

We also ran lines for orders 2, 1, 2, 3 in that order, with no sort. Order 2 returned only "1 x A" instead of both of its lines, "1 x A, 3 x C". The template's **Lines!F** column showed Order split, the split count became 1 and CHECK turned to Check. Sorting by order number fixed it.

## Troubleshooting

### An order shows only some of its items
The order's rows are not together. A second group of the same order number later in the list starts new text, and the lookup finds only the last group. Sort the Lines by order number. In the template, **Lines!F** says Order split on the first row of the second group.

### Two orders that look the same are split
A number 2001 and a text "2001" look alike but are different values, so they break an order into two groups. Make the whole column numbers or the whole column text. The template flags the mix as Order split.

### The Orders tab says Not found
That order number is not on the Lines tab, or it is spelled differently. Check for a typo, or a stray space in a text order number.

### The result is cut off or I get an error about 255 characters
The 255 character limit applies to text typed inside a formula, not to a result. A 528 character result worked in LibreOffice. A narrow column only hides the rest, so widen it or turn on wrap text.

## Template

[Download the .xlsx](templates/tidy-tabs-items-per-order-one-cell.xlsx)

The file has three tabs. **Lines** has 12 fictional sample lines and room for 100, with the running text in **D**, a last line flag in **E** and a sort check in **F**. **Orders** has 5 sample order numbers in yellow and room for 40, with the items, line count and lookup check in **B:D**, and a summary in **G1:G6**: orders listed, not found, lines found, lines on the Lines tab, order numbers split apart, and a **CHECK** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The file uses no TEXTJOIN and no dynamic arrays.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
