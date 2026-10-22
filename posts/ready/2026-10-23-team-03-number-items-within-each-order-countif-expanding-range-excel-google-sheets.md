---
title: How to Number Items Within Each Order in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model, including edited copies: rows of one order not next to each other, a1003 in lower case, a trailing and a leading space, a blank order ID with an item, an ID with no item, a blank row in the middle, one row only, no rows, 199 random IDs, numeric IDs, IDs with * or ?, and "007" against the number 7. Excel and Google Sheets were not opened.
search_description: Number the items within each order in Excel and Google Sheets with COUNTIF. Free template with item numbers and first and last item flags.
image_prompts: A stack of small kraft shipping boxes with a roll of twine and a hand-tied paper tag on a packing table, soft window light, no screens, no readable text, no logos
image_alt: Stacked kraft shipping boxes with a twine roll and blank paper tag on a packing table
threads: Order export with one row per item, and you want 1, 2, 3 that restarts for each order?\n=COUNTIF($A$2:A2,A2) in row 2, fill down. The range grows, so a new order ID starts at 1 again.\nCareful: "A1003 " with a trailing space counts as a different order.
---
Number the items inside each order with `=COUNTIF($A$2:A2,A2)` in row 2, then fill down, where column A holds the order ID. The range grows by one row each time, so the count starts at 1 for each new order ID.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Order list with item numbers restarting at 1 for each order ID and First, Last and Only labels](images/2026-10-23-team-03-number-items-within-each-order-countif-expanding-range-excel-google-sheets/order-line-numbers-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: an order export has one row per item, so the same order ID repeats. Each row counts how many times its own ID has shown up so far, from the top of the list down to that row.

1. Put the order ID in column **A** and the item in column **B**, starting in row 2. Row 1 holds headings.
2. In **C2**, type `=COUNTIF($A$2:A2,A2)`.
3. Select C2, then double-click the small square at its bottom right corner to fill the formula down. Or drag it down yourself.
4. Optional: in **D2**, type `=COUNTIF($A$2:$A$200,A2)` to get the total number of items in that order.

Here is why it restarts at 1. The first end of the range is locked as `$A$2`. The second end, `A2`, has no dollar signs, so it moves down one row as you fill. Row 2 counts A2 only. Row 3 counts A2:A3, row 4 counts A2:A4, and so on. When a new order ID appears, there are no earlier rows with that ID, so the count is 1.

In D2, both ends are locked, so every row counts the whole list. That gives the total for the order.

### What the template adds

The template wraps the same core in `IF(A2="","",...)`, so rows with no order ID stay blank. It also fills three more columns:

- **E**, Position: "Only" for a one item order, "First" for item 1, "Last" for the final item, and blank in between.
- **F**, Order label: the ID, the item number and the total, such as A1003-2 of 3.
- A summary block in **H1:I8** with a check cell in **I8**.

### In Excel

The formulas work as typed. COUNTIF is described on Microsoft's [COUNTIF function page](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34). That is the help page, not our test.

### In Google Sheets

The same formulas should work as typed. Google Sheets was not opened for this post.

## Example

Sample data is fictional: Etsy style orders from a small candle and soap shop.

| Row | Order ID | Item | Item number | Items in order | Position |
|---|---|---|---|---|---|
| 2 | A1001 | Lavender soy candle | 1 | 2 | First |
| 3 | A1001 | Cedar soy candle | 2 | 2 | Last |
| 4 | A1002 | Oatmeal soap bar | 1 | 1 | Only |
| 5 | A1003 | Brass hoop earrings | 1 | 3 | First |
| 6 | A1003 | Letterpress birthday card | 2 | 3 | |
| 7 | A1003 | Kraft gift box | 3 | 3 | Last |

Row 6 is a middle item, so its Position stays blank. Column F labels read A1003-1 of 3, A1003-2 of 3, and so on.

The template holds 24 items in 9 orders, A1001 to A1009. The summary shows 24 rows, 9 orders (the number of rows numbered 1), and 4 as the most items in one order. **I8** shows OK.

The rows of one order do not have to sit next to each other. In an edited copy with rows swapped, each row still got the right item number going down the list, and the totals were right. **I6** counts separated rows, and it counted 2 there. Sort by Order ID first for a clean packing list.

## Troubleshooting

### One order shows up as two orders
Look for a space at the end or the start of the ID, such as "A1003 " with a trailing space. COUNTIF treats it as a different order. In the edited copy, the sample went from 9 orders to 10, the cell turned red, and the check cell said CHECK. In a spare column, `=TRIM(A2)` gives a clean copy. Microsoft describes it on the [TRIM function page](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9).

### Different orders get merged, or the count is too high
A `*` or `?` inside an order ID acts as a wildcard in COUNTIF. In LibreOffice Calc, a probe with the IDs A*1, A1, AB1 and A? numbered A*1 as 4 items, because it also matched A1 and AB1. A? was counted as 3. The template counts these IDs in **I7** and shows CHECK. Rename them, for example swap the * for a dash. Not checked in Excel or Google Sheets.

### 007 and 7 count as the same order
In LibreOffice Calc, the text "007" and the number 7 were counted as one order. If your IDs are all digits, type them the same way everywhere, all as text or all as numbers. This was not checked in Excel or Google Sheets.

### a1003 and A1003 count as one order
That is expected. COUNTIF ignores capitalization, and an edited copy with a1003 confirmed it in LibreOffice.

## Template

[Download the .xlsx](templates/tidy-tabs-order-line-numbers.xlsx)

The file has two tabs. **Orders** holds the yellow inputs in **A2:B200** (order ID and item), the formulas in **C2:F200** (item number, items in order, position, order label), and the summary and check in **H1:I8**. **How to use** has short notes. Both tabs print portrait on US Letter, and the print area is A1:F25.

The orders and items are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file. The formulas cover rows 2 to 200.
