---
title: How to Find the Most Common Value in a List in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and every result was compared with a Python Counter on the sample and on 15 edited copies (a tie, all items different, one item only, blank rows in the middle, upper and lower case, a trailing space, quantities typed as text, numbers as items). Excel and Google Sheets were not opened.
search_description: Find the most common value in a list in Excel and Google Sheets with COUNTIF, INDEX and MATCH. Free template with sample orders, a top values table and a tie message.
image_prompts: A small candle maker's packing table with a row of lidded amber jars in three scents, a stack of kraft shipping boxes with twine, and a brass kitchen scale, soft window light, no screens, no readable text, no logos
image_alt: Amber candle jars beside kraft shipping boxes and a brass scale on a packing table
threads: Which scent shows up in your orders the most? MODE only works on numbers.\nFor text, COUNTIF each row in a helper column, then INDEX/MATCH the biggest count.\nIn the sample, Lavender wins with 11 of 40 orders. Tie? The template says so.
---
Count how often each row's value appears with `=IF(B2="","",COUNTIF($B$2:$B$500,B2))` in a helper column, then return the most common value with `=IFERROR(INDEX(Orders!$B$2:$B$500,MATCH(MAX(Orders!$D$2:$D$500),Orders!$D$2:$D$500,0)),"")`. Here B holds the item and D holds the count.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and the results below were recalculated and compared with a Python Counter. The steps for Excel and Google Sheets use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Summary sheet showing the most common candle scent, its count, share of orders and a top five table](images/2026-10-26-team-01-most-common-value-mode-index-match-countif-excel-google-sheets/most-common-value-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

`MODE` finds the most common number, but it returns an error on text. For text, you count each row yourself and then look up the row with the biggest count.

1. Put your list on a sheet named **Orders**. Type the order in **A**, the item (a size, scent or color) in **B** and the quantity in **C**, starting in row 2.
2. In **D2**, type `=IF(B2="","",COUNTIF($B$2:$B$500,B2))`. Copy it down to row 500. Every row of the same item now shows the same count.
3. On a sheet named **Summary**, type the formula from the top of this post in any cell. It returns the item from the first row that has the biggest count.
4. To see how often it appears, type `=MAX(Orders!$D$2:$D$500)`.
5. For the most common number in column C, type `=IFERROR(MODE(Orders!$C$2:$C$500),"No repeats")`.

How the formula reads:

- `COUNTIF($B$2:$B$500,B2)` counts how many cells in the list equal this row's item.
- `MAX(...)` finds the biggest of those counts.
- `MATCH(...,0)` finds the first row that has that count.
- `INDEX(...)` returns the item from that row.
- `IFERROR(...,"")` keeps the cell empty when the list is empty.

The template adds helper columns **E** to **H** and a few more Summary cells. They exist only for the tie message, the top 5 table and the quantity tie. You can skip them if you only want the answer.

### In Excel

The formulas should work as typed. COUNTIF, INDEX, MATCH, MAX, MODE and IFERROR are old functions, so no newer function is needed. Menu paths and function details are on Microsoft's pages: [COUNTIF](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [MODE](https://support.microsoft.com/en-us/office/mode-function-e45192ce-9122-4980-82ed-4bdc34973120). Those pages describe the functions, not our test.

### In Google Sheets

The same formulas should work as typed. Google's page for [COUNTIF](https://support.google.com/docs/answer/3093480) describes the function. To start from our file, use **File > Import > Upload**.

## Example

Sample data is fictional: 40 candle orders, E-2001 to E-2040, for an online shop. The first rows look like this.

| Order | Item size or scent | Quantity | Times this item is listed |
|---|---|---|---|
| E-2001 | Lavender | 2 | 11 |
| E-2002 | Lavender | 4 | 11 |
| E-2003 | Sandalwood | 1 | 5 |
| E-2004 | Vanilla | 3 | 8 |
| E-2005 | Cedar | 6 | 9 |
| E-2006 | Sandalwood | 1 | 5 |

The Summary tab then showed these results after recalculation:

| Cell | Result |
|---|---|
| Orders counted | 40 |
| Most common item | Lavender |
| Times it appears | 11 |
| Share of orders | 27.5% |
| Tie? | No tie |
| Most common quantity | 1 (13 times) |
| Check | OK |

The top 5 table lists every item once, biggest count first.

| Rank | Item | Orders | Share |
|---|---|---|---|
| 1 | Lavender | 11 | 27.5% |
| 2 | Cedar | 9 | 22.5% |
| 3 | Vanilla | 8 | 20.0% |
| 4 | Citrus | 7 | 17.5% |
| 5 | Sandalwood | 5 | 12.5% |

The sheet counts. It does not say what to stock or buy, and it does not rank dollars.

## Troubleshooting

### Two items have the same count
Only one item can be named, and `MATCH` picks the one listed first. The template counts the items that reach the biggest count and says so. In a copy with S and M at 2 orders each, S was named and the tie cell read "2 items tie at 2 orders; the first one listed is shown". The same message says "1 orders" when the count is 1, because the text does not change the word.

### Cedar and cedar are counted together
`COUNTIF` ignores upper and lower case. In a copy with Cedar, cedar and CEDAR, the template found one item with 3 orders and showed the spelling listed first.

### One item is counted separately
A trailing space makes a different item. In a copy with "Cedar " (a space at the end) next to Cedar, they were counted apart: Cedar 2 and "Cedar " 1. Retype the cell, or clean the column with `=TRIM(B2)` in a spare column and paste the values back.

### A quantity is ignored
`MODE` skips numbers stored as text, and the check cell shows CHECK. In a copy with quantities typed as text, they were not counted. Retype them as numbers.

### Most common quantity says No repeats
Every quantity is different, so `MODE` has no answer. Without `IFERROR`, it returns #N/A. In a copy where every item was different, the item cell named the first one listed and the tie message said all of them tied at 1.

## Template

[Download the .xlsx](templates/tidy-tabs-most-common-value.xlsx)

The file has three tabs. **Orders** holds the yellow order, item and quantity columns in **A:C** (rows 2 to 500) and the gray helper formulas in **D:H**. **Summary** holds the most common item, its count and share, the tie message, the most common quantity and its tie message, the top 5 table in **D1:G6**, and the check in **B9**. **How to use** has short notes.

The shop and the orders are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
