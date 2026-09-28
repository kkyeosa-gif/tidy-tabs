---
title: How to Count Unique Values in a Column in Excel and Google Sheets
labels: excel-formulas, data-cleanup
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A printed customer order list on a desk with a yellow highlighter marking repeated names and a sticky note flag
image_alt: Printed order list with highlighter marks over repeated customer names
search_description: How to count unique values in a column using COUNTUNIQUE in Google Sheets and UNIQUE or COUNTIF in Excel, with formulas and troubleshooting.
threads: Your customer list has 400 rows but how many are actual different people? Google Sheets has a one-word answer: COUNTUNIQUE. Excel needs UNIQUE wrapped in COUNTA, or SUMPRODUCT if you're on an older version without dynamic arrays
---

Count unique values with `=COUNTUNIQUE(range)` in Google Sheets. In Excel with a Microsoft 365 subscription, use `=COUNTA(UNIQUE(range))`. On older Excel versions without the UNIQUE function, use `=SUMPRODUCT(1/COUNTIF(range,range))`.

> Works in: Google Sheets (any current version) and Excel for Microsoft 365. The SUMPRODUCT formula also works in Excel 2016 and 2019. Steps come from Microsoft and Google help pages, not hands-on testing.

![Spreadsheet example with columns Row, Customer Name](images/2026-09-28-zauto-01-how-to-count-unique-values-in-a-column-i/how-to-count-unique-values-in-a-column-in-excel-an-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

### In Google Sheets

1. Click an empty cell.
2. Type `=COUNTUNIQUE(` and then select the range, for example `A2:A150`.
3. Close the parenthesis and press **Enter**.
4. Blank cells in the range are ignored automatically.

### In Excel (Microsoft 365 or Excel 2021)

1. Click an empty cell.
2. Type `=COUNTA(UNIQUE(A2:A150))` and press **Enter**.
3. `UNIQUE` returns a spilled list of the distinct values, and `COUNTA` counts how many there are.
4. If you want to see the actual list of unique values instead of just the count, type `=UNIQUE(A2:A150)` on its own and let it spill into the cells below.

### In Excel (2019, 2016, or no UNIQUE function)

1. Click an empty cell.
2. Type `=SUMPRODUCT(1/COUNTIF(A2:A150,A2:A150))` and press **Enter**.
3. This formula divides 1 by the count of each value's occurrences, then adds the results, which totals to the number of distinct values.
4. If the range has any blank cells, this formula returns a `#DIV/0!` error. Narrow the range to exact data rows, or use `=SUMPRODUCT((A2:A150<>"")/COUNTIF(A2:A150,A2:A150&""))` instead.

## Example

Sample data below is fictional, a list of customer names from order rows on a farmers market sign-up sheet.

| Row | Customer Name |
|---|---|
| 1 | Dana Ruiz |
| 2 | Marcus Lee |
| 3 | Dana Ruiz |
| 4 | Priya Shah |
| 5 | Marcus Lee |
| 6 | Dana Ruiz |

This range has 6 entries but only 3 distinct names: Dana Ruiz, Marcus Lee, and Priya Shah. `COUNTUNIQUE(A1:A6)` and `COUNTA(UNIQUE(A1:A6))` both return `3`.

## Troubleshooting

### The count is higher than expected because of trailing spaces
"Dana Ruiz " with a trailing space counts as different from "Dana Ruiz". Wrap the range in `TRIM()` first, for example `=COUNTUNIQUE(ARRAYFORMULA(TRIM(A2:A150)))` in Sheets.

### Names with different capitalization count as separate values
`COUNTUNIQUE` and `UNIQUE` are not case-sensitive by default in most cases, but mixed data from a CSV import can still behave inconsistently. If you suspect this, convert the range to one case first with a helper column using `=LOWER(A2)`, then count uniques on that column.

### #DIV/0! error with the SUMPRODUCT formula
This happens when the range includes blank cells, since `COUNTIF` returns 0 for a blank and dividing by 0 fails. Shrink the range to exclude blanks, or use the blank-safe version shown in the Excel 2016/2019 steps above.

### UNIQUE returns a #SPILL! error
Excel needs empty cells below and to the side of the formula cell to spill the results. Clear any data in that area, or move the formula to an empty part of the sheet.

### The count includes a header row
If your range starts at row 1 and row 1 is a header like "Customer Name," that header is a text value and gets counted too. Start the range at row 2 instead, for example `A2:A150`.

## Copy-paste setup

Type these formulas directly into a blank cell, adjusting the range to match your data:

- Google Sheets: `=COUNTUNIQUE(A2:A150)`
- Excel 365 / 2021: `=COUNTA(UNIQUE(A2:A150))`
- Excel 2019 / 2016: `=SUMPRODUCT(1/COUNTIF(A2:A150,A2:A150))`
- Excel 2019 / 2016, blank-safe: `=SUMPRODUCT((A2:A150<>"")/COUNTIF(A2:A150,A2:A150&""))`
- To see the list of unique values instead of just a count: `=UNIQUE(A2:A150)` (Excel 365) or `=UNIQUE(A2:A150)` (Google Sheets, which supports the same function name)

Related help pages:
- [COUNTUNIQUE function - Google Docs Editors Help](https://support.google.com/docs/answer/3094289)
- [UNIQUE function - Microsoft Support](https://support.microsoft.com/en-us/office/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e)
- [COUNTIF function - Microsoft Support](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34)
