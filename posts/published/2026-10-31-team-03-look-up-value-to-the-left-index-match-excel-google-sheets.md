---
threads_url: https://www.threads.com/@tin_ylab/post/DfKpiMklLRL
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-look-up-value-to-left-in-excel.html
title: How to Look Up a Value to the Left in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model of the lookups, in 20 workbooks (7,320 cell comparisons, 0 mismatches). Checked: the first and last product, a SKU not in the list, capital letters, leading, trailing and double spaces, a SKU stored as a number or as text, a duplicate SKU, blank keys, a blank row in the list, names with commas and quotes, wildcard characters and a non-breaking space. The VLOOKUP-to-the-left demo was confirmed to return Not found for every SKU. XLOOKUP was not tried. Excel and Google Sheets were not used.
search_description: Look up a product name from a SKU on its right with INDEX and MATCH, since VLOOKUP cannot look left. Free template with a TRIM and Not found guard.
image_prompts: A small stack of kraft shipping boxes with a SKU sticker on each beside a roll of packing tape and a pencil on a packing table, soft daylight, no screens, no readable text, no logos
image_alt: Stack of shipping boxes with stickers beside a tape roll and pencil
threads: VLOOKUP can't return a column to the left of the one you search.\nINDEX and MATCH can: INDEX picks the column you want, MATCH finds the row.\nThe template puts both side by side, so you can watch VLOOKUP say Not found on every SKU while INDEX/MATCH returns the product name.
---
Use `=INDEX($A$5:$A$44,MATCH(E5,$B$5:$B$44,0))` to return the product name from column A when the SKU you are searching sits in column B. VLOOKUP only returns columns to the right of the search column, so INDEX and MATCH fills the gap in both Excel and Google Sheets.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every result below was recalculated and compared with a Python model. INDEX, MATCH and VLOOKUP are old, standard functions in Excel and Google Sheets, and the steps link to the Microsoft and Google help pages. Neither Excel nor Google Sheets was opened for this post.

![Fifteen SKUs with product name, price, a VLOOKUP-to-the-left column showing Not found, and a summary](images/2026-10-31-team-03-look-up-value-to-the-left-index-match-excel-google-sheets/lookup-to-the-left-index-match-template.png) *caption: the Lookup sheet of the template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

The layout matches the template: **Products** has name, SKU, price and stock in columns A to D, rows 5 to 44. **Lookup** has the SKU you search for in column A from row 5. The steps are the same in Excel and Google Sheets.

1. On **Products**, enter your catalog in **A5:D44**. Keep the SKU column to the right of the name.
2. On **Lookup**, type or paste a SKU in **A5**.
3. In **B5**, type the formula below.
4. Select **B5**, then drag the fill handle down to the last row of SKUs.
5. For the price, copy the formula into **C5** and change `Products!$A$5:$A$44` to `Products!$C$5:$C$44`. Do the same with column D for stock.

The formula in **B5**:

`=IF(A5="","",IFERROR(INDEX(Products!$A$5:$A$44,MATCH(TRIM(A5),Products!$B$5:$B$44,0)),"Not found"))`

`MATCH(TRIM(A5),Products!$B$5:$B$44,0)` finds the row number of the SKU in the SKU column. The `0` means exact match. `INDEX(Products!$A$5:$A$44, row)` returns the cell in that row from the name column. The two ranges are separate, so the name column can be on either side of the SKU column.

`TRIM` removes stray spaces from a pasted SKU, `IFERROR` shows Not found when there is no match, and the `IF` part keeps blank rows blank.

### Why VLOOKUP cannot do this

VLOOKUP searches the first column of its range and returns a column to the right. With the SKU in column B and the name in column A, no range starts at the SKU and also includes the name. Column E of the template tries `=VLOOKUP(A5,Products!$A$5:$D$44,1,FALSE)`. That searches the names in column A, so a SKU is never found, and it shows Not found for all 15 lookups. Column F uses `=VLOOKUP(TRIM(A5),Products!$B$5:$D$44,2,FALSE)`, which works for the price because the price is to the right of the SKU.

Newer Excel and Google Sheets also have XLOOKUP, which can look left. The team did not try it.

Help: Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd), [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [VLOOKUP function](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1), and Google's [function list](https://support.google.com/docs/table/25273), which includes INDEX, MATCH and VLOOKUP.

## Example

Sample data is fictional: a 25 product candle and soap shop and 15 SKU lookups.

| SKU typed | Product name | Price | Stock | VLOOKUP to the left |
|---|---|---|---|---|
| CND-001 | Lavender candle, 8 oz | $18.00 | 24 | Not found |
| SOP-014 | Lemon poppy seed soap | $8.50 | 15 | Not found |
| CND-004 (trailing space) | Citrus grove candle, 8 oz | $18.00 | 9 | Not found |
| snd-010 | Sandalwood reed diffuser | $26.00 | 11 | Not found |
| TRV-003 | Sample card, assorted scents | $5.00 | 70 | Not found |
| ZZZ-999 | Not found | Not found | Not found | Not found |

The summary beside the list shows 15 lookups, 13 found and 2 not found, and the check cell says OK. The SKUs `CND-004 ` with a trailing space and `snd-010` in lower case still match, because of TRIM and because MATCH ignores capital letters in LibreOffice Calc.

## Troubleshooting

### A SKU that is in the list shows Not found
In the tests a non-breaking space after the SKU (common in pasted web text) gave Not found, and so did a SKU typed without its hyphen. TRIM does not remove that kind of space. Retype the SKU in the lookup cell.

### The wrong product comes back
The SKU is listed twice in Products. MATCH returns the first one only. The summary counts the rows involved and the check cell shows CHECK. Remove or rename the duplicate.

### Numeric SKUs do not match
If one list stores 1001 as a number and the other as text, a match can fail in some apps. In LibreOffice Calc both versions matched, but that was not checked in Excel or Google Sheets, so keep SKUs stored the same way in both places.

### The price or stock shows 0
The cell in Products is blank, and INDEX returns 0 for a blank cell. Fill in the value, or wrap the formula in an IF that checks for a blank.

### A lookup below the table is ignored
The template reads rows 5 to 64 on Lookup and rows 5 to 44 on Products. Data typed below those rows is not read, and the check cell shows CHECK.

## Template

[Download the .xlsx](templates/tidy-tabs-lookup-to-the-left.xlsx)

The file has three tabs. **Products** holds 25 fictional products in yellow cells **A5:D44**. **Lookup** has yellow SKU cells in **A5:A64** and gray formulas for name, price, stock, the VLOOKUP demos and the row found, plus a summary and a check cell. **How to use** has short notes.

Limits: 40 products and 60 lookups. Larger lists were not tested. The file was recalculated only in LibreOffice Calc 24.2.7.2. Lookup prints landscape and Products prints portrait, both on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
