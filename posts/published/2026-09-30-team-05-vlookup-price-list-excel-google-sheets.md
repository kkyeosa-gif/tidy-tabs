---
threads_url: https://www.threads.com/@tin_ylab/post/DeHfIqnljRJ
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-look-up-price-with-vlookup-in.html
title: How to Look Up a Price With VLOOKUP in Excel and Google Sheets
labels: formulas, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Pull a price from a price list by SKU with VLOOKUP exact match, show "SKU not found" with IFERROR, and fix the trailing-space #N/A. Free .xlsx.
image_prompts: A wooden shelf in a small home-based shop holding stacked handmade soap bars and a few candle jars with blank stickers, a wooden clipboard with a printed price sheet leaning against the shelf, all text blurred and unreadable, no logos, soft daylight
image_alt: Soap bars and candle jars on a shelf beside a clipboard price sheet
threads: VLOOKUP says SKU not found but the SKU is right there?\nCheck for a space at the end. FALSE means exact match, so SOP-OAT plus a trailing space does not match SOP-OAT.\nWrap the lookup in IFERROR so the cell says SKU not found instead of a raw error.
---
Type `=VLOOKUP(A2,'Price list'!$A$2:$C$500,3,FALSE)` to pull the price for a SKU; FALSE forces an exact match, and wrapping it in IFERROR shows a clear message when the SKU is missing.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template's formulas were recalculated and checked. The Excel and Google Sheets steps below follow Microsoft Support and Google Docs Editors Help. Neither app was hands-on tested here.

![LibreOffice Calc order lines sheet showing SKUs, quantities, looked-up prices, and line totals](images/2026-09-30-team-05-vlookup-price-list-excel-google-sheets/vlookup-price-list-template.png) *LibreOffice Calc 24.2 PDF export of the Order lines sheet with sample data (fake).*

## Steps

Set up two tabs. **Price list** holds your products. **Order lines** is where you type a SKU and a quantity and let the formulas fill in the rest.

1. On the **Price list** tab, put **SKU** in column A, **Item** in column B, and **Price** in column C. VLOOKUP searches the first column of the range, so SKU must come first.
2. On the **Order lines** tab, put **SKU** in column A and **Qty** in column B. Leave C, D, and E for formulas.
3. In C2, look up the item name: `=IFERROR(VLOOKUP(A2,'Price list'!$A$2:$C$500,2,FALSE),"SKU not found")`. The 2 means "return column 2 of the range", which is Item.
4. In D2, look up the unit price: `=IFERROR(VLOOKUP(A2,'Price list'!$A$2:$C$500,3,FALSE),"")`. The 3 means column 3, which is Price. The empty quotes leave the cell blank when the SKU is missing.
5. In E2, calculate the line total: `=IF(D2="","",B2*D2)`.
6. Format D and E as currency, then copy C2:E2 down for every order line.
7. Add an order total somewhere outside the table, for example `=SUM(E2:E500)` in I1.

The `$` signs in `$A$2:$C$500` lock the range, so it does not slide down when you copy the formula. The range runs to row 500 so you can add products without editing the formula.

### In Excel

1. Click C2 and type the formula from step 3.
2. In the last argument, Range_lookup, enter `FALSE` for an exact match.
3. Press **Enter**, then copy the formula down with the fill handle.

Microsoft documents the arguments on its [VLOOKUP function page](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1) and covers the error wrapper on its [IFERROR function page](https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611).

### In Google Sheets

1. Click C2 and type the formula from step 3. Sheets suggests the function name as you type.
2. Use `FALSE` for the fourth argument (Google calls it is_sorted), so Sheets looks for an exact match.
3. Press **Enter**, then drag the fill handle down to copy the formula.

Google describes the arguments on its [VLOOKUP help page](https://support.google.com/docs/answer/3093318?hl=en).

## Example

Sample data is fictional. The price list has six products:

| SKU | Item | Price |
|---|---|---|
| CND-LAV8 | Lavender soy candle, 8 oz | $18.00 |
| CND-CED8 | Cedar soy candle, 8 oz | $18.00 |
| SOP-OAT | Oatmeal goat milk soap bar | $8.00 |
| SOP-CHR | Charcoal soap bar | $8.50 |
| EAR-HOOP | Brass hoop earrings | $24.00 |
| CRD-BDAY | Letterpress birthday card | $6.00 |

Six order lines go through the formulas:

| Row | SKU typed | Qty | Item | Unit price | Line total |
|---|---|---|---|---|---|
| 2 | CND-LAV8 | 2 | Lavender soy candle, 8 oz | $18.00 | $36.00 |
| 3 | SOP-OAT | 3 | Oatmeal goat milk soap bar | $8.00 | $24.00 |
| 4 | EAR-HOOP | 1 | Brass hoop earrings | $24.00 | $24.00 |
| 5 | SOP-OAT (trailing space) | 1 | SKU not found | blank | blank |
| 6 | CRD-BDAY | 5 | Letterpress birthday card | $6.00 | $30.00 |
| 7 | CND-CED8 | 1 | Cedar soy candle, 8 oz | $18.00 | $18.00 |

Order total: **$132.00**. Row 5 is left out of the total because it has no price. These values match what LibreOffice Calc 24.2.7 calculated from the template, and they match a hand calculation.

## Troubleshooting

### A SKU that is clearly in the list shows "SKU not found"
Look at row 5 above. The typed SKU is `SOP-OAT ` with a space at the end, so it is not an exact match for `SOP-OAT`, and FALSE demands an exact match. Delete the extra space. Pasted data from a web page or a CSV export often carries these spaces, and they are invisible in the cell.

### You see #N/A instead of a message
Without IFERROR, VLOOKUP returns #N/A whenever it finds no match. Wrap the formula as `=IFERROR(VLOOKUP(...),"SKU not found")`. Be aware that IFERROR also hides other errors, so a wrong range would show the same message.

### The formula returns the wrong item or price
Check the column number. The 2 and the 3 count columns from the left edge of the range, not from column A of the sheet. If you start the range at B, column 3 is one column further right than you expect.

### The price is right in row 2 and wrong or missing lower down
The range was not locked. Use `$A$2:$C$500` with dollar signs so every copied row searches the same block. Without them, the range shifts down one row per row copied and eventually skips products.

### Your SKUs are not in the first column
VLOOKUP only looks right of the column it searches. Use INDEX and MATCH instead: `=IFERROR(INDEX('Price list'!$C$2:$C$500,MATCH(A2,'Price list'!$A$2:$A$500,0)),"")`. The 0 means exact match. In the template, column F uses this formula, and it returned the same prices as the VLOOKUP column in every row of the sample.

## Template

[Download the .xlsx](templates/tidy-tabs-vlookup-price-list.xlsx)

The file has three tabs. **Order lines** has the SKU and Qty columns, the VLOOKUP formulas for Item, Unit price, and Line total, an INDEX/MATCH price check in column F, and the order total in I1. **Price list** has the six sample products. **How to use** has short notes on each formula.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
