---
title: How to Find Each Product's Best and Worst Month in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 13 workbooks and 5,746 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Return the month name of each product's highest and lowest sales with INDEX, MATCH and MAX or MIN, with ties and blank months handled. Free .xlsx template.
image_prompts: A small gift shop's seasonal shelf with glass jars of wax melts, a wicker basket of glass holiday ornaments and a stack of folded linen tea towels, price tags blank and unreadable, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Wax melts, holiday ornaments and stacked linen tea towels on a shop shelf
threads: Which month did each product sell best?\n=INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0)) reads the month name off the header row. Swap MAX for MIN for the worst month.\nTies go to the first month from the left, so Wax melts shows Jun even though Jul also hit $310.00.
---
Return the month of each product's highest sales with `=INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0))`, where B4:M4 holds the month names and B5:M5 is that product's row, and swap MAX for MIN to get the worst month. In the made-up sample, Soy candle 8 oz is best in Oct at $890.00 and worst in Feb at $380.00, and the download wraps the formula so an empty row says No sales instead of #N/A.

This is the "return the column header of the max value" trick: MAX finds the biggest number in the row, MATCH finds which column it sits in, and INDEX reads the header at the top of that column.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with an independent Python model (13 workbooks, 5,746 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets have the same functions but were not opened for this post; the open questions there are COUNTIF with a decimal MAX criterion, the mmm format on text headers and COUNTA on cells whose formulas return "".

![Monthly product sales grid with best month, worst month and tie count for each row](images/2026-11-14-team-01-find-best-and-worst-month-per-product-column-header-of-max-index-match-excel-google-sheets/best-and-worst-month-per-product-template.png) *Eight products over ten months; Wax melts ties for best in Jun and Jul, and Discontinued soap shows No sales. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Products** tab and a **How to use** tab.

1. Type the month names across row 4: **Jan** in **B4** through **Dec** in **M4**.
2. Type product names in **A5** and down, and each month's sales in **B5:M5** and down. The template holds rows 5 to 54.
3. Leave a month blank if the product was not for sale. Type 0 only if it was for sale and sold nothing.
4. In **O5**, type `=INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0))` for the best month.
5. In **P5**, type `=MAX(B5:M5)` for the best month's sales.
6. In **Q5**, type `=COUNTIF(B5:M5,MAX(B5:M5))` to count how many months tie for best.
7. In **R5**, type `=INDEX($B$4:$M$4,MATCH(MIN(B5:M5),B5:M5,0))` for the worst month, and in **S5**, `=MIN(B5:M5)`.
8. Select **O5:S5** and drag the fill handle down to your last product.

To flip the question and find the top product for each month, go down a column instead of across a row. In **B56**, type `=INDEX($A$5:$A$54,MATCH(MAX(B5:B54),B5:B54,0))` and copy it right to **M56**.

The template wraps these core formulas. **O5** and **R5** show No sales for a row with no numbers and Check when a sales cell holds text. **U5** gives the row's Status (OK, No sales or Check), and **P**, **Q** and **S** fill in only when Status is OK. **O** and **R** are formatted as mmm, so if you type real dates in row 4 instead of names, the results still show Jan, Feb and so on.

Help pages: Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [MATCH](https://support.google.com/docs/answer/3093378). They describe the functions, not our test.

## Example

Sample data is fictional: 8 made-up products with 2026 sales from Jan to Oct, and Nov and Dec still blank.

| Product | Total | Best month | Best sales | Tied for best | Worst month | Worst sales | Status |
|---|---|---|---|---|---|---|---|
| Soy candle 8 oz | $5,295.00 | Oct | $890.00 | 1 | Feb | $380.00 | OK |
| Wax melts | $2,465.00 | Jun | $310.00 | 2 | Feb | $165.00 | OK |
| Holiday ornament | $855.00 | Oct | $640.00 | 1 | Jan | $0.00 | OK |
| Gift box | $1,705.00 | Oct | $720.00 | 1 | Aug | $280.00 | OK |
| Market tote | $1,120.00 | Oct | $175.00 | 1 | Feb | -$25.00 | OK |
| Lip balm | $1,820.00 | Jul | $235.00 | 1 | Jan | $140.00 | OK |
| Linen tea towel | $3,920.00 | May | $565.00 | 1 | Feb | $240.00 | OK |
| Discontinued soap | $0.00 | No sales | | | No sales | | No sales |

Wax melts sold $310.00 in both Jun and Jul, so Best month says Jun (the first from the left) and Tied for best says 2.

Gift box started in July, so Jan to Jun are blank. Blank months are skipped, and its worst month is Aug at $280.00. Holiday ornament has 0 typed for Jan to Aug, and a 0 counts, so its worst month is Jan at $0.00.

Row 56 shows Linen tea towel as the top product in Mar, Apr and May and Soy candle 8 oz in the other months to Oct.

The summary in **W4:X9** shows 8 products listed, Oct as the best month for all products together ($3,305.00 in the Month total row), 1 product with a tie for best, 0 rows showing Check, a grand total of $17,180.00 and CHECK OK.

## Troubleshooting

### A product with no sales shows #N/A
MAX of an empty row is 0, and MATCH cannot find a 0 in a row of blanks. In LibreOffice, the core formula on an all-blank row returned #N/A. The template shows No sales instead.

### The best month ignores a big month
A sale typed as text, such as "$120", is skipped by MAX and MIN without a warning. In LibreOffice, the core formula on 100, "$120" (text) and 90 returned Jan. Retype the amount as a number; the template shows Check on that row.

### Two months tie but only one shows
MATCH stops at the first match from the left, so Wax melts shows Jun and not Jul. Look at Tied for best: anything above 1 means another month has the same total. The template has no tie count for the worst month.

### The worst month is a month the product was not for sale
A 0 was typed in that month. A 0 is a real $0.00 month and can win as the worst, as Holiday ornament's Jan does. Clear the cell if the product was not for sale yet.

## Template

[Download the .xlsx](templates/tidy-tabs-best-month-per-product.xlsx)

**Products** has product names in **A5:A54**, sales for Jan to Dec in **B5:M54**, and the formula columns **N:U** (Total, Best month, Best sales, Months tied for best, Worst month, Worst sales, Months filled in, Status). Row 56 shows the top product for each month, row 57 its sales and row 58 the month total. The summary sits in **W4:X9**. It prints Landscape on US Letter with the header row repeated. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Status is red for Check, and Months tied for best is orange when it is above 1. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
