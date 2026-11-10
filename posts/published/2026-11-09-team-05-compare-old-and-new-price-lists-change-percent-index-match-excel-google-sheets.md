---
threads_url: https://www.threads.com/@tin_ylab/post/Dfj-1-cG7Hc
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-compare-two-price-lists-and-see.html
title: How to Compare Two Price Lists and See What Changed in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 19 workbooks and 720 Compare rows with 0 mismatches left. Checked: the sample, empty lists, all new, all removed, capitals and extra spaces in SKUs, a SKU typed as a number on one list and text on the other, a SKU twice in Old, old and new price 0, a text price, a blank old price, a blank row between SKUs, a SKU twice in New, and four random sets of 30 to 40 SKUs. Excel and Google Sheets were not used. Not run: SKUs containing wildcard characters, more than 40 products per list, a non-US language setting.
search_description: Compare an old and a new price list by SKU with INDEX and MATCH, then show the dollar change, percent change and a status. Free Excel template.
image_prompts: A few small kraft price tags, a pencil and a stack of folded cotton tea towels on a wooden craft fair table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Kraft price tags and a pencil beside folded tea towels on a wooden table
threads: Updated your prices and can't remember what moved?\nPull the old price for each SKU: =IFERROR(INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0)),"")\nIn the sample, CND-8 went from $18.00 to $19.50, up 8.3% (made-up prices).
---
For each SKU on the new list, pull the old price with `=IFERROR(INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0)),"")`, then show the change with `=ROUND(B4-C4,2)` and the percent with `=D4/C4`. In the sample, CND-8 went from $18.00 to $19.50, up $1.50 or 8.3%, and a SKU with no old price is labeled New (made-up prices).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (19 workbooks, 720 Compare rows). The formulas use only IF, IFERROR, INDEX, MATCH, ROUND, TRIM, COUNTIF and SUMPRODUCT, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Compare tab with sample prices, dollar change, percent change and a status for each SKU](images/2026-11-09-team-05-compare-old-and-new-price-lists-change-percent-index-match-excel-google-sheets/compare-old-and-new-price-lists-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. The template has three working tabs. **Old** and **New** each hold a SKU in **A** and a price in **B**, from row 2 down to row 41 (40 products). **Compare** follows the New tab and does the lookups.

1. On **Old**, paste the SKUs in **A** and the prices in **B**. Do the same on **New**.
2. On both tabs, fill **C2** down to row 41 with `=IF(TRIM(A2)="","",TRIM(A2))`. This is the SKU with stray outer spaces removed, so "CND-8 " still matches "CND-8".
3. On **Compare**, type `=IF(New!C2="","",New!C2)` in **A4** and `=IF(A4="","",New!B2)` in **B4**. Compare row 4 reads New row 2.
4. In **C4**, type `=IF(A4="","",IFERROR(INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0)),""))`. `MATCH` finds the row of this SKU on the old list, `INDEX` returns the price from that row, and `IFERROR` leaves the cell blank when the SKU is not there.
5. In **D4**, type `=IF(OR(NOT(ISNUMBER(B4)),NOT(ISNUMBER(C4))),"",ROUND(B4-C4,2))` for the dollar change.
6. In **E4**, type `=IF(D4="","",IF(C4=0,"",D4/C4))` and format the column as a percentage with one decimal. An old price of $0.00 leaves the percent blank instead of showing #DIV/0!.
7. In **F4**, type `=IF(A4="","",IF(G4<>"OK","Check",IF(ISNA(MATCH(A4,Old!$C$2:$C$41,0)),"New",IF(D4>0,"Raised",IF(D4<0,"Lowered","Same")))))`. This is the status.
8. In **G4**, type the row check, which says why a row is flagged:
   `=IF(A4="","",IF(NOT(ISNUMBER(B4)),"Check new price",IF(ISNA(MATCH(A4,Old!$C$2:$C$41,0)),"OK",IF(SUMPRODUCT(--(Old!$C$2:$C$41=A4))>1,"Duplicate in Old",IF(ISNUMBER(INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0))),"OK","Check old price")))))`
9. To list SKUs that disappeared, type `=IF(Old!C2="","",Old!C2)` in **I4**, `=IF(I4="","",Old!B2)` in **J4**, and `=IF(I4="","",IF(ISNUMBER(MATCH(I4,New!$C$2:$C$41,0)),"","Removed"))` in **K4**.
10. Fill rows 4 to 43 down. Count each status with `=COUNTIF($F$4:$F$43,"Raised")`, and count removals with `=COUNTIF($K$4:$K$43,"Removed")`.

Rows are locked to a fixed range (row 41 on the lists, row 43 on Compare). If you have more than 40 products, extend every range to the same last row.

Help pages: Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a). They describe the functions, not our test. We could not load Google's help pages when writing this, so no Google link is given.

## Example

Sample data is fictional: 8 SKUs on each list, with made-up prices.

| SKU | New | Old | Change | Percent | Status |
|---|---|---|---|---|---|
| CND-8 | $19.50 | $18.00 | $1.50 | 8.3% | Raised |
| LIP-1 | $3.50 | $3.50 | $0.00 | 0.0% | Same |
| TOT-1 | $24.00 | $24.00 | $0.00 | 0.0% | Same |
| PIN-1 | $7.50 | $8.00 | -$0.50 | -6.3% | Lowered |
| STK-4 | $4.50 | $4.50 | $0.00 | 0.0% | Same |
| BOX-1 | $34.00 | $32.00 | $2.00 | 6.3% | Raised |
| SOAP-1 | $9.00 | blank | blank | blank | New |
| MUG-1 | $16.00 | $16.00 | $0.00 | 0.0% | Same |

CRD-1 is on the old list but not the new one, so the Removed block shows it. The counts at the top of Compare read Raised 2, Lowered 1, Same 4, New 1, Removed 1, Check 0, and the CHECK cell says OK.

The percent cells hold exact values, so PIN-1 is -6.25% and BOX-1 is 6.25%. The one-decimal format displays both as 6.3%. Use two decimals if that matters to you.

The sheet shows what changed. It does not say what to charge.

## Troubleshooting

### A SKU shows New but it is on the old list
The two spellings differ by more than outer spaces, such as "CND 8" against "CND-8". The template's SKU used column removes outer spaces and collapses doubled spaces, but it does not remove non-breaking spaces or tabs. Retype the SKU on one list.

### Status says Check
Read the **Row check** column. "Check new price" means the new price is text or blank. "Check old price" means the old price is text or blank. A blank old price is not treated as $0.00. "Duplicate in Old" means the SKU appears twice on the old list; the first price is used, so fix the list.

### A SKU appears twice on the new list
Both rows are compared with the same old price. The template does not flag this, so scan the SKU column if the totals look high.

### Percent is blank but the change shows
The old price is $0.00, and dividing by zero would be meaningless. The template leaves the percent empty and still sets the status from the dollar change.

### A tiny change shows Same
Prices with half a cent, such as 1.005 against 1.00, showed Same in the test because of how computers store decimals, where exact decimal math says $0.01 lower. Prices with two decimals are not affected.

### SKUs with * ? or ~
`MATCH` reads these as wildcards. The template does not escape them, so SKUs that contain them are not supported. Capitals do not matter: "cnd-8" matches "CND-8".

## Template

[Download the .xlsx](templates/tidy-tabs-compare-price-lists.xlsx)

The file has four tabs. **Old** and **New** have 40 input rows each with the SKU used helper in column **C**. **Compare** has the counts and CHECK cell in row 2, the table in **A:G**, and the Removed block in **I:K**. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Status cells are shaded, and the sheet prints on US Letter in landscape.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
