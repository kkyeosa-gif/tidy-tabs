---
threads_url: https://www.threads.com/@tin_ylab/post/DgGTUzngY1r
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-sort-list-alphabetically-with.html
title: How to Sort a List Alphabetically With a Formula (No SORT Function) in Excel and Google Sheets
labels: excel-formulas, small-business, lookup
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 26 workbooks and 41,705 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Sort names A to Z without SORT: rank with =COUNTIF(names,"<"&name)+COUNTIF(names_so_far,name), then list them with INDEX, MATCH and ROWS.
image_prompts: A clipboard vendor sign-up sheet with handwritten rows lying on a folding table at an outdoor farmers market, next to a metal cash box and a small stack of blank paper booth number cards, morning light under a white canopy tent, realistic photo, no screens, no readable text, no logos
image_alt: Clipboard sign-up sheet on a folding table beside a cash box and booth cards
threads: Stuck on Excel 2016 or 2019 with no SORT function?\nRank each name with =COUNTIF(names,"<"&name), plus a second COUNTIF on the names so far so ties get 5 and 6, not 5 and 5.\nThen INDEX and MATCH pull rank 1, 2, 3 down the column.
---
Give each name its A to Z place with `=COUNTIF(names,"<"&name)+COUNTIF(names_so_far,name)`, where the second COUNTIF breaks ties, then pull place 1, 2, 3 in order with `=INDEX(names,MATCH(ROWS(counter),rank,0))`. In the sample vendor list of 19 names, the download prints a directory that starts 1st Street Coffee, apple Annie's Pies, Apple Annie's Pies, Ash Grove Apiary and carries each booth fee along (1st Street Coffee $45.00), and it re-sorts as soon as you type a new vendor.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (26 workbooks, 41,705 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![A to Z tab listing 19 vendors and booth fees beside Z to A](images/2026-11-22-team-01-sort-list-alphabetically-with-formula-no-sort-function-countif-index-match-excel-google-sheets/alphabetical-sort-formula-template.png) *A to Z tab with 19 sample market vendors sorted A to Z in columns B:C and Z to A in columns E:F, each with its booth fee, starting with 1st Street Coffee at $45.00. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Names go on the **Vendors** tab, one per row in rows 5 to 204, in any order.

1. Type the vendor names in **A5** and down, and a second value (booth fee, booth number) in **B5** and down. Blank rows are fine.
2. In **C5**, type `=IF(A5="","",COUNTIF($A$5:$A$204,"<"&A5)+COUNTIF($A$5:A5,A5))`. The first COUNTIF counts names that come before this one; the second counts this name so far, so two identical names get ranks 5 and 6, not 5 and 5.
3. In **D5**, type the same formula with `">"` in place of `"<"` for a Z to A rank.
4. Fill **C5:D5** down to row 204.
5. On the **A to Z** tab, in **B5**, type `=IFERROR(INDEX(Vendors!$A$5:$A$204,MATCH(ROWS(A$5:A5),Vendors!$C$5:$C$204,0)),"")`. ROWS counts 1, 2, 3 down the column; MATCH finds the name with that rank.
6. In **C5**, type `=IFERROR(IF(INDEX(Vendors!$B$5:$B$204,MATCH(ROWS(A$5:A5),Vendors!$C$5:$C$204,0))="","",INDEX(Vendors!$B$5:$B$204,MATCH(ROWS(A$5:A5),Vendors!$C$5:$C$204,0))),"")` so each fee stays with its vendor and a blank fee stays blank.
7. For Z to A, put the same two formulas in **E5** and **F5**, pointed at `Vendors!$D$5:$D$204`.
8. Fill **B5:F5** down to row 204.

Microsoft's [COUNTIF help](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and Google's [COUNTIF help](https://support.google.com/docs/answer/3093480) cover the criteria syntax, and Microsoft's [MATCH help](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) explains match type 0. The same rank-then-lookup idea is used to [rank products within each category](https://tidytabs.blogspot.com/2026/11/how-to-rank-products-within-each.html).

Newer Excel and Google Sheets have a SORT function. The template does not use it, so it does not need a version with SORT (Excel 2016 and 2019 lack it).

## Example

The vendors are fictional, typed in sign-up order with a blank row 15. The A to Z tab shows:

| # | Vendor (A to Z) | Booth fee | Vendor (Z to A) | Booth fee |
|---|---|---|---|---|
| 1 | 1st Street Coffee | $45.00 | zinnia Flower Co | $30.00 |
| 2 | apple Annie's Pies | $40.00 | Zeke's Honey | $25.00 |
| 3 | Apple Annie's Pies | $40.00 | Zeke's Honey | $25.00 |
| 4 | Ash Grove Apiary | $25.00 | The Candle Barn | $35.00 |
| 5 | Bea's Bees | $25.00 | Oak & Ivy Soap | $30.00 |
| 6 | Bea's Bees | $25.00 | Maple Hollow Farm | $35.00 |
| 7 | Blue Door Bakery | $30.00 | Maple Hollow Farm | $35.00 |
| 8 | blue door bakery | $30.00 | Kettle Corn King | $35.00 |
| 9 | Cedar Ridge Eggs | $20.00 | Dee-Lish Jams | $25.00 |
| 10 | De Luca Pasta | $40.00 | De Luca Pasta | $40.00 |

Rows 11 to 19 continue from Dee-Lish Jams to zinnia Flower Co. The blank row is left out, so 19 names show from 19 entered.

In LibreOffice Calc, upper and lower case counted as the same letter, and a tie kept the order typed: "apple Annie's Pies" was entered first, so it is listed first. A space sorted before letters, so De Luca Pasta comes before Dee-Lish Jams. "The Candle Barn" sorts under T; the formula does not skip "The". Spaces, symbols, accents and case ties follow each app's own text order, so check them in Excel or Sheets.

## Troubleshooting

### A name is missing and a blank row appears in the list
Two names got the same rank. In a 12-name test list, the number 42 and "  Sunny Acres" (leading spaces) both got rank 1, and "*Star* Kettle Corn" and "~Tilde Tea" both got rank 2, because COUNTIF reads * as a wildcard and ~ as an escape character. The A to Z tab showed 10 names and lost "  Sunny Acres" and "~Tilde Tea". **Vendors I6** counts repeated ranks (it showed 4). Remove the extra spaces and symbols, or use the rank in column **F**, `=IF(A5="","",SUMPRODUCT(($A$5:$A$204<>"")*($A$5:$A$204<A5))+SUMPRODUCT(($A$5:A5<>"")*($A$5:A5=A5)))`, which gave 1 to 12 with no gaps, and point the MATCH on the A to Z tab at `Vendors!$F$5:$F$204`.

### A name starting with < or > lands at the start or the end
COUNTIF reads a leading < > or = as a comparison. ">Zed" got rank 200 in LibreOffice. The Status column marks it Check: * ? ~ or < > = in name.

### Booth 10 comes before Booth 2
The names are text, so the order was Booth 1, Booth 10, Booth 2, Booth 20, Booth 3. The formula sorts text, not numbers inside text.

### A row says Check: extra spaces
The name has a space before, after or doubled inside it. Retype it.

## Template

[Download the .xlsx](templates/tidy-tabs-alphabetical-sort-formula.xlsx)

The file has three tabs. **Vendors** takes up to 200 names in rows 5 to 204 with Vendor name and Booth fee (inputs in **A:B**), Rank (A to Z), Rank (Z to A), Status and the SUMPRODUCT rank in **C:F**. The checks in **H4:I9** (names entered, rows showing Check, repeated ranks, rows where F differs from C, names typed below row 204, CHECK) read 19, 0, 0, 0, 0, OK with the sample data.

**A to Z** holds the A to Z list in **B:C** and the Z to A list in **E:F**, numbered in **A**, with checks in **H4:I8** that compare names shown with names entered. It prints Landscape on US Letter. **How to use** explains each formula. For more than 200 names, insert rows above row 204 on both tabs so every range grows.

Yellow cells are inputs and gray cells are formulas. Every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
