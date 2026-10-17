---
title: How to Look Up a Value With Two Criteria in Excel and Google Sheets
labels: excel-formulas, small-business, lookup
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and all 9 size and finish pairs were checked against a Python dictionary lookup, plus a pair not on the list, a duplicate pair, a trailing space, lower case text, a price typed as text and a blank size. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google.
search_description: Look up a value with two criteria using a helper key and INDEX MATCH, or SUMIFS, in Excel and Google Sheets. Free price list template.
image_prompts: Three framed art prints in small, medium and large sizes leaning against a studio wall beside a fanned stack of matte and glossy photo paper samples, soft window daylight, no screens, no readable text, no logos
image_alt: Three art prints in graduated sizes leaning on a wall beside paper finish samples
threads: Need a price by size AND finish, but a lookup only checks one column?\nJoin the two into a helper key, =A2&"|"&B2, then run INDEX/MATCH on the key.\nIn the sample sheet Large + Gloss returns $30.00. SUMIFS gives the same number but shows $0.00 when the pair is missing.
---
Join the two lookup columns into a helper key, then match on that key with `=INDEX(Prices!$C$2:$C$30,MATCH(B1&"|"&B2,Prices!$D$2:$D$30,0))`. The helper in **D2** is `=IF(A2="","",A2&"|"&B2)`, filled down the price list.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Price lookup table with sizes and finishes beside a Large Gloss result of $30.00](images/2026-10-18-team-01-lookup-with-two-criteria-index-match-sumifs-excel-google-sheets/lookup-two-criteria-price-lookup-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

Two tabs do the work. **Prices** holds one row per size and finish. **Lookup** holds the two inputs in **B1** (size) and **B2** (finish) and the answer below them.

1. On a tab named **Prices**, type the headers Size, Finish, Price and Key in **A1:D1**.
2. From row 2 down, type one row per size and finish: Large in **A**, Gloss in **B**, 30 in **C**. Format **C** as dollars.
3. In **D2**, type `=IF(A2="","",A2&"|"&B2)`. This joins the two cells into one text key such as Large|Gloss. The `|` keeps the key readable and stops two pairs from running together.
4. Select **D2** and drag the fill handle down to row 30 so you can add rows later. Empty rows stay blank.
5. On a tab named **Lookup**, type Size in **A1** and Finish in **A2**. Type a size in **B1** and a finish in **B2**, for example Large and Gloss.
6. In **B4**, type `=IF(OR($B$1="",$B$2=""),"",IFERROR(INDEX(Prices!$C$2:$C$30,MATCH(B1&"|"&B2,Prices!$D$2:$D$30,0)),"Not on price list"))`. The part that does the work is `INDEX(Prices!$C$2:$C$30,MATCH(B1&"|"&B2,Prices!$D$2:$D$30,0))`. MATCH finds the row where the key equals Large|Gloss, and INDEX returns the price from that row.
7. Change **B1** or **B2** and the price updates.

`IFERROR` replaces the #N/A error with a plain message when the pair is not listed. The leading `IF` leaves the cell blank until both inputs are filled.

### Option 2: SUMIFS, no helper column

For a number, `SUMIFS` checks both columns directly. In **B5**, type:

`=IF(OR($B$1="",$B$2=""),"",SUMIFS(Prices!$C$2:$C$30,Prices!$A$2:$A$30,B1,Prices!$B$2:$B$30,B2))`

It adds the prices of every row that matches both criteria. That works for a price list with one row per pair. It only works for numbers, and it returns 0 instead of an error when nothing matches.

### Option 3: INDEX and MATCH without the Key column

In **B6**, the template joins the two columns inside the formula:

`=IF(OR($B$1="",$B$2=""),"",IFERROR(INDEX(Prices!$C$2:$C$30,MATCH(B1&"|"&B2,INDEX(Prices!$A$2:$A$30&"|"&Prices!$B$2:$B$30,0),0)),"Not on price list"))`

The inner `INDEX(...,0)` makes the joined list work as a normal formula, with no Ctrl+Shift+Enter. This form was run in LibreOffice only. I can't say how Excel or Google Sheets treat it, so use the helper column if you want the safe version.

### In Excel

All three formulas use functions Excel has. Microsoft describes them on its [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) pages.

### In Google Sheets

The same formulas should work as typed. Google documents the [MATCH function](https://support.google.com/docs/answer/3093378), where `0` as the last argument means an exact match. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small print shop with three sizes and three finishes, so the price list has 9 rows.

| Size | Finish | Price | Key |
|---|---|---|---|
| Small | Matte | $12.00 | Small\|Matte |
| Small | Gloss | $14.00 | Small\|Gloss |
| Small | Canvas | $32.00 | Small\|Canvas |
| Medium | Matte | $18.00 | Medium\|Matte |
| Medium | Gloss | $21.00 | Medium\|Gloss |
| Medium | Canvas | $45.00 | Medium\|Canvas |
| Large | Matte | $26.00 | Large\|Matte |
| Large | Gloss | $30.00 | Large\|Gloss |
| Large | Canvas | $64.00 | Large\|Canvas |

With Large in **B1** and Gloss in **B2**, **B4**, **B5** and **B6** each show $30.00, **B7** (the count of matching rows) shows 1, and **B8** shows OK. All 9 pairs gave the listed price in all three formulas, and every one matched a Python dictionary lookup.

## Troubleshooting

### Not on price list, but the pair is on the list
There is probably an extra space, as in "Large " with a trailing space. The key becomes `Large |Gloss`, which matches nothing. Retype the cell. Upper and lower case do not matter: large and gloss still returned $30.00.

### SUMIFS shows $0.00 and the other formulas show a message
The pair is not on the price list, for example Large with Foil. `SUMIFS` returns 0 instead of an error, which looks like a real price. In the template **B8** says Not on price list in this case.

### SUMIFS shows $63.00 and INDEX MATCH shows $30.00
The same pair is listed twice. In the test, a second Large and Gloss row at $33.00 made `SUMIFS` add both, while `MATCH` took the first row. **B7** showed 2, **B8** showed Duplicate rows, and **E** on the Prices tab flagged both rows. Delete or fix one of them.

### The price comes back, but B8 says CHECK
A price was typed as text, such as "30" with quote marks or a text format. `INDEX` returns the text, `SUMIFS` skips it and shows $0.00, and **Prices!E** says Price not a number. Retype the price as a number.

### Everything is blank
**B1** or **B2** is empty. The formulas are built to stay blank, and **B8** says Pick a size and finish.

## Limits

This is a lookup, not pricing advice. It assumes one price per size and finish: duplicates are flagged, not resolved. The two dropdowns hold three fixed names each, so edit their lists for your own sizes. The formulas read rows 2 to 30 only.

## Template

[Download the .xlsx](templates/tidy-tabs-lookup-two-criteria.xlsx)

The file has three tabs. **Prices** has the 9 sample rows in yellow input cells **A2:C10**, the Key in **D** and a Row check in **E**, with formulas ready to row 30. **Lookup** has the size and finish dropdowns in **B1:B2** and the three formulas, the match count and the check in **B4:B8**. **How to use** has short notes, and each tab prints on US Letter.

The prices and shop are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
