---
title: How to Convert Ounces to Pounds and Ounces in Excel and Google Sheets
labels: excel-formulas, shipping, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and 10 packages were checked against Python divmod, plus edge cases for 0, 16, 0.5 and 15.99 oz, a 1,700 oz package, a number typed as text, a word in the weight cell and an empty sheet. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Help.
search_description: Ounces to pounds and ounces: =INT(A2/16) and =MOD(A2,16), or join them as text. A 50 oz package shows as 3 lb 2 oz. Free .xlsx with a check cell.
image_prompts: A small kitchen scale with a kraft shipping box on it, a roll of packing tape and a stack of flat mailers on a packing table, soft window light, no screens, no readable text, no logos
image_alt: Kitchen scale holding a kraft box next to packing tape and flat mailers
threads: Shipping label wants pounds and ounces, but your sheet has 50 oz?\n=INT(A2/16) gives the 3 lb, =MOD(A2,16) gives the 2 oz left over.\nOr join them: =INT(A2/16)&" lb "&MOD(A2,16)&" oz". The template also adds up a whole shipment and checks itself.
---
Get the whole pounds with `=INT(A2/16)` and the leftover ounces with `=MOD(A2,16)`, or join them with `=INT(A2/16)&" lb "&MOD(A2,16)&" oz"`. A 50 oz package shows as 3 lb 2 oz.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Weights table with packed ounces, whole pounds, leftover ounces and a pounds and ounces label for 10 packages, plus a whole shipment total](images/2026-10-14-team-01-convert-ounces-to-pounds-and-ounces-int-mod-excel-google-sheets/ounces-to-pounds-and-ounces-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: 16 ounces make 1 pound. Dividing by 16 and dropping the decimals gives the pounds. The remainder of that division is the ounces left over.

1. In row 1, type these headers: Item, Packed ounces, Pounds, Leftover ounces, Label.
2. Type each item in column **A** and its packed weight in ounces in column **B**, starting in row 2. Use plain numbers such as 50 or 12.25, not "50 oz".
3. In **C2**, type `=IF(B2="","",INT(B2/16))` and fill it down. INT drops everything after the decimal point, so 50 / 16 = 3.125 becomes 3.
4. In **D2**, type `=IF(B2="","",MOD(B2,16))` and fill it down. MOD returns what is left after dividing by 16, so 50 gives 2.
5. In **E2**, type `=IF(B2="","",INT(B2/16)&" lb "&MOD(B2,16)&" oz")` and fill it down. The `&` joins the pieces into text such as 3 lb 2 oz.

The IF part leaves the row empty until you type a weight.

For a whole shipment, add the ounces first and split the total, not the individual pounds. In **H3**, type `=ROUND(SUM(B2:B30),2)`. In **H4**, type `=INT(H3/16)`. In **H5**, type `=ROUND(MOD(H3,16),2)`. The ROUND keeps long decimals such as 14.3399999999997 out of the label.

Add a safety check. In **H7**, type `=IF(ROUND(SUM(C2:C30)*16+SUM(D2:D30)-H3,6)=0,"OK","CHECK")`. It turns every pounds and ounces pair back into ounces and compares the result with the total.

### In Excel

The formulas work as typed. Microsoft describes the two functions on its [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) pages.

### In Google Sheets

The same formulas should work as typed. Google describes them on its [INT help page](https://support.google.com/docs/answer/3093490) and [MOD help page](https://support.google.com/docs/answer/3093497). This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is 10 packages packed by a small online shop.

| Item | Packed ounces | Pounds | Leftover ounces | Label |
|---|---|---|---|---|
| Lavender candle box | 50 | 3 | 2 | 3 lb 2 oz |
| Oat soap bundle | 16 | 1 | 0 | 1 lb 0 oz |
| Earring card | 3.5 | 0 | 3.5 | 0 lb 3.5 oz |
| Canvas tote bag | 22 | 1 | 6 | 1 lb 6 oz |
| Ceramic mug pair | 41 | 2 | 9 | 2 lb 9 oz |
| Wax melt case | 96 | 6 | 0 | 6 lb 0 oz |
| Greeting card set | 12.25 | 0 | 12.25 | 0 lb 12.25 oz |
| Tea towel | 9 | 0 | 9 | 0 lb 9 oz |
| Gift crate | 135 | 8 | 7 | 8 lb 7 oz |
| Print tube | 28 | 1 | 12 | 1 lb 12 oz |

| Result | Cell | Value |
|---|---|---|
| Packages with a weight | H2 | 10 |
| Total ounces | H3 | 412.75 |
| Total pounds | H4 | 25 |
| Leftover ounces | H5 | 12.75 |
| Total label | H6 | 25 lb 12.75 oz |
| Check | H7 | OK |

All 10 rows matched Python `divmod(ounces, 16)`. Exactly 16 oz shows 1 lb 0 oz, not 0 lb 16 oz.

## Troubleshooting

### The label shows #VALUE!
The weight cell holds a word instead of a number. Typing "50 oz" in the cell is the usual cause. In the recalculated file, a word in **B12** gave #VALUE! in C, D and E, and in **H7**. Type only the number, and put the unit in the header.

### The CHECK cell says CHECK
A weight is stored as text. In the recalculated file, "20" typed as text in **B12** was still converted by the row formulas (1 lb 4 oz), but SUM skips text, so **H3** stayed at 412.75 and H7 showed CHECK. Retype the number or change the cell format to Number. This conversion is how LibreOffice behaved; keep weights as real numbers in any app.

### The label has a long decimal such as 2.1000000000000014
Computers store some decimals with tiny errors. Wrap the ounces in ROUND, for example `=ROUND(MOD(B2,16),2)`. The template does this only for the shipment total. In the recalculated file, 50.1 oz gave 3 lb 2.1 oz, and 2,206.34 oz total gave 14.3399999999997 until it was rounded.

### A very heavy package gives a huge pounds number
Nothing caps it. In the recalculated file, 1,700 oz gave 106 lb 4 oz. This sheet only converts units. It does not know any carrier weight limit or rate, so read those on your carrier's own pages.

### A negative weight shows -1 lb 11 oz
INT rounds down, so -5 oz becomes -1 lb and 11 oz. A negative weight is a typo, so fix the cell.

## Template

[Download the .xlsx](templates/tidy-tabs-ounces-to-pounds-ounces.xlsx)

The file has two tabs. **Weights** holds the 10 sample packages in A2:B11, formulas ready for rows 2 to 30 (C to E), and the whole shipment summary in G1:H7. Yellow cells are yours to fill in. **How to use** has short fill-in notes.

The items and weights are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
