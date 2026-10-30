---
threads_url: https://www.threads.com/@tin_ylab/post/DfH3AbXFCX9
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-list-all-orders-for-one-customer.html
title: How to List All Orders for One Customer From a Data List in Excel and Google Sheets
labels: excel-formulas, small-business, lookup
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python filter (27 workbooks, 1,326 cell comparisons). Checked: customers with 0, 1, 11, 12, 13, 14 and 60 orders, a blank pick, a name not in the data, capital letters, a trailing space, wildcard text, a blank amount and a blank row. Excel and Google Sheets were not used, and FILTER was not tested.
search_description: List all orders for one customer in Excel and Google Sheets with COUNTIF, INDEX and MATCH. Free template with a customer dropdown and an orders count.
image_prompts: A stack of paper order slips held by a binder clip next to a small rubber stamp and a pile of kraft shipping boxes on a packing table, soft daylight, no screens, no readable text, no logos
image_alt: Paper order slips clipped together beside a rubber stamp and kraft shipping boxes on a table
threads: VLOOKUP only gives you the first order for a customer?\nNumber the matching rows first: =COUNTIF($B$5:B5,pick) on each row that matches.\nThen INDEX and MATCH fetch match 1, 2, 3 and so on. No FILTER function, so it works in older Excel too.
---
Number each matching row on the Orders sheet with `=IF(AND('Customer view'!$B$1<>"",B5='Customer view'!$B$1),COUNTIF($B$5:B5,'Customer view'!$B$1),"")`, then pull the nth order with `=IFERROR(INDEX(Orders!$A$5:$A$104,MATCH(ROWS($A$4:A4),Orders!$F$5:$F$104,0)),"")`. These use only `COUNTIF`, `INDEX` and `MATCH`, so no `FILTER` function is needed.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python filter. The Excel and Google Sheets steps use the same long-standing functions as typed. Neither Excel nor Google Sheets was opened for this post, and the `FILTER` function was not tested.

![Customer view sheet listing twelve orders for one customer with count and total](images/2026-10-30-team-02-list-all-orders-for-one-customer-countif-index-match-excel-google-sheets/all-orders-for-customer-template.png) *caption: PDF render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

`VLOOKUP` and a single `MATCH` stop at the first hit. The fix is to give every matching row its own running number (1, 2, 3 ...), then look up that number instead of the name.

The template has two tabs. **Orders** holds Order # in **A**, Customer in **B**, Order date in **C**, Item in **D** and Amount in **E**, from row 5 to row 104. **Customer view** holds the picked customer in **B1** and the list from row 5.

### In Excel

1. On **Orders**, type the data in **A5:E104**. Spell each customer the same way every time.
2. On **Customer view**, type a customer name in **B1**.
3. On **Orders**, in **F5**, type `=IF(AND('Customer view'!$B$1<>"",B5='Customer view'!$B$1),COUNTIF($B$5:B5,'Customer view'!$B$1),"")`. The range `$B$5:B5` grows as you copy it down, so the third match gets a 3. The blank check keeps empty rows from matching an empty pick.
4. Select **F5**, then drag the fill handle down to row 104.
5. On **Customer view**, in **A5**, type `=IFERROR(INDEX(Orders!$A$5:$A$104,MATCH(ROWS($A$4:A4),Orders!$F$5:$F$104,0)),"")`. `ROWS($A$4:A4)` is 1 in this row, 2 in the next, and so on, so each row asks for the next match.
6. In **B5**, **C5** and **D5**, type the same formula with `Orders!$C$5:$C$104`, `Orders!$D$5:$D$104` and `Orders!$E$5:$E$104` as the first argument of `INDEX`.
7. Select **A5:D5** and drag down to row 16. That gives room for 12 orders.
8. In **B2**, type `=IF(B1="",0,COUNTIF(Orders!$B$5:$B$104,B1))`. In **B3**, type `=IF(B1="",0,SUMIF(Orders!$B$5:$B$104,B1,Orders!$E$5:$E$104))`.

Past the last match, `MATCH` returns #N/A inside the formula and `IFERROR` turns it into a blank, so no error shows on the sheet. Microsoft's pages: [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a). Those pages describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 8 above, typed exactly as written.
2. Fill **F5** down to row 104 and **A5:D5** down to row 16.

Google's help page: [INDEX](https://support.google.com/docs/answer/3098242). It describes the function, not our test.

### A faster option in newer apps

Current Excel 365 and Google Sheets have a `FILTER` function that returns every matching row from one cell. The template does not use it, and the team did not test it. The method above is for older Excel or for a sheet where you want each step visible.

## Example

Sample data is fictional: 40 orders from 9 customers at a small online shop. Picking Alma Reyes in **B1** gives **B2** = 12 orders found, **B3** = $258.75 and **D1** = 12 rows listed. The check in **D2** says OK. This table shows the first 6 of the 12 rows.

| Order # | Order date | Item | Amount |
|---|---|---|---|
| 5010 | 07/16/2026 | Cedar candle | $36.00 |
| 5011 | 07/19/2026 | Brass hoops | $24.00 |
| 5012 | 07/20/2026 | Cedar candle | $54.00 |
| 5013 | 07/23/2026 | Birthday card | $6.00 |
| 5020 | 08/08/2026 | Birthday card | $6.00 |
| 5024 | 08/16/2026 | Kraft mailer box | $6.50 |

Other picks from the sample: Ben Tran has 7 orders ($178.00), Cora Lind 6 ($68.25), Dev Kapoor 5 ($143.25), Grace Hale 1 ($6.50). Jonah Bell is in the pick-list but has no orders, so the note says "No orders found for this name". The matching rows turn green on the Orders tab.

## Troubleshooting

### Only 12 orders show and D2 says CHECK
The list has 12 rows. A customer with 13 orders is found (**B2** = 13), but only 12 are listed, and **D3** says so. With exactly 12 orders, all show and the check says OK. To make room, copy the last row of results down. If you go past 16 rows, check that the number inside `ROWS` keeps counting.

### Nothing is listed for a name that is in the data
A trailing space makes a different name. In the test, "Alma Reyes " with a space at the end found 0 orders and no warning, only the note "No orders found for this name". The same goes for a space inside one Orders cell: that order was dropped and 11 were found. Retype the name, or clean the data first.

### A typed name with * or ? gives CHECK
`COUNTIF` treats `*` and `?` as wildcards, but the helper column compares exact text. Typing `Al*` made **B2** count 12 while nothing was listed, so **D2** showed CHECK. The check caught it. Pick the full name instead.

### The amount or item shows 0
If a matching row has a blank Amount or Item, `INDEX` returns 0 for it. The total is still right. Fill the blank cell or ignore the 0.

### Capital letters
They do not matter. "alma reyes" and "ALMA REYES" found the same 12 orders as "Alma Reyes".

## Template

[Download the .xlsx](templates/tidy-tabs-all-orders-for-customer.xlsx)

The file has three tabs. **Orders** has 100 slots in **A5:E104**, 40 filled with fictional orders, and the gray helper column **F**. **Customer view** has the customer cell **B1** with a dropdown from the names in **H5:H24**, the counts in **B2** and **B3**, the check in **D2**, and 12 result rows. A typed name is allowed even if it is not in the dropdown. **How to use** has short notes. Yellow cells are inputs and gray cells are formulas.

It prints on US Letter. It was checked in LibreOffice Calc 24.2.7.2 only. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
