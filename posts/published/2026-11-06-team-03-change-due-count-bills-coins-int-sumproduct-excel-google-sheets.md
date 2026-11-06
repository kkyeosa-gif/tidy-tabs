---
threads_url: https://www.threads.com/@tin_ylab/post/DfahxuqFAbJ
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-change-due-and-count.html
title: How to Calculate Change Due and Count Bills and Coins in Excel and Google Sheets
labels: excel-formulas, small-business, cash
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 18 workbooks and 12,150 comparisons with 0 mismatches, and no error value appeared in any recalculated workbook. Checked: the sample, no sales, exact change and a zero sale, 1 cent of change, 60 amounts from $0.00 to $0.99 of change, whole dollar amounts, float traps (1.00 minus 0.30, 10.00 minus 9.99, 5.00 minus 4.35, 20.00 minus 8.20), 4.355 and 4.345 from 5.00, cash short, negative amounts, text, a blank cell, spaces, $10,000 in cash, 60 rows, and six random lists. In every row the pieces added back to the change. Found in a standalone run: a difference of exactly half a cent (2.505 minus 2.50, 5.00 minus 4.995) gives $0.00. Excel and Google Sheets were not used. Not run: more than 60 sales.
search_description: Calculate change due and count the bills and coins to hand back in Excel and Google Sheets with ROUND, INT and SUMPRODUCT. Free template.
image_prompts: An open metal cash box with folded bills and sorted quarters, dimes and nickels in its tray, beside a small stack of kraft paper bags and a bunch of fresh carrots on a wooden farmers market table, soft morning light, realistic photo, no screens, no readable text, no logos
image_alt: Open cash box with folded bills and sorted coins next to paper bags and carrots
threads: Making change at a market stand without a calculator gets old.\nROUND the difference, turn it into cents, then INT and SUMPRODUCT count the pieces.\nA $31.35 sale paid with $50.00 gives one $10, one $5, three $1 bills, two quarters, a dime and a nickel.
---
Get the change with `=ROUND(C8-B8,2)`, turn it into whole cents with `=ROUND(D8*100,0)`, then count each bill or coin with `=INT($E8/F$4)` and the matching formulas in the columns to its right. A $31.35 sale paid with $50.00 gives $18.65 change: one $10, one $5, three $1 bills, two quarters, one dime and one nickel.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. ROUND, INT and SUMPRODUCT exist in current Excel and Google Sheets, but neither app was opened for this post.

![Sales with change due and the bills and coins to hand back for each](images/2026-11-06-team-03-change-due-count-bills-coins-int-sumproduct-excel-google-sheets/change-due-bills-and-coins-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Change** has one sale per row from row 8: sale number in **A**, sale total in **B**, cash given in **C**. Type amounts in whole cents, such as 31.35.

1. In **F4:M4**, type the value of each piece in cents: 2000, 1000, 500, 100, 25, 10, 5, 1. These are the $20, $10, $5 and $1 bills, then quarters, dimes, nickels and pennies. Put the labels in row 7 (**F7:M7**).
2. In **D8**, type `=IF(AND(B8="",C8=""),"",IF(AND(ISNUMBER(B8),ISNUMBER(C8)),IF(AND(B8>=0,C8>=0),IF(ROUND(C8-B8,2)>=0,ROUND(C8-B8,2),"Short"),"Check"),"Check"))`. This is the change due. It shows **Short** when the cash is less than the total and **Check** for text or negative amounts.
3. In **E8**, type `=IF(ISNUMBER(D8),ROUND(D8*100,0),"")`. Whole cents are easier to divide than dollars.
4. In **F8**, type `=IF(ISNUMBER($E8),INT($E8/F$4),"")`. This is the number of $20 bills.
5. In **G8**, type `=IF(ISNUMBER($E8),INT(($E8-SUMPRODUCT($F8:F8,$F$4:F$4))/G$4),"")`. It takes the cents left after the pieces to its left and divides by this column's value.
6. Fill **G8** right to **M8**. Each column widens the range by one: **M8** uses `$F8:L8` and `$F$4:L$4`.
7. In **N8**, type `=IF(ISNUMBER(E8),IF(SUMPRODUCT(F8:M8,$F$4:$M$4)=E8,"OK","Check"),D8)`. It adds the pieces back up and shows OK when they match the change.
8. Fill **D8:N8** down to row 67.
9. In **F5:M5**, type `=SUM(F8:F67)` and fill right. In **D5**, type `=SUM(D8:D67)`. Row 5 is how many of each piece to hand out for all sales.

INT drops the fraction, so each column takes as many whole pieces as fit. The largest values come first, which gives the fewest pieces.

This uses only $20, $10, $5 and $1 bills plus the four coins. It does not know about $2 bills, half dollars, $50 or $100 bills, or what is actually in your cash box.

Help pages: Microsoft's [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e). and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 10 sales at a farmers market stand, totaling $72.71 in change. The table leaves out the $20 column because no sale needed one.

| Sale | Total | Cash given | Change due | $10 | $5 | $1 | Quarters | Dimes | Nickels | Pennies |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | $7.50 | $20.00 | $12.50 | 1 | 0 | 2 | 2 | 0 | 0 | 0 |
| 2 | $12.25 | $20.00 | $7.75 | 0 | 1 | 2 | 3 | 0 | 0 | 0 |
| 3 | $18.00 | $20.00 | $2.00 | 0 | 0 | 2 | 0 | 0 | 0 | 0 |
| 4 | $4.75 | $5.00 | $0.25 | 0 | 0 | 0 | 1 | 0 | 0 | 0 |
| 5 | $23.40 | $40.00 | $16.60 | 1 | 1 | 1 | 2 | 1 | 0 | 0 |
| 6 | $9.99 | $10.00 | $0.01 | 0 | 0 | 0 | 0 | 0 | 0 | 1 |
| 7 | $15.00 | $15.00 | $0.00 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| 8 | $31.35 | $50.00 | $18.65 | 1 | 1 | 3 | 2 | 1 | 1 | 0 |
| 9 | $6.85 | $10.00 | $3.15 | 0 | 0 | 3 | 0 | 1 | 1 | 0 |
| 10 | $8.20 | $20.00 | $11.80 | 1 | 0 | 1 | 3 | 0 | 1 | 0 |

Row 5 shows what to hand out across all ten sales: 0 $20 bills, 4 $10 bills, 3 $5 bills, 14 $1 bills, 13 quarters, 3 dimes, 3 nickels and 1 penny. Check it by hand: 4 x $10 + 3 x $5 + 14 x $1 + 13 x $0.25 + 3 x $0.10 + 3 x $0.05 + $0.01 = $72.71.

The summary shows 10 sales listed, 1 with exact change (sale 7), 0 where cash is short and a largest change of $18.65. The **CHECK** cell says OK.

## Troubleshooting

### The change is $0.00 when it should be $0.01
This happens when the cash and total differ by exactly half a cent, such as 5.00 and 4.995. A spreadsheet stores 0.005 as a number slightly below it, so ROUND goes down. In a standalone run, 2.505 minus 2.50 and 5.00 minus 4.995 both gave $0.00. Type amounts in whole cents and the problem does not come up. In the test, 4.355 from 5.00 gave $0.65 and 4.345 gave $0.66.

### The row says Short
The cash given is less than the sale total. **D** and **N** both show Short, the pieces stay blank and **CHECK** says Check. Fix the cash amount or the total.

### The row says Check
The cell holds text, a negative number, or only one of the two amounts is filled in. In the test, each of these gave Check in **D** and **N**. A stray space in an amount counts as text.

### A piece count looks too high or too low
The count follows the eight values in row 4. If you do not carry quarters, you cannot hand them out, and the sheet does not know that. The sheet only knows the eight values in row 4.

### The CHECK cell says Check
**Q12** says OK when no row shows Check, every sale has both a total and cash, and no cash is short. Look in column **N** for the first row that is not OK.

## Template

[Download the .xlsx](templates/tidy-tabs-change-due-bills-coins.xlsx)

The file has two tabs. **Change** has 60 sale rows from row 8 with 10 fictional sample sales, yellow input cells for the sale total and cash given, the piece values in row 4, the hand-out totals in row 5, and a summary in **Q7:Q12**: sales listed, sales with exact change, sales where cash is short, the largest change due, rows showing Check, and a **CHECK** cell. **How to use** has notes on each formula, including the whole cents tip. The Change tab prints on US Letter paper.

The template has 60 sale rows. A longer list was not tested. Cash up to $99,999.99 was tested.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
