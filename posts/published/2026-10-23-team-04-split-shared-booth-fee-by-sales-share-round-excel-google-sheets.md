---
threads_url: https://www.threads.com/@tin_ylab/post/De4F7BBoML2
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-split-shared-booth-fee-between.html
title: How to Split a Shared Booth Fee Between Sellers by Sales Share in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 32 edited variants (three equal sellers on $100.00, a tie for the largest sales, leftovers of minus one, plus one and minus two cents, a seller with $0.00 sales, all sales 0, negative and text sales, a text total, a total with a fraction of a cent, a $0.05 bill). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Split a shared booth fee or bill between sellers by sales share in Excel and Google Sheets. Free template that keeps the cents adding up.
image_prompts: A stack of small folded market tablecloths, a brass cash tin with a few bills sticking out and a rolled strip of paper tickets on a folding table at a craft fair, soft morning light, no screens, no readable text, no logos
image_alt: Folded tablecloths, an open cash tin and a ticket roll on a craft fair table
threads: Splitting a booth fee by sales share and the cents never add up?\nWrap each share in ROUND, then put the leftover penny on the largest share.\nOn a $240.00 fee, the rounded shares add to $239.99 until that fix. The sample sellers are made up.
---
Split a shared bill by each seller's share of sales with `=ROUND($B$1*B4/SUM($B$4:$B$7),2)`, where B1 is the total and B4:B7 are the sales. Each share is rounded to cents, so a second column adds any leftover penny to the largest share and the parts add back to the total.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Split sheet with a $240.00 booth fee divided among four sellers by sales](images/2026-10-23-team-04-split-shared-booth-fee-by-sales-share-round-excel-google-sheets/split-booth-fee-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: find each seller's share of total sales, multiply it by the bill, round to cents, then fix any penny the rounding lost.

1. In **B1**, type the total bill, such as 240.
2. In **A4:A7**, type the seller names. In **B4:B7**, type each seller's sales for the event. The template has room for 4 sellers.
3. In **C4**, type `=ROUND($B$1*B4/SUM($B$4:$B$7),2)` and fill it down to **C7**.
4. In **C8**, type `=SUM(C4:C7)`. If it does not equal B1, the rounding lost or added a cent or two.
5. In **D4**, type the penny fix below and fill it down to **D7**.
6. In **D8**, type `=SUM(D4:D7)`. It should equal B1.

The penny fix in D4 is:

`=C4+IF(AND(B4=MAX($B$4:$B$7),COUNTIF($B$4:B4,MAX($B$4:$B$7))=1),ROUND($B$1-SUM($C$4:$C$7),2),0)`

Here is how it reads:

- `ROUND($B$1-SUM($C$4:$C$7),2)` is the leftover: the bill minus the rounded shares. It can be a cent or two, plus or minus.
- `B4=MAX($B$4:$B$7)` is true only on the row with the largest sales.
- `COUNTIF($B$4:B4,MAX($B$4:$B$7))=1` keeps the fix to the first row if two sellers tie for the largest sales, so the cent is added once.
- Every other row just repeats its rounded share from column C.

### In Excel

ROUND, SUM, MAX and COUNTIF have been in Excel for many years, so the formulas should work as typed. See Microsoft's [ROUND function help page](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c). That is the help page, not our test.

### In Google Sheets

The same formulas should work as typed. Google's [ROUND help page](https://support.google.com/docs/answer/3093440) describes the function. Again, that is the help page, not a test we ran.

## Example

Sample data is fictional: a made up October craft fair booth fee and four made up sellers.

| Seller | Sales | Share | Rounded | Final |
|---|---|---|---|---|
| Maple Street Candles | $1,310.00 | 42.26% | $101.42 | $101.43 |
| Birch and Bloom Soap | $890.00 | 28.71% | $68.90 | $68.90 |
| Harbor Knit Co. | $560.00 | 18.06% | $43.35 | $43.35 |
| Pine Hill Prints | $340.00 | 10.97% | $26.32 | $26.32 |
| Total | $3,100.00 | 100.00% | $239.99 | $240.00 |

The first row is $240.00 x 1,310 / 3,100 = 101.419..., which rounds to $101.42. The four rounded shares add to $239.99, one cent short of the $240.00 fee. The leftover cent goes to the seller with the largest sales, so the Final column adds to exactly $240.00.

The template also shows an equal split ($60.00 each) and the difference for each seller. That compares two methods. It does not say which is right.

Other cases run in LibreOffice: $100.00 among three sellers with equal sales gives $33.34, $33.33 and $33.33. Sales of $1,200.00, $800.00 and $500.00 on a $100.00 bill give $48.00, $32.00 and $20.00 with no leftover. The leftover can also be negative, so the largest seller pays a cent or two less.

## Troubleshooting

### My rounded amounts add to $239.99, not $240.00
Each share was rounded on its own, and rounding can drop or add a cent. That is what column C does. Column D adds the difference to the largest share. In the template, **J1** shows the leftover and **J2** names the seller who gets it.

### The cell says Check sales
The sales range has text, a negative amount, or adds to 0, or the total in B1 is not a number. The template shows "Check sales" instead of #DIV/0!. Retype the amounts as plain numbers.

### The check cell says CHECK
**J5** compares the Final total with the bill. A bill with a fraction of a cent, such as 86.455, cannot add up to cents, so it shows CHECK. Type the bill to the cent.

### The extra cent went to the biggest seller
That is how the fix works here: the leftover goes to the largest sales, or to the first one listed if two tie. Whether that is acceptable is for the sellers to decide. Agree on the method with the other sellers before you send anyone a number.

### A fifth seller is not counted
The ranges stop at row 7, so the template handles 4 sellers. For more, insert rows inside the range, above row 7, and copy the formulas into them.

## Template

[Download the .xlsx](templates/tidy-tabs-split-shared-cost.xlsx)

The file has two tabs. **Split** holds the yellow inputs (the total in **B1**, seller names in **A4:A7** and sales in **B4:B7**) and the gray formulas: rounded share in **C**, final share in **D**, percent share in **E**, equal split in **F** and the difference in **G**, with totals in row 8. A small summary in **I1:J5** shows the leftover cents, who received them and a check cell. **How to use** has short notes. Split prints landscape on US Letter.

The fee and the sellers are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, not in Excel or Google Sheets. The template does arithmetic only. It does not decide what split is fair or how a booth fee is treated for taxes. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
