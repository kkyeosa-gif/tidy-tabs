---
title: How to Charge Shipping for the First Item and Each Additional Item in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 24 workbooks and 46,842 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Charge one first-item shipping price per order: flag the line with two COUNTIFS, charge each line with IF, and total each order with SUMIF.
image_prompts: A packed shipping box on a home studio packing table holding two wrapped jar candles and a bubble-wrapped ceramic mug, with a small digital postal scale and a roll of brown kraft packing paper beside it, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Open shipping box with two wrapped candles and a mug beside a postal scale
threads: $14.50 is what 2 candles and 1 mug cost to ship when every line pays its own first-item price.\nOnly the mug should pay first, so it is $8.00 + 2 x $1.50 = $11.00.\nTwo COUNTIFS flag that one line per order, and SUMIF adds up the order.
---
Flag the one line per order that pays the first-item price (no line in the order with a higher first-item price, and the first of any ties) with two COUNTIFS, then charge `=IF(flag=1,first+(qty-1)*additional,qty*additional)` on each line and add the lines per order with SUMIF. In the sample order with 2 candles ($5.00 first, $1.50 each additional) and 1 mug ($8.00 first, $3.00 each additional), the mug pays first, so shipping is $8.00 + 2 x $1.50 = $11.00 instead of the $14.50 you get by giving every line its own first-item price.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (24 workbooks, 46,842 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Fourteen order lines with first-item flags, line shipping and order shipping totals](images/2026-11-22-team-02-combined-shipping-first-item-each-additional-item-countifs-sumif-excel-google-sheets/combined-shipping-first-additional-item-template.png) *Orders tab with 7 sample orders and 14 lines, where order 1001 (2 candles and 1 mug) ships for $11.00 and column K shows the $14.50 overcharge. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The sheet's rule: the item with the highest First item price pays it once; every other unit pays its Each additional price. On a tie, the first line listed pays it. Check your shop's shipping settings help for its own rule. The sheet has no carrier rates or tax on shipping.

The steps are the same in Excel and Google Sheets.

1. On the **Profiles** tab, type one item type per row from row 5: the name in **A**, the First item price in **B** and the Each additional price in **C**, such as Small, $5.00, $1.50.
2. On the **Orders** tab, type one order line per row from row 5: Order # in **A**, Item in **B**, Profile in **C** (a dropdown from the Profiles list) and Qty in **D**.
3. In **E5**, look up the first-item price with `=IFERROR(INDEX(Profiles!$B$5:$B$24,MATCH(C5,Profiles!$A$5:$A$24,0))+0,"")`. **F5** is the same formula on `Profiles!$C$5:$C$24`. The +0 turns a blank price into $0.00.
4. In **G5**, flag the line that pays the first-item price with `=IF(AND(D5>0,COUNTIFS($A$5:$A$304,A5,$D$5:$D$304,">0",$E$5:$E$304,">"&E5)=0,COUNTIFS($A$5:A5,A5,$D$5:D5,">0",$E$5:E5,E5)=1),1,0)`. The first COUNTIFS checks that no line in the same order has a higher first-item price. The second, with its growing range, lets only the first of any tied lines through. See the COUNTIFS help from [Microsoft](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [Google](https://support.google.com/docs/answer/3256550).
5. In **H5**, charge the line with `=IF(G5=1,E5+(D5-1)*F5,D5*F5)`. The flagged line pays the first item once plus each additional for its other units; every other line pays each additional for every unit.
6. In **I5**, total the order on its first line only with `=IF(COUNTIF($A$5:A5,A5)=1,SUMIF($A$5:$A$304,A5,$H$5:$H$304),"")`. An order's lines can be anywhere in the list. Microsoft's [SUMIF help](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) explains the range and sum range.
7. Fill **E5:I5** down to row 304. In the template, guards keep blank rows blank and **J** shows a Status.

## Example

The profiles and orders are made up. Profiles: Small $5.00 / $1.50, Fragile $8.00 / $3.00, Flat $4.00 / $0.75, Heavy $12.00 / $6.00.

| Order # | Lines | Pays first | Order shipping | Every line pays first (K) |
|---|---|---|---|---|
| 1001 | Soy candle Small x2, Mug Fragile x1 | Mug | $11.00 | $14.50 |
| 1002 | Art print Flat x3, Sticker pack Flat x1 | Art print | $6.25 | $9.50 |
| 1003 | Soy candle Small x1, Cast iron trivet Heavy x1, Art print Flat x2 | Trivet | $15.00 | $21.75 |
| 1004 | Mug Fragile x2, Vase Fragile x1 | Mug | $14.00 | $19.00 |
| 1005 | Soy candle Small x0, Art print Flat x1 | Art print | $4.00 | $4.00 |
| 1006 | Soy candle Small x1 | Candle | $5.00 | $5.00 |
| 1007 | Vase Fragile x1, Mug Fragile x1 | Vase (listed first) | $11.00 | $16.00 |

In order 1003 the trivet pays $12.00, the candle $1.50 and the two prints $0.75 each: $15.00. The sticker pack for order 1002 sits on row 16 and still lands in its $6.25 total.

The summary in **N4:N12** shows 14 lines, 7 orders, $66.25 in both order totals and line charges, a $0.00 Difference, and CHECK OK.

## Troubleshooting

### The order total is too high
Every line was charged its own first-item price, as column **K** shows on purpose: $14.50 instead of $11.00 for order 1001. Only the line with a 1 in **G** pays the first item.

### The Difference in N10 is not $0.00
An order number may be typed as text, such as 2 entered with an apostrophe next to a real number 2. In LibreOffice Calc, the order total in **I** then left out the text line: order totals summed to $5.00 against $13.00 in line charges, a Difference of -$8.00, and **J** said Check: order # is text. Retype the number (SO-1001 style text is fine). Excel and Google Sheets were not checked.

### A line shows Check: profile not found
The profile has a typo or an extra space, such as "Fragile ". That line gets no charge, so the order total is short. Use the dropdown.

### Two lines have the same first-item price
The first line listed pays it. In order 1007 the vase and the mug are both $8.00, so the vase on row 17 pays $8.00 and the mug pays $3.00, for $11.00.

### The rule gives a different total than my shop
With profile A at $6.00 / $5.00 and profile B at $5.50 / $1.00 in one order, this sheet charges $7.00. A rule that picks the largest gap between the two prices would charge $10.50. That is a different rule; match your shop's settings.

## Template

[Download the .xlsx](templates/tidy-tabs-combined-shipping-calculator.xlsx)

The file has three tabs. **Profiles** holds up to 20 item types in rows 5 to 24 with Profile name, First item $ and Each additional $ in **A:C** and a Status in **D** that flags blank names, extra spaces, repeats and bad prices.

**Orders** holds up to 300 lines in rows 5 to 304. You type **A:D**; **E:J** hold the lookups, the first-item flag, line shipping, order shipping and Status, and **K** shows the every-line-pays-first total for comparison. The summary in **M4:N12** checks that order totals add up to line charges. **How to use** explains each formula. Both data tabs print Landscape on US Letter.

Yellow cells are inputs, gray cells are formulas, and every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
