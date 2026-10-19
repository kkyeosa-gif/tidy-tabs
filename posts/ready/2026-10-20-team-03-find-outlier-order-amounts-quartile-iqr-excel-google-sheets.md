---
title: How to Find Unusually High or Low Order Amounts in Excel and Google Sheets
labels: excel-formulas, small-business, data-cleanup
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 30 sample orders were checked against a Python model (inclusive quartiles, Decimal), plus 20 edited variants (amounts exactly on a fence, one cent past it, equal amounts, text amounts, a refund, a blank gap, multiplier 3, fewer than 4 amounts). The Review color was seen only in the LibreOffice PDF export and in the saved XML. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Find unusually high or low order amounts with QUARTILE and the 1.5 x IQR rule in Excel and Google Sheets. Free outlier checker for a sales column.
image_prompts: Rows of folded packing slips on a packing table with one slip pulled aside under a round magnifying glass, a handheld price tag gun and a roll of mailing tape beside them, soft natural window light, no screens, no readable text, no logos
image_alt: One packing slip set apart under a magnifying glass among neat rows of slips
threads: Did a $50.00 order get typed as $500.00?\nFlag it with fences built from QUARTILE: above Q3 + 1.5 x IQR or below Q1 - 1.5 x IQR means Review.\nIt also caught a $4.90 slip that a mean plus 2 standard deviations rule missed.
---
Flag an amount for review when it is above Q3 + 1.5 x IQR or below Q1 - 1.5 x IQR. `=QUARTILE(Orders!C2:C100,1)` gives Q1, `=QUARTILE(Orders!C2:C100,3)` gives Q3, and the IQR is Q3 minus Q1. A typed $500.00 among orders of $35.50 to $74.00 lands outside the fences. A mean and standard deviation rule also catches the $500.00, but the typo inflates the standard deviation so much that a $4.90 slip passes it.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The formulas use only QUARTILE, IF, OR, ISNUMBER and COUNTIF, so Excel and Google Sheets should accept them as typed. Neither Excel nor Google Sheets was opened for this post. The amber Review color was seen only in the LibreOffice PDF export.

![Fictional order list with two amber Review rows flagged by quartile fences](images/2026-10-20-team-03-find-outlier-order-amounts-quartile-iqr-excel-google-sheets/outlier-order-amounts-template.png) *PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

This catches typing slips, such as an extra zero, before they reach a monthly total. It does not decide which orders are real. Review means look at the row, not delete it.

1. Make a tab named **Orders**. Type Order ID in **A1**, Date in **B1**, Amount in **C1** and Flag in **D1**. Enter orders from row 2 down, with dates as MM/DD/YYYY.
2. Make a tab named **Fences**. In **A1** to **A7**, type the labels Q1, Q3, IQR, Multiplier, Amounts counted, Upper fence and Lower fence.
3. In **Fences!B1**, type `=IF(COUNT(Orders!$C$2:$C$100)<4,"Need 4 amounts",QUARTILE(Orders!$C$2:$C$100,1))`. In **B2**, type the same formula ending in `,3)`. These are Q1 and Q3. The `IF` shows a message instead of #NUM! when there are fewer than 4 numbers.
4. In **B3**, type `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),B2-B1,"")`. This is the IQR, the spread of the middle half of your amounts.
5. In **B4**, type 1.5. This is the multiplier you can change. A larger number such as 3 flags fewer rows.
6. In **B5**, type `=COUNT(Orders!$C$2:$C$100)`.
7. In **B6**, type `=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B4)),B2+B4*B3,"Check")`. This is the upper fence.
8. In **B7**, type `=IF(AND(ISNUMBER(B1),ISNUMBER(B3),ISNUMBER(B4)),B1-B4*B3,"Check")`. This is the lower fence.
9. In **Orders!D2**, type `=IF(C2="","",IF(NOT(ISNUMBER(C2)),"Check",IF(NOT(AND(ISNUMBER(Fences!$B$6),ISNUMBER(Fences!$B$7))),"Check",IF(OR(C2>Fences!$B$6,C2<Fences!$B$7),"Review","OK"))))`, then drag the fill handle down to row 100.
10. Select **A2:D100**, then add a conditional formatting rule with the formula `=$D2="Review"` and an amber fill.

An amount exactly on a fence counts as OK. One cent past it is Review. A blank row stays blank, and an amount typed as text shows Check.

### In Excel

Select **A2:D100**, then **Home > Conditional Formatting > New Rule > Use a formula to determine which cells to format**. Type `=$D2="Review"`, click **Format**, pick a fill and click **OK**. Microsoft describes the dialog on its [conditional formatting help page](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-in-excel-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f). QUARTILE is the same as [QUARTILE.INC](https://support.microsoft.com/en-us/office/quartile-inc-function-1bbacc80-5075-42f1-aed6-47d735c4819d), the inclusive method used here.

### In Google Sheets

Select **A2:D100**, then **Format > Conditional formatting**. Under **Format cells if**, choose **Custom formula is**, type `=$D2="Review"` and pick a fill color. Google Sheets was not opened for this post, so check the menu names on your screen. Google documents the [QUARTILE function](https://support.google.com/docs/answer/3094041) with the same `QUARTILE(data, quartile_number)` form.

## Example

The sample data below is fictional: 30 orders, ORD-4101 to ORD-4130, dated 09/14/2026 to 10/13/2026. Two amounts are typing slips on purpose. The $500.00 was meant as $50.00 and the $4.90 was meant as $49.00.

| Fences cell | Meaning | Value |
|---|---|---|
| B1 | Q1 | $44.38 |
| B2 | Q3 | $60.91 |
| B3 | IQR | $16.54 |
| B6 | Upper fence | $85.72 |
| B7 | Lower fence | $19.57 |

| Orders row | Order ID | Amount | Flag |
|---|---|---|---|
| 12 | ORD-4111 | $500.00 | Review |
| 22 | ORD-4121 | $4.90 | Review |

The other 28 orders show OK. The mean of all 30 amounts is $65.97 and the standard deviation is $83.07. A mean plus 2 standard deviations rule gives an upper limit of $232.11, so it still catches the $500.00. Its lower limit is -$100.17, so it cannot catch the $4.90. That rule counts 1 amount outside its limits (**Fences!B19**), and the quartile fences count 2 (**Fences!B8**).

**Fences!B12** shows OK when Review, OK and Check add up to the number of rows with an amount.

## Troubleshooting

### Every row says Check
Either the fences are not numbers or the amounts are text. With fewer than 4 numbers, Q1 and Q3 show Need 4 amounts and every amount shows Check. A multiplier typed as text does the same. Add amounts or retype 1.5 as a number.

### One amount says Check and the rest are fine
That amount is text, such as "$50" typed with a leading apostrophe or pasted as a label. Retype it as a plain number. Text amounts are left out of the fences, so they do not skew them.

### A normal order says Review
The fences move when you add or remove amounts, and they follow your typical order size. A real $120.00 order beside mostly $50.00 orders is outside the fences even though nothing is wrong. Check the row against the order itself, or raise **Fences!B4** to 3.

### Almost every amount is the same and small changes get flagged
When most amounts are equal, the IQR is 0 and both fences sit on that amount. In a test with 20 equal amounts, $50.01 and $49.99 both showed Review. If every amount is equal, all of them show OK.

### A refund shows Review
A negative amount such as -$20.00 and a $0.00 row fall below the lower fence of a normal order list. That is expected. Keep refunds on a separate tab if you want clean fences.

## Template

[Download the .xlsx](templates/tidy-tabs-outlier-order-amounts.xlsx)

The file has three tabs. **Orders** holds the 30 sample orders in yellow input cells **A2:C100**, with the Flag formula ready to row 100. **Fences** has Q1, Q3, the IQR, the multiplier in **B4**, the two fences, the counts, the check in **B12** and the mean and standard deviation comparison in **B14:B19**. **How to use** has short notes. All tabs print portrait on US Letter, and the Orders print area is the first 40 rows.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
