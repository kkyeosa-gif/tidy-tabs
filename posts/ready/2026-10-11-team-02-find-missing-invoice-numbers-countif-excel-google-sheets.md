---
title: How to Find Missing Invoice Numbers in a Sequence in Excel and Google Sheets
labels: excel-formulas, invoicing, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against hand math, including a text-number case and an out-of-range case. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Find missing invoice numbers with =IF(COUNTIF($A$2:$A$200,D2)=0,"Missing","OK") against a list of expected numbers. Flags duplicates too. Free .xlsx template.
image_prompts: Small office desk with a loose stack of paper invoices fanned out in order with one gap where a sheet is absent, and a rubber stamp beside it, soft window light, no screens, no readable text, no logos
image_alt: Fanned stack of paper invoices with one empty gap beside a rubber stamp
threads: Invoice numbers 1001 to 1012, but you only issued ten?\nList every number you should have, then =IF(COUNTIF($A$2:$A$200,D2)=0,"Missing","OK") flags the gaps.\nThe template also marks a number used twice, so 1009 shows up as Duplicate.
---
List every number you should have, then flag the gaps with `=IF(COUNTIF($A$2:$A$200,D2)=0,"Missing","OK")`, where D2 is one expected number and A2:A200 holds the invoice numbers you actually issued. Any number that returns Missing is a gap to look into.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same function and link to their help pages. Neither Excel nor Google Sheets was opened for this post.

![Invoice numbers 1001 to 1012 with three flagged Missing and one Duplicate](images/2026-10-11-team-02-find-missing-invoice-numbers-countif-excel-google-sheets/missing-invoice-numbers-check-template.png) *Render of the Check tab of the template, a PDF export from LibreOffice Calc 24.2 using fictional sample data*

## Steps

`COUNTIF` counts how many cells in a range equal a value. If an expected number shows up zero times, it is missing. If it shows up twice, it is a duplicate.

This works when your invoice numbers are plain numbers such as 1001, in one column.

1. On a sheet named Invoices, type the header Invoice number in **A1**.
2. Type every invoice number you issued in **A2:A200**, in any order.
3. Add a second sheet named Check. In row 1, type these headers in A to C: Expected number, Times found, Status.
4. In **Check!A2**, type `=MIN(Invoices!$A$2:$A$200)+ROW()-2`. This is the lowest number you issued.
5. Fill **A2** down to row 13. Each row adds one, so the column counts up from the lowest number. Row 2 gives the lowest, row 3 the next one, and so on.
6. In **B2**, type `=COUNTIF(Invoices!$A$2:$A$200,A2)` and fill it down to row 13. This counts how many times each expected number appears on the Invoices sheet.
7. In **C2**, type `=IF(B2=0,"Missing",IF(B2>1,"Duplicate","OK"))` and fill it down to row 13.
8. Select **C2:C13** and add conditional formatting: red fill when the cell equals Missing, yellow when it equals Duplicate.

The `$` signs lock the invoice range, so it stays `$A$2:$A$200` as you fill down. The reference to `A2` has no `$`, so it moves to `A3`, `A4` and so on.

The sheet only looks between your lowest and highest number. A number missing before the first one or after the last one is not detected, because the list starts at the lowest number you entered.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) page. Conditional formatting is under **Home > Conditional Formatting > New Rule**.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [COUNTIF](https://support.google.com/docs/answer/3093480) help page. Conditional formatting is under **Format > Conditional formatting**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is ten invoices from a small studio, numbered 1001 to 1012.

| Invoice number | Client | Amount |
|---|---|---|
| 1001 | Maple Street Bakery | $450.00 |
| 1002 | Dana Ruiz Design | $300.00 |
| 1003 | Harbor Yoga | $180.00 |
| 1005 | Cedar Hill Farm | $620.00 |
| 1006 | Maple Street Bakery | $450.00 |
| 1008 | Harbor Yoga | $180.00 |
| 1009 | Sam Patel LLC | $975.00 |
| 1009 | Sam Patel LLC | $975.00 |
| 1011 | Dana Ruiz Design | $300.00 |
| 1012 | Cedar Hill Farm | $620.00 |

The Check tab lists 1001 to 1012 and reports this:

| Expected number | Times found | Status |
|---|---|---|
| 1003 | 1 | OK |
| 1004 | 0 | Missing |
| 1007 | 0 | Missing |
| 1009 | 2 | Duplicate |
| 1010 | 0 | Missing |
| 1012 | 1 | OK |

Only a few rows are shown. The rest are OK. So 1004, 1007 and 1010 are missing, and 1009 was entered twice.

The template also has a summary in **E1:F8**. It shows lowest number 1001, highest 1012, 12 numbers expected, 10 invoices entered, 3 missing and 1 number used twice. Row 8 checks the math: 10 entered, minus 1 extra copy, plus 3 missing is 12, which equals the 12 expected. In the recalculated file, every value matched a hand check.

## Troubleshooting

### A number you know you entered shows as Missing
It was probably typed as text. In the recalculated file, a text "1005" in the Invoices column was not matched by `COUNTIF`, so 1005 showed Missing and the count of invoices entered dropped to 9. Retype it as a plain number. How Excel and Google Sheets treat text numbers was not tested.

### Invoice numbers with a prefix, like INV-1005, all show Missing
`MIN`, `MAX` and `COUNTIF` need numbers here. A value such as INV-1005 is text, so it is ignored and the expected list cannot be built. Keep the number in its own column and the prefix in another.

### The summary check in F8 does not equal F4
The Check tab lists only 12 numbers. When the real run is longer, such as a highest number of 1020, F4 says 20 numbers are expected but only 12 rows are checked. Fill the formulas in A to C down further and widen the ranges in **F5:F8**.

### The list starts at the wrong number
**A2** starts at the lowest number on the Invoices sheet. A stray low number, such as 1000, moves the start to 1000 and pushes 1012 off the 12 rows. Fix or remove the stray number, or fill down one more row.

### A Missing result does not mean something went wrong
A gap might be a voided invoice, a test invoice or a number skipped on purpose. The sheet shows where the gaps are. It does not say why, and it does not give an accounting or legal answer. Check your invoice records for each one.

## Template

[Download the .xlsx](templates/tidy-tabs-missing-invoice-numbers.xlsx)

The file has three tabs. **Invoices** holds ten sample invoice numbers with a client and an amount. **Check** has the expected numbers in A2:A13, the count in column B, the status in column C with red and yellow highlights, and the summary in E1:F8. **How to use** has short fill-in notes.

The clients and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
