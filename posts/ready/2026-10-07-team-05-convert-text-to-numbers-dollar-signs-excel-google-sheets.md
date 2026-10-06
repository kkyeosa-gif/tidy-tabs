---
title: How to Convert Text With Dollar Signs to Numbers in Excel and Google Sheets
labels: excel-formulas, data-cleanup, small-business-finance
tested_in: LibreOffice Calc 24.2 only (Linux, US locale). Not tested in Excel or Google Sheets.
search_description: Pasted $ amounts adding up to $0.00? Convert text to numbers with =VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",","")). Free .xlsx.
image_prompts: A stack of printed bank statement pages with a few amounts circled in pencil, a paper receipt spike with receipts on it, and a pocket calculator on a wooden desk, soft light, all printed text blurred and unreadable, no screens, no logos
image_alt: Printed statement pages with pencil circles, a receipt spike, and a pocket calculator
threads: SUM says $0.00 on a column of pasted dollar amounts?\nThey are text. =VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",","")) turns them into numbers.\nThe ten sample amounts add up to $5,095.60.
---
Strip the symbols and convert: `=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",",""))` turns the text "$1,250.00" into the number 1250. A column of pasted text amounts adds up to $0.00 until you do.

> Works in: LibreOffice Calc 24.2 (Linux, US locale), where the template was built and every result below was recalculated. Excel and Google Sheets use functions with the same names and arguments (links below), but neither app was opened for this post. Other regional settings can read `$`, commas and periods differently, and that was not checked.

![Pasted amounts sheet where text dollar values show FALSE and clean columns sum to $5,095.60](images/2026-10-07-team-05-convert-text-to-numbers-dollar-signs-excel-google-sheets/convert-text-to-numbers-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Say the pasted amounts are in **A2:A11** and `SUM(A2:A11)` shows $0.00. SUM skips text, so every cell in the column is text, not a number.

1. In **B2**, type `=ISNUMBER(A2)` and fill it down. FALSE means the cell holds text. TRUE means it is a real number.
2. In **C2**, type `=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",",""))`.
3. Read the formula from the inside out. TRIM removes stray spaces at the start and end. The first SUBSTITUTE removes the `$`. The second removes the commas. VALUE turns the remaining text into a number.
4. Fill **C2** down to the last row.
5. Format column C as currency. Type `=SUM(C2:C11)` somewhere to get the total.
6. To keep only the numbers, select column C and copy it. Then paste it over itself as values only (**Paste Special > Values**). Delete the original text column and the helper columns.

### Amounts in parentheses

Banks often show a negative as "(45.50)". If your paste has those, the formula below handles them explicitly. Use it in **D2** instead of the formula in step 2:

`=IF(LEFT(TRIM(A2),1)="(",-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",",""),"(",""),")",""))`

It removes the parentheses, then multiplies by -1 when the text started with one.

In LibreOffice Calc 24.2 with a US locale, VALUE already reads "(45.50)" as -45.50, so the shorter formula in step 2 works there too. That was seen in LibreOffice only. Excel and Google Sheets were not tested and may treat parentheses differently, so the longer formula is the safer choice when you do not know. It does not depend on that behavior.

### In Excel

The formulas work as typed. Microsoft documents the arguments on the [VALUE function](https://support.microsoft.com/en-us/office/value-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2) and [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) pages. The paste-as-values step is on the Home tab under **Paste > Values**. This was not tested in Excel here.

### In Google Sheets

Sheets has SUBSTITUTE, VALUE, ISNUMBER and a [TRIM function](https://support.google.com/docs/answer/3094140) with the same arguments. The paste-as-values step is **Edit > Paste special > Values only**. None of this was tested in Sheets here.

## Example

Sample data below is fictional. It is ten amounts pasted from a bank statement and an Etsy payout report.

| Pasted text (column A) | ISNUMBER | Clean with negatives (column D) |
|---|---|---|
| $1,250.00 | FALSE | $1,250.00 |
| 18.00 (leading space) | FALSE | $18.00 |
| (45.50) | FALSE | -$45.50 |
| $89.95 | FALSE | $89.95 |
| $2,400.00 (trailing space) | FALSE | $2,400.00 |
| 12.50 | FALSE | $12.50 |
| $7.25 | FALSE | $7.25 |
| $1,075.40 | FALSE | $1,075.40 |
| $310.00 | FALSE | $310.00 |
| (22.00) | FALSE | -$22.00 |

`=SUM(A2:A11)` on the text column returns $0.00. `=SUM(D2:D11)` returns $5,095.60.

Check it by hand: 1,250.00 + 18.00 - 45.50 + 89.95 + 2,400.00 + 12.50 + 7.25 + 1,075.40 + 310.00 - 22.00 = 5,095.60. In the template, the two cleaned columns gave the same total, and `=ISNUMBER` on the cleaned values is TRUE for all ten rows.

## Troubleshooting

### SUM still shows $0.00 after the formula
The formula column is fine but you are summing the old column. Point SUM at column C or D, not A.

### The cleaned cells are numbers but break when you delete column A
They are still formulas that read column A. Copy them and paste as values only (step 6), then delete the original column.

### #VALUE! in a cleaned cell
VALUE could not read what was left after the substitutions. Look for a letter, an extra symbol, or a space that TRIM does not remove, such as a non-breaking space copied from a web page. Removing it with `SUBSTITUTE(A2,CHAR(160),"")` is a common fix, but that case was not reproduced here.

### The result is off by a factor of 10 or 100
Your regional settings may use a comma as the decimal mark, so "1,250.00" reads differently. Only the US locale was checked here. Look at your system's number settings before you trust the totals.

### You have the opposite problem: ZIP codes lose their leading zeros
That is numbers where you wanted text. See the earlier post "How to Keep Leading Zeros in Excel ZIP Codes" on this blog.

## Template

[Download the .xlsx](templates/tidy-tabs-convert-text-to-numbers.xlsx)

The file has two tabs. **Pasted amounts** holds the ten sample amounts stored as text in column A, with `ISNUMBER` in B, the basic clean formula in C, the clean-with-negatives formula in D, and a second `ISNUMBER` check in E. Three totals sit in G1:H3: the SUM of the text, the SUM of column C and the SUM of column D. **How to use** has short notes on each column.

The amounts are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
