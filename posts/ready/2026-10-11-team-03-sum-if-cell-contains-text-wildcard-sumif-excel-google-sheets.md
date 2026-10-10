---
title: How to Sum Cells That Contain Certain Text With SUMIF Wildcards in Excel and Google Sheets
labels: excel-formulas, expense-tracking, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against hand math, including the edge cases listed in the post. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Total every amount whose description contains a word with =SUMIF(B2:B13,"*"&E1&"*",C2:C13). Wildcards, starts with, ends with, COUNTIF and traps. Free .xlsx template.
image_prompts: Small home office desk with a spread of paper receipts and an open shoebox of crumpled receipts, soft window light, no screens, no readable text, no logos
image_alt: Paper receipts spread across a desk beside an open shoebox of crumpled receipts
threads: Need the total of every expense that mentions "Uber"?\n=SUMIF(B2:B13,"*"&E1&"*",C2:C13) adds any row whose description contains the word in E1.\nHeads up: "ink" also catches "Pink ribbon". It matches letters, not words.
---
Wrap the word in asterisks with `=SUMIF(B2:B13,"*"&E1&"*",C2:C13)` to total every amount whose description contains the word in E1. The `*` stands for any text, and the match ignores capital letters.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same function and link to their help pages. Neither Excel nor Google Sheets was opened for this post.

![SUMIF wildcard totals for six search words with contains, starts with and ends with](images/2026-10-11-team-03-sum-if-cell-contains-text-wildcard-sumif-excel-google-sheets/sumif-contains-text-totals-template.png) *Render of the Totals tab of the template, a PDF export from LibreOffice Calc 24.2 using fictional sample data*

## Steps

Set up the data first. Descriptions go in one column, amounts in the next, and the word you want to find goes in its own cell.

1. In row 1, type these headers in A to C: Date, Description, Amount.
2. Type the expenses in **A2:C13**. Descriptions go in **B2:B13** and amounts in **C2:C13** as plain numbers.
3. In **E1**, type the word to find, for example `Uber`.
4. In **F1**, type `=SUMIF(B2:B13,"*"&E1&"*",C2:C13)`.
5. In **G1**, type `=COUNTIF(B2:B13,"*"&E1&"*")` to count how many rows matched.
6. Change the word in E1 and both results update.

Here is how the formula reads. `B2:B13` is where to look. `"*"&E1&"*"` glues an asterisk, the word and another asterisk into one pattern such as `*Uber*`. `C2:C13` holds the amounts to add up for each matching row.

Drop one asterisk to change the rule:

- `E1&"*"` matches text that starts with the word.
- `"*"&E1` matches text that ends with the word.
- A question mark stands for exactly one character, so `Ub?r` matches Uber.

To look for a real `*` or `?`, put a `~` in front of it. The pattern `Stamps~*` finds the text `Stamps*`.

If you will keep adding rows, use a longer range such as `B2:B200` and `C2:C200`, as the template does.

### In Excel

The formulas work as typed. Microsoft documents the arguments and the wildcard characters on its [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) page.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [SUMIF](https://support.google.com/docs/answer/3093583) help page. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is twelve September expenses for a small shop, with the word to find in a separate cell.

| Date | Description | Amount |
|---|---|---|
| 09/01/2026 | Uber to client meeting | $18.40 |
| 09/02/2026 | Printer ink cartridge | $42.99 |
| 09/04/2026 | UBER EATS lunch | $14.25 |
| 09/07/2026 | Pink ribbon spool | $6.50 |
| 09/09/2026 | Etsy listing fees | $3.20 |
| 09/12/2026 | Etsy shipping labels | $27.80 |
| 09/15/2026 | Domain renewal | $12.00 |
| 09/18/2026 | Uber airport | $36.10 |
| 09/21/2026 | Packing tape and ink pen | $9.75 |
| 09/24/2026 | Stamps* | $11.60 |
| 09/26/2026 | Ink refill kit | $19.00 |
| 09/30/2026 | Software subscription | $15.00 |

All twelve add up to $216.59. The template's Totals tab runs six search words through the same formulas:

| Text to find | Contains | Count | Starts with | Ends with |
|---|---|---|---|---|
| Uber | $68.75 | 3 | $68.75 | $0.00 |
| Etsy | $31.00 | 2 | | |
| ink | $78.24 | 4 | $19.00 | |
| Ink pen | $9.75 | 1 | | $9.75 |
| Stamps~* | $11.60 | 1 | | |
| Zoom | $0.00 | 0 | | |

Check Uber by hand: $18.40 plus $14.25 plus $36.10 is $68.75. The $14.25 row is "UBER EATS lunch", so capital letters do not matter. Etsy is $3.20 plus $27.80, which is $31.00.

Every value above matched hand math in LibreOffice Calc.

## Troubleshooting

### The total includes things I did not mean
The match is on letters, not whole words. Searching `ink` also finds "Pink ribbon spool", which is why that row is in the $78.24. Use a longer phrase such as `Printer ink`, or put a space in the pattern.

### A row with the word is left out
`SUMIF` only matches text with a wildcard. If the description cell holds a number, such as 2024, the `*` pattern does not find it. Retype it as text. In the check, a number in a description dropped the Uber total from $68.75 to $50.35.

### The count is higher than the number of amounts added
An amount typed as text, such as "18.40" with quote marks or a stray letter, still counts in `COUNTIF` but `SUMIF` leaves it out. In the check, the count stayed 3 while the total fell to $50.35. Retype the amount as a plain number.

### A blank search cell returns the grand total
An empty E1 turns the pattern into `**`, which matches every row that has text. In the check that gave $216.59 and a count of 12. Fill in E1, or wrap the formula as `=IF(E1="","",SUMIF(B2:B13,"*"&E1&"*",C2:C13))`. On a copy of the template in LibreOffice Calc 24.2, that version showed an empty cell when E1 was blank and $68.75 for Uber, while the plain formula showed $216.59 for the blank E1.

### A search for * or ? gives the wrong rows
Those two characters are wildcards. Type `~*` or `~?` in the search word to match the real character.

## Template

[Download the .xlsx](templates/tidy-tabs-sumif-contains-text.xlsx)

The file has three tabs. **Expenses** holds the twelve sample rows in columns A to C. **Totals** lists six search words in A2:A7, with contains in column B, a count in C, starts with in D and ends with in E, plus the all expenses total in B9. **How to use** has short fill-in notes. The formulas read ranges down to row 200.

The expenses are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
