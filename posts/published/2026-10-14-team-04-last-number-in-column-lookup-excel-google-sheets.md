---
threads_url: https://www.threads.com/@tin_ylab/post/Deg7SmHlAgK
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-get-last-number-in-column-in.html
title: How to Get the Last Number in a Column in Excel and Google Sheets
labels: excel-formulas, small-business, lookup
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python, including blank rows, a text note, a number typed as text, a 0 as the last entry, a new row and an empty log. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Help.
search_description: Last number in a column: =LOOKUP(9.99E+307,B2:B200). A cash log ending at $1,402.75 returns $1,402.75 even with blank rows. Free .xlsx.
image_prompts: Small business cash box with a few folded dollar bills and a paper receipt spike beside a pencil on a wooden counter, soft window light, no screens, no readable text, no logos
image_alt: Cash box with folded bills and a receipt spike on a wooden shop counter
threads: Need the latest balance from a log that keeps growing?\n=LOOKUP(9.99E+307,B2:B200) returns the last number in the column, blank rows and text notes included.\nThe sample cash log ends at $1,402.75 and the formula finds it.\nA number typed as text gets skipped, so the post covers that too.
---
Return the last number in a growing list with `=LOOKUP(9.99E+307,B2:B200)`, and the last date the same way from the date column. In a cash log whose balances end at $1,402.75, the formula returns $1,402.75 even with blank rows in between.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same function as typed. Neither Excel nor Google Sheets was opened for this post.

![Cash log with dates, balances and notes next to a summary showing the latest balance of $1,402.75, the latest date and 7 entries](images/2026-10-14-team-04-last-number-in-column-lookup-excel-google-sheets/last-number-in-column-lookup-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: 9.99E+307 is a number bigger than any balance you will ever type. LOOKUP cannot find it, so it settles on the last number in the range. Blank cells and text are skipped.

1. On a tab named **Log**, type these headers in row 1: Date, Balance, Note.
2. Type each date as MM/DD/YYYY in column **A** and the balance in column **B**. Notes go in column **C**. Blank rows between entries are fine.
3. On a second tab named **Summary**, in **B2**, type `=LOOKUP(9.99E+307,Log!B2:B200)`. This is the latest balance.
4. In **B3**, type `=LOOKUP(9.99E+307,Log!A2:A200)` and format it as a date. Dates are numbers, so the same formula finds the last date.
5. In **B4**, type `=COUNT(Log!B2:B200)`. This counts the entries.
6. For a log that may be empty, wrap the formula: `=IFERROR(LOOKUP(9.99E+307,Log!B2:B200),"No entries")`.

The range stops at row 200. Make it longer if your log will grow past that.

### In Excel

The formula works as typed. Microsoft describes the function on its [LOOKUP function](https://support.microsoft.com/en-us/office/lookup-function-446d94af-663b-451d-8251-369d5e3864cb) page and the wrapper on its [IFERROR function](https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611) page. Format **B3** as a date from **Home > Number Format**.

### In Google Sheets

The same formula should work as typed. Google describes the wrapper on its [IFERROR help page](https://support.google.com/docs/answer/3093304). Format **B3** with **Format > Number > Date**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small shop's cash log with blank rows between some entries.

| Date | Balance | Note |
|---|---|---|
| 09/01/2026 | $1,000.00 | Opening cash |
| 09/03/2026 | $1,085.50 | Saturday market |
| (blank) | | |
| 09/08/2026 | $1,210.25 | Online orders |
| 09/12/2026 | $1,148.00 | Supply order |
| (blank) | | |
| (blank) | | |
| 09/19/2026 | $1,320.40 | Saturday market |
| 09/24/2026 | $1,257.90 | Booth fee |
| (blank) | | |
| 10/02/2026 | $1,402.75 | Online orders |

| Result | Cell | Value |
|---|---|---|
| Latest balance | B2 | $1,402.75 |
| Latest date | B3 | 10/02/2026 |
| Entries | B4 | 7 |
| Latest balance, empty-log version | B5 | $1,402.75 |

LOOKUP returns the last number by position, not the biggest one. In the recalculated copy, a $900.00 entry added at the bottom came back as 900.

## Troubleshooting

### The result is #N/A
There is no number in the range, so LOOKUP has nothing to return. In the recalculated copy, clearing the Log tab gave #N/A in B2 and B3, 0 in B4, and "No entries" in the IFERROR versions. Use the IFERROR form from step 6.

### The result is the old balance, not the new one
The new balance is probably stored as text. In the recalculated copy, "$1,500.00" typed as text in **B14** was skipped and B2 stayed $1,402.75. Text usually lines up on the left of the cell. Retype it as a plain number.

### The date and balance do not match
LOOKUP checks each column on its own. In the same test, the date in **A14** was a real date, so B3 moved to 10/05/2026 while the balance stayed behind. If one of the two is wrong, fix the text entry.

### A note in the balance column is ignored
That is on purpose. A text note such as "check recount" in **B14** did not change the result. This formula finds numbers only. A text version needs a different formula.

### A 0 as the last entry still counts
A typed 0 is a number, not a blank. In the recalculated copy, a 0 typed as the last balance was returned as 0, and B4 counted it. Clear the cell if you want the previous balance back.

## Template

[Download the .xlsx](templates/tidy-tabs-last-number-in-column.xlsx)

The file has three tabs. **Log** holds the 7 sample entries with blank rows between them, ready for rows 2 to 30. **Summary** holds the latest balance, latest date, entry count and the two IFERROR versions. **How to use** has short fill-in notes. Formulas read Log rows 2 to 200.

The dates and balances are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
