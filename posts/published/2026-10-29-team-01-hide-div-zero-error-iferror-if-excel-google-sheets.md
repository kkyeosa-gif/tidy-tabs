---
threads_url: https://www.threads.com/@tin_ylab/post/DfE-bBpGjjm
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-fix-and-hide-div0-error-in-excel.html
title: How to Fix and Hide the #DIV/0! Error in Excel and Google Sheets
labels: excel-formulas, small-business, error-handling
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model (14 workbooks, 728 cell comparisons). Checked: all Visits zero, blank or text, 0 divided by 0, negative Visits, blank or text Orders, a typo in a formula, and Visits typed as the text "0". Excel and Google Sheets were not used.
search_description: Fix and hide the #DIV/0! error in Excel and Google Sheets with IF and IFERROR. Free template with side by side formulas, an error count and a check that the result column has no errors.
image_prompts: A paper weekly sales ledger with a few rows circled in red pencil, a red pencil and a small stack of shipping labels on a craft table, soft daylight, no screens, no readable text, no logos
image_alt: Red pencil resting on a paper ledger with several rows circled beside shipping labels
threads: Your conversion rate column is full of #DIV/0! again?\nSkip IFERROR if you can: =IF(C5=0,"",B5/C5) blanks only the zero-visits rows.\nIFERROR hides every error, including a wrong cell reference, so a typo in the formula would stay invisible.
---
Hide the error only when the divisor is zero with `=IF(C5=0,"",B5/C5)`, where B5 is Orders and C5 is Visits. `=IFERROR(B5/C5,"")` also works, but it hides every error, including a wrong cell reference.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Weekly rates with #DIV/0! errors in the raw column and blanks in three fixed columns](images/2026-10-29-team-01-hide-div-zero-error-iferror-if-excel-google-sheets/hide-div-zero-errors-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

#DIV/0! means a formula divided by zero or by an empty cell. Here that happens when a week has 0 visits, so the conversion rate (Orders divided by Visits) has no answer. Test the divisor first, then divide.

The layout below matches the template. Orders are in **B5:B16** and Visits are in **C5:C16**, one row per week.

### In Excel

1. In **D5**, type `=B5/C5`. This is the plain division. It shows #DIV/0! when Visits is 0 or empty.
2. In **E5**, type `=IFERROR(B5/C5,"")`. It returns a blank for any error.
3. In **F5**, type `=IF(C5=0,"",B5/C5)`. It returns a blank only when Visits is 0 or empty, so other errors still show.
4. In **G5**, type `=IF(AND(ISNUMBER(B5),ISNUMBER(C5)),IF(C5<>0,B5/C5,""),"")`. It also returns a blank when Orders or Visits is text or empty.
5. Select **D5:G5**, then drag the fill handle down to row 16.
6. Select **D5:G16**, press **Ctrl+1**, choose **Percentage** and set 1 decimal place.

Use column **F** when you want to see surprises, and column **G** when the data may hold text. Microsoft's page: [IFERROR function](https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611). That page describes the function, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 4 above, typed exactly as written.
2. Fill **D5:G5** down to row 16, then choose **Format > Number > Percent**.

Google's help page: [IFERROR function](https://support.google.com/docs/answer/3093304). That page describes the function, not our test.

### Why IFERROR can hide a real mistake

IFERROR replaces every error with the same blank. A typo looks the same as a zero. The template's typo demo uses a name that does not exist, `Visits`, in place of the cell reference C5:

- `=B5/Visits` shows #NAME?.
- `=IFERROR(B5/Visits,"")` shows a blank, so the mistake is hidden.
- `=IF(C5=0,"",B5/Visits)` still shows #NAME?, so you see it.

A blank can mean "no visits this week" or "your formula is broken". With IF, only the first one is blank.

## Example

Sample data is fictional: 12 weeks of orders and visits. This table shows 9 of the weeks. Blank means the formula returned an empty string.

| Week | Orders | Visits | D raw | E IFERROR | F IF zero | G IF number |
|---|---|---|---|---|---|---|
| 1 | 14 | 410 | 3.4% | 3.4% | 3.4% | 3.4% |
| 2 | 9 | 0 | #DIV/0! | blank | blank | blank |
| 3 | 22 | 655 | 3.4% | 3.4% | 3.4% | 3.4% |
| 4 | 0 | 0 | #DIV/0! | blank | blank | blank |
| 5 | 18 | (empty) | #DIV/0! | blank | blank | blank |
| 6 | 11 | n/a | #VALUE! | blank | #VALUE! | blank |
| 7 | 30 | 812 | 3.7% | 3.7% | 3.7% | 3.7% |
| 9 | 0 | 380 | 0.0% | 0.0% | 0.0% | 0.0% |
| 11 | 25 | 0 | #DIV/0! | blank | blank | blank |

The summary is in **I5:J9**. **J5** counts 5 errors in the raw column, **J6** counts 7 rates shown in column G, and **J7** is their average, 2.8%. **J8** averages the raw column and returns #DIV/0! on purpose, because one error spoils an average. **J9** says OK when columns E and G hold no errors and G has a rate for every row with numbers in both Orders and Visits.

Week 4 shows 0 divided by 0, which gives the same error as 9 divided by 0. Week 9 has 0 orders and 380 visits, a real 0.0%, which stays visible. Week 6 has the text "n/a" in Visits. Column F still shows #VALUE! there, which is a different error and a useful warning.

## Troubleshooting

### The average or total at the bottom still shows an error
SUM and AVERAGE return an error if any cell in the range has one. Point them at a fixed column such as G, not at the raw column D. The template's **J7** is `=IF(COUNT(G5:G16)=0,"No rates",AVERAGE(G5:G16))`, which skips blanks and says No rates if nothing is left.

### A formula that returned a blank breaks the next formula
The blank is empty text, not a number. In the test, `=G9*2` on a blank cell gave #VALUE!. Wrap the next formula in IF, or use SUM and AVERAGE, which skip text.

### The rate is blank but the data looks fine
Check for a typo like the one above, then check whether Visits is text. A number stored as text left-aligns. In the test, a Visits cell holding the text "0" gave #DIV/0! in columns D and F, because LibreOffice treated the text as 0 in the division. Column G stayed blank. Retype the cell as a number.

### A blank Orders cell shows 0.0%
A blank counts as 0 in a division. Columns D, E and F showed 0.0% for blank Orders with 50 Visits, and column G stayed blank. Use column G if a blank should not look like a real 0.0%.

### Negative Visits give a negative rate
There is no error to hide, so the rate shows, such as -5. The formulas do not judge whether a number makes sense.

## Template

[Download the .xlsx](templates/tidy-tabs-hide-div-zero-errors.xlsx)

The file has two tabs. **Examples** holds yellow Orders in **B5:B16** and Visits in **C5:C16** with 12 fictional weeks filled in. Gray cells are formulas in **D** to **G** and **J5:J9**. Errors in **D** to **F** show in red. **How to use** has short notes. To add weeks, insert rows inside the table, not below it.

Error names and wording were checked in LibreOffice Calc 24.2.7.2 only. It prints on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
