---
title: How to Sum Expenses by Category With SUMIF in Excel and Google Sheets
labels: excel-formulas, expenses, freelancers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same SUMIF, SUM and TRIM functions; the links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Sum expenses by category with =SUMIF($C$2:$C$40,F2,$D$2:$D$40), check it against =SUM(D:D), and catch "Supplies " typos. Free .xlsx template.
image_prompts: Close-up of a small freelancer's workspace with a stack of paper receipts held by a binder clip, a pocket calculator, and a few manila folders fanned out on a wooden table, soft window light, all paper text blurred and unreadable, no screens, no logos
image_alt: Stack of paper receipts with a binder clip beside a pocket calculator and folders
threads: Category totals off by $84.50?\nA trailing space makes "Supplies " a different word than "Supplies", so =SUMIF($C$2:$C$40,F2,$D$2:$D$40) skips that row without any error.\nAdd a check: log total minus the category total should be $0.00.
---
Use `=SUMIF($C$2:$C$40,F2,$D$2:$D$40)` next to a list of categories: it adds every amount in column D whose category in column C matches the category in F2. Then compare the category totals to `=SUM(D2:D40)` so a misspelled category shows up as a difference.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Expense log with category totals, percent of total and a check cell that reads OK](images/2026-10-08-team-02-sum-expenses-by-category-sumif-excel-google-sheets/sum-expenses-by-category-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Keep the log in columns A to D and the summary table in F to H. The columns match the template.

1. In row 1 of A to D, type these headers: Date, Vendor, Category, Amount. Log one expense per row.
2. In **F1:H1**, type Category, Total spent, % of total.
3. In **F2:F8**, type each category once: Supplies, Software, Shipping, Travel, Meals, Utilities, Advertising.
4. In **G2**, type `=SUMIF($C$2:$C$40,F2,$D$2:$D$40)`. The first range is where SUMIF looks for the match, the second is what it adds. The dollar signs keep both ranges fixed when you copy down.
5. Fill **G2** down to **G8**.
6. In **F9**, type Total. In **G9**, type `=SUM(G2:G8)`.
7. In **H2**, type `=G2/$G$9` and fill down to **H8**. Format column H as a percentage with one decimal.
8. In **F11**, type Log total, and in **G11** type `=SUM($D$2:$D$40)`. This adds every amount in the log, no matter what the category says.
9. In **F12**, type Difference, and in **G12** type `=ROUND(G11-G9,2)`. It should be $0.00.

The ranges stop at row 40, so you can add rows 24 to 40 without editing a formula. Change 40 to a bigger row number if your log will be longer.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) page.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [SUMIF function](https://support.google.com/docs/answer/3093583) page.

## Example

Sample data below is fictional. It is September 2026 spending for a one-person craft business, 22 expenses in all, with the first on 09/01/2026 for $84.50 at Greenline Office Supply.

The summary table the formulas produce:

| Category | Total spent | % of total |
|---|---|---|
| Supplies | $317.30 | 22.2% |
| Software | $90.99 | 6.4% |
| Shipping | $142.25 | 9.9% |
| Travel | $246.00 | 17.2% |
| Meals | $74.35 | 5.2% |
| Utilities | $240.74 | 16.8% |
| Advertising | $320.50 | 22.4% |
| Total | $1,432.13 | 100.0% |

Check Supplies by hand: $84.50 + $37.80 + $128.60 + $66.40 = $317.30. As a share of the total, that is $317.30 / $1,432.13 = 22.2%.

The log total in **G11** is also $1,432.13, so the difference is $0.00 and the check cell says OK. The percentages shown add up to 100.1% by eye because each one is rounded. The real values add to exactly 100%.

## Troubleshooting

### One category total is too low, and there is no error
Look for a trailing space. In a copy of the template, we changed **C2** from `Supplies` to `Supplies ` (with a space after it). In LibreOffice Calc 24.2.7, the Supplies total dropped from $317.30 to $232.80, which is $84.50 short. The Total fell to $1,347.63 and nothing showed an error. SUMIF treats "Supplies " as a different word, so it skips that row. Excel and Google Sheets were not tested for this.

### The difference cell is not $0.00
This is the check doing its job. The log total in **G11** counts every amount, but the category total in **G9** only counts rows that match the list. The gap is the money in rows with a category that is not in **F2:F8**, usually a typo, a trailing space, or a new category you forgot to add. In the template, **G13** also turns red and says CHECK CATEGORIES.

### I can't see which row is wrong
Sort the log by Category with the Sort command on the **Data** menu. A misspelled value usually lands next to the right one or at the end of the list. For a hidden trailing space, clean the column with `=TRIM(C2)` in a spare column, copy the results, and paste them back as values. Microsoft documents [TRIM](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) and Google documents [TRIM](https://support.google.com/docs/answer/3094140).

### The total is $0.00 for a category I know I used
Check the spelling in **F2:F8** against the log. SUMIF needs the same letters and spaces, so "Shipping" and "Shiping" are different. Also confirm the Amount column holds numbers and not text: left-aligned amounts are a clue. Text amounts were not tested in this template.

### A new category has no total
SUMIF only totals categories in the summary list. Insert a row between rows 2 and 8, type the category, and copy the formulas from the row above so the Total row still includes it.

## Template

[Download the .xlsx](templates/tidy-tabs-sum-expenses-by-category.xlsx)

The file has two tabs. **Expenses** holds the 22-row sample log in A to D, the SUMIF totals and percent of total in F to H, and the Log total, Difference and Check cells in **G11:G13**, with conditional formatting that turns the check red when the categories do not add up. **How to use** has short fill-in notes.

The vendors and amounts are fictional, and the file only adds up what you type. It is not tax or accounting advice, so ask your accountant how to group expenses for filing. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
