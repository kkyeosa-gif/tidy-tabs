---
threads_url: https://www.threads.com/@tin_ylab/post/De9QCnTFZ3C
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-auto-categorize-expenses-from.html
title: How to Auto-Categorize Expenses From Bank Descriptions in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python lowercase substring search on all 199 expense rows in 7 edited copies (duplicate keywords, blank keyword rows, a keyword with no category, case changes, blank descriptions, numeric cells, wildcard keywords). Excel and Google Sheets were not opened.
search_description: Auto-categorize bank descriptions in Excel and Google Sheets with INDEX, MATCH and SEARCH. Free keyword template that marks unmatched lines Review.
image_prompts: A small craft studio desk with a stack of folded paper bank statements held by a binder clip, a shoebox of crumpled gas station and post office receipts, and a row of color-coded manila folders, soft window light, no screens, no readable text, no logos
image_alt: Shoebox of crumpled receipts beside colored folders and a clipped stack of statements
threads: Bank export says "SHELL OIL 57442 MANCHESTER NH" and you want it to say Fuel?\nA Keywords tab plus SEARCH inside INDEX/MATCH does it, no rules to click through.\nAnything with no matching keyword shows Review. In the sample, 4 of 25 lines landed there.
---
Put your keywords and categories in a two-column table, then return the category of the first keyword found in each description with `=IF(B2="","",IFERROR(INDEX(Keywords!$B$2:$B$40,MATCH(1,INDEX((Keywords!$A$2:$A$40<>"")*(Keywords!$B$2:$B$40<>"")*ISNUMBER(SEARCH(Keywords!$A$2:$A$40,B2)),0),0)),"Review"))`. A description like "SHELL OIL 57442 MANCHESTER NH" becomes Fuel, and anything with no matching keyword shows Review so you can sort it by hand.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and the results below were recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Expenses sheet with bank descriptions, automatic categories and category totals](images/2026-10-25-team-01-categorize-expenses-by-keyword-search-lookup-excel-google-sheets/categorize-expenses-by-keyword-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: `SEARCH` looks for each keyword inside the description. The first keyword in your table that is found decides the category.

1. Add a sheet named **Keywords**. Type a word that appears in your bank descriptions in **A2:A40**, such as `shell`, and its category in **B2:B40**, such as `Fuel`. Put row 1 headers above them.
2. On a sheet named **Expenses**, paste the date, bank description and amount into **A:C**, starting in row 2.
3. In **D2**, type the formula from the top of this post.
4. Copy **D2** down next to every pasted line.
5. In **F2:F6**, type your category names, and in **G2** type `=SUMIF($D$2:$D$200,F2,$C$2:$C$200)`. Copy it down. Add a row for `Review` too.

How the formula reads:

- `SEARCH(Keywords!$A$2:$A$40,B2)` tries every keyword against the description. It ignores capital letters, so `shell` finds "SHELL OIL" and "shell oil".
- `ISNUMBER(...)` turns each hit into TRUE and each miss into FALSE.
- The two `<>""` tests multiply the result by 0 for any keyword row where the keyword or the category is empty.
- `MATCH(1,...,0)` finds the first row that still equals 1.
- `INDEX(Keywords!$B$2:$B$40,...)` returns the category from that row.
- `IFERROR(...,"Review")` supplies the Review label when nothing matched.
- `IF(B2="","",...)` leaves the category empty on a blank line.

The `INDEX(...,0)` wrapper is the usual way to make this work without pressing Ctrl+Shift+Enter. We did not test that in Excel.

Help pages for the functions: [INDEX in Excel](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd), [MATCH in Google Sheets](https://support.google.com/docs/answer/3093378) and [SUMIF in Excel](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b). Those pages describe the functions, not our test.

### In Excel

The formula should work as typed in the steps above. SEARCH, INDEX, MATCH, IFERROR and SUMIF are old functions, and no newer function is needed.

### In Google Sheets

The same formula should work as typed. Import the template with **File > Import > Upload** if you want to start from our file.

## Example

Sample data is fictional: 25 bank lines for a small online shop in New Hampshire, 09/01/2026 to 09/30/2026. The Keywords tab has 12 keywords.

| Bank description | Amount | Category |
|---|---|---|
| SHELL OIL 57442 MANCHESTER NH | $48.20 | Fuel |
| USPS PO 0245 BOSTON MA 02108 | $18.65 | Postage |
| ETSY INC FEES BROOKLYN NY | $12.40 | Fees |
| Staples Inc 0118 NASHUA NH | $18.20 | Supplies |
| ADOBE *CREATIVE CLOUD | $22.99 | Software |
| SQ *JOES DINER 02108 BOSTON MA | $14.50 | Review |
| CHECK 1042 | $120.00 | Review |

The category totals for all 25 lines:

| Category | Total |
|---|---|
| Fuel | $225.02 |
| Postage | $73.90 |
| Fees | $26.60 |
| Supplies | $171.78 |
| Software | $91.97 |
| Review (4 lines) | $217.95 |
| All amounts | $807.22 |

The check cell shows OK because the six totals add up to $807.22. The 4 Review lines are the diner, a market booth deposit, check 1042 and a Dunkin charge. Add a keyword such as `dunkin` with the category Meals, then add Meals to the totals list, and that line leaves the Review pile.

The categories are your choice. The sheet does not decide what is deductible, so ask your tax preparer about that.

## Troubleshooting

### Every line says Review
Check the Keywords tab first. Both the keyword and the category must be filled in on a row. In LibreOffice, the same formula without the two `<>""` tests returned #N/A for every line as soon as one keyword cell in the range was empty, which `IFERROR` then turned into Review. We did not check how Excel or Google Sheets handle that, so keep the tests.

### A line got the wrong category
Two keywords matched and the one higher in the table won. Move the more specific keyword above the general one on the Keywords tab.

### Everything matched one category
A keyword is too short or a wildcard. `SEARCH` treats `*` and `?` as wildcards, so in LibreOffice a keyword of `*` matched every description. Put `~` before a symbol to match it literally, as in `a~*b`.

### The category totals do not match the total
The check cell shows CHECK. A category name on the Keywords tab is missing from the totals list, an amount is stored as text, or a line has an amount but no description. In the template, a blank description gives an empty category and its amount stays out of the six totals.

### A new keyword does nothing
The formula reads rows 2 to 40 on the Keywords tab and rows 2 to 200 on Expenses. Keep new keywords inside those rows, or widen the ranges in D2 and copy it down again.

## Template

[Download the .xlsx](templates/tidy-tabs-categorize-expenses-by-keyword.xlsx)

The file has three tabs. **Expenses** holds the yellow date, description and amount columns in **A:C** (rows 2 to 200), the category formula in **D**, category totals in **F2:G7**, the total in **G8**, the Review count in **G9** and the check in **G10**. **Keywords** holds the yellow keyword and category table in **A2:B40**, with 12 sample keywords. **How to use** has short notes.

The shop and the bank lines are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
