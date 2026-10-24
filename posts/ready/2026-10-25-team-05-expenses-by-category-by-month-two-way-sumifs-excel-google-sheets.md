---
title: How to Total Expenses by Category and Month With SUMIFS in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and all 30 category-by-month cells were compared with a Python Decimal sum. Edited copies checked: a header typed as 01/15/2026, a start month that crosses a year, a misspelled category, a date outside the six months, a text date, a text amount, a date with a time, month-end dates, row 500 versus row 501, a blank row, and an empty Expenses tab. Excel and Google Sheets were not opened.
search_description: Total expenses by category and month in one table with SUMIFS and EDATE in Excel and Google Sheets. Free template with 60 sample lines and a check cell.
image_prompts: A row of five glass jars holding coins and folded receipts on a wooden shelf beside a blank corkboard and a pen cup in a home office, soft window light, no screens, no readable text, no logos
image_alt: Five glass jars of coins and folded receipts on a shelf beside a pen cup
threads: No pivot table needed for expenses by category and month.\nCategories down the side, month starts across the top, then one SUMIFS filled across: ">="&B$1 and "<"&EDATE(B$1,1).\nThe sample's 30 cells add up to $2,711.40.
---
Put categories down the side and the first day of each month across the top, then fill one formula across the whole table: `=SUMIFS(Expenses!$C$2:$C$500,Expenses!$B$2:$B$500,$A2,Expenses!$A$2:$A$500,">="&B$1,Expenses!$A$2:$A$500,"<"&EDATE(B$1,1))`. It adds each category's amounts dated within that month, with no pivot table.

The template uses a slightly longer version that wraps each header in `DATE(YEAR(B$1),MONTH(B$1),1)`. That version still works if a header is typed as the 15th. Both are explained below.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Summary sheet with five expense categories across six months, row and column totals and a grand total of $2,711.40](images/2026-10-25-team-05-expenses-by-category-by-month-two-way-sumifs-excel-google-sheets/expenses-by-category-and-month-template.png) *caption: Summary tab of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: for each cell, add the amounts that match the row's category and fall on or after the 1st of the column's month and before the 1st of the next month. Using "before the next 1st" instead of "on or before the last day" also catches dates that carry a time.

1. On a sheet named **Expenses**, type the date in **A2:A500**, the category in **B2:B500** and the amount in **C2:C500**. Format column A as dates and column C as dollars.
2. On a sheet named **Summary**, type your categories in **A2:A6**, spelled exactly as they appear in column B of Expenses.
3. In **B1**, type the first month, such as 01/01/2026. Format it as **mmm yyyy** so it shows Jan 2026.
4. In **C1**, type `=EDATE(B1,1)` and copy it right through **G1**. Each header is one month after the one before.
5. In **B2**, type the formula from the top of this post.
6. Copy **B2** right to **G2**, then copy **B2:G2** down to row 6.

How the formula reads:

- `$A2` locks the category column, so every cell in a row uses that row's category.
- `B$1` locks the header row, so every cell in a column uses that column's month.
- `">="&B$1` keeps dates on or after the 1st of the month.
- `"<"&EDATE(B$1,1)` keeps dates before the 1st of the next month. `EDATE` handles 28, 30 and 31 day months for you.

### The version in the template

The template's B2 is the same formula with each header wrapped in `DATE(YEAR(B$1),MONTH(B$1),1)`:

`=SUMIFS(Expenses!$C$2:$C$500,Expenses!$B$2:$B$500,$A2,Expenses!$A$2:$A$500,">="&DATE(YEAR(B$1),MONTH(B$1),1),Expenses!$A$2:$A$500,"<"&EDATE(DATE(YEAR(B$1),MONTH(B$1),1),1))`

The wrapper turns any date in a month into the 1st of that month. In LibreOffice, a header typed as 01/15/2026 gave the same numbers as 01/01/2026. The shorter formula needs every header to be a 1st.

Help pages for the functions: [SUMIFS in Excel](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), [EDATE in Excel](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) and [SUMIFS in Google Sheets](https://support.google.com/docs/answer/3238496). Those pages describe the functions, not our test.

### Totals and the check

Type `=SUM(B2:G2)` in **H2** and copy it down, then `=SUM(B2:B6)` in **B7** and copy it right through **H7**.

The template also adds a check. **K3** totals every amount on the Expenses tab and **K4** subtracts the table's grand total from it. If K4 is not $0.00, some rows are not in the table.

### In Excel

The formulas work as typed. SUMIFS and EDATE are old functions, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed.

## Example

Sample data is fictional: 60 expense lines for a small online shop, 01/2026 to 06/2026, in five categories.

| Cell | What it shows | Result |
|---|---|---|
| B2 | Fuel, Jan 2026 | $75.00 |
| C2 | Fuel, Feb 2026 | $50.00 |
| H2 | Fuel, six months | $492.00 |
| H3 | Postage, six months | $397.35 |
| H4 | Supplies, six months | $502.35 |
| H5 | Fees, six months | $607.35 |
| H6 | Software, six months | $712.35 |
| H7 | Grand total | $2,711.40 |
| K1 | Biggest month | Jun 2026 |
| K2 | Biggest month total | $507.15 |
| K5 | Check | OK |

January Fuel is two lines, $40.00 and $35.00. February Fuel is a single $50.00 line. The 30 cells matched an independent Python sum line for line.

The table adds up what you typed. It does not say which expenses are deductible. The categories are yours.

## Troubleshooting

### K5 says CHECK and K4 shows missing dollars
A row on the Expenses tab is not landing in any cell. The usual causes are a category spelled differently from A2:A6 (such as "Postge"), a category missing from the table, or a date outside the six months. K4 shows the dollar amount that is missing, which helps you find the row.

### A row with a date is ignored
The date was pasted as text, so it left-aligns and does not compare with the month dates. Retype it as a real date. In the template, a text date also turns K5 to CHECK.

### An amount is ignored
An amount stored as text, such as "$18.00" with a leading apostrophe, is skipped by SUMIFS. K5 shows CHECK. Retype it as a number.

### The last day of the month is missing
You used `"<="&EOMONTH(B$1,0)`, which drops a date with a time on it such as 03/31/2026 11:59 PM. Use `"<"&EDATE(B$1,1)` as above. In the template, that date counted in March.

### New rows are not counted
The formulas read rows 2 to 500. Row 500 is counted and row 501 is not. Extend the ranges in every formula if you have more lines, or start a new sheet for each year.

## Template

[Download the .xlsx](templates/tidy-tabs-expenses-by-category-and-month.xlsx)

The file has three tabs. **Summary** holds the yellow first month in **B1** and the yellow categories in **A2:A6**. Gray cells are formulas: the grid in **B2:G6**, totals in column **H** and row **7**, and the right-hand block with the biggest month in **K1**, its total in **K2**, all Expenses amounts in **K3**, the difference in **K4** and the check in **K5**. **Expenses** holds 60 fictional lines in **A2:C61**, with yellow input cells down to row 500. **How to use** has short notes.

The shop and the amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Summary tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
