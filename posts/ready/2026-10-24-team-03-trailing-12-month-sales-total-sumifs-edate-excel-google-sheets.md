---
title: How to Calculate Trailing 12 Month Sales in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 26 edited variants (end months at a year boundary, windows with 5, 9 and 11 months of data, a missing month, a mid-month end date, a mid-month date in the data, a duplicate month, text in a sales cell, a text date, unsorted rows, a previous 12 total of 0, 199 monthly rows). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Calculate trailing 12 month sales in Excel and Google Sheets with SUMIFS and EDATE. Free template with the previous 12 months for comparison.
image_prompts: A brass shop bell, a stack of folded kraft shopping bags with twine handles and a small potted succulent on a wooden gift shop counter, warm afternoon light, no screens, no readable text, no logos
image_alt: Brass shop bell, folded kraft bags and a small succulent on a gift shop counter
threads: Want last 12 months of sales that does not change every time you reopen the file?\nType the end month in B1, then SUMIFS the months after EDATE(B1,-12) and up to B1.\nThe sample shop totals $67,090.00, up 5.7% on the 12 months before. A missing month shows up as 11 found.
---
Add the 12 months ending at a month you type in B1 with `=SUMIFS(Sales!$B$2:$B$200,Sales!$A$2:$A$200,">"&EDATE(B1,-12),Sales!$A$2:$A$200,"<="&B1)`. Column A of the Sales sheet holds the first day of each month, so B1 must also be a first of the month such as 09/01/2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Sales sheet with a $67,090.00 trailing total, 12 months found and previous 12 total](images/2026-10-24-team-03-trailing-12-month-sales-total-sumifs-edate-excel-google-sheets/trailing-12-month-sales-total-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: pick the last month you want, then total every month that is later than 12 months before it and no later than it. Typing the end month, instead of using today's date, keeps last quarter's report from changing when you reopen the file.

1. On a sheet named **Sales**, type the first day of each month in **A2:A200**, such as 10/01/2025, 11/01/2025 and so on. Format the cells as dates.
2. In **B2:B200**, type that month's total sales as plain numbers.
3. On a second sheet, type the last month you want in **B1**, such as 09/01/2026. It must be the 1st.
4. In **B2**, type the formula from the top of this post.

How the formula reads:

- `EDATE(B1,-12)` is the same day 12 months earlier, 09/01/2025 for the sample.
- `">"&EDATE(B1,-12)` keeps months after that date. It leaves out 09/01/2025 and starts at 10/01/2025.
- `"<="&B1` keeps months up to and including the end month.
- Both conditions test column A, so `SUMIFS` adds the column B amounts for exactly 12 months.

Because the conditions use dates, the rows can be in any order. Start with the first empty row below your data and the range will pick up new months as you type them, up to row 200.

Help pages for the functions: [SUMIFS in Excel](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), [EDATE in Excel](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) and [SUMIFS in Google Sheets](https://support.google.com/docs/answer/3238496). Those pages describe the functions, not our test.

### Compare with the previous 12 months

The template's **B4** moves both dates back another 12 months: `">"&EDATE(B1,-24)` and `"<="&EDATE(B1,-12)`. **B5** subtracts it from the total and **B6** divides that change by B4.

### What the template adds

The built B2 is longer than the bare formula. It shows "Type the month in B1" when B1 is blank or text. The counting core is the same `SUMIFS`.

**B3** counts the months found in the 12 month window with `COUNTIFS`, and **B7** does the same for the previous 12. If either count is below 12, **B5** says "Need 12 months in both" instead of a change, so a short window never looks like a drop in sales. Column **C** on the Sales tab shows a running trailing total next to each month, and stays blank until 12 months of data exist.

### In Excel

The formulas work as typed. SUMIFS, COUNTIFS and EDATE are old functions, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed.

## Example

Sample data is fictional: 36 months of sales for a small gift shop, 10/01/2023 to 09/01/2026. The end month in B1 is 09/01/2026.

| Cell | What it shows | Result |
|---|---|---|
| B1 | End month | 09/01/2026 |
| B2 | Sales 10/2025 to 09/2026 | $67,090.00 |
| B3 | Months found | 12 |
| B4 | Previous 12 (10/2024 to 09/2025) | $63,500.00 |
| B5 | Change in dollars | $3,590.00 |
| B6 | Change in percent | 5.7% |
| B8 | Check | OK |

The sheet is not smooth: December is the biggest month at $12,880.00 and January is the smallest at $3,470.00. A 12 month window always includes one of each, which is why this total moves slowly from month to month.

Change B1 to 08/01/2026 and the window slides one month. The total becomes $66,820.00 and the previous 12 become $63,220.00. Change it to 12/01/2025 and the window crosses a year boundary: $64,910.00 and $61,310.00.

This adds up past sales you typed. It is not a forecast and it does not value a business.

## Troubleshooting

### The total is too low by one month
A month is missing from column A, or it was typed as a mid-month date such as 09/15/2025. **B3** shows 11 instead of 12, and the check in **B8** says CHECK. Type each month as the 1st.

### B8 says CHECK
One of these is true: B1 is not the 1st, a month in column A is not the 1st, a month appears twice, or a cell in column A or B holds text. A duplicate month is counted twice, so B3 shows 13.

### The change shows Need 12 months in both
Either window has fewer than 12 months of data. With an end month of 02/01/2025 in the sample, the total covers only 5 months in the previous window, so no change is shown. Move B1 later or add earlier months.

### The total ignores a sales cell
A cell holding text such as "n/a" is skipped by SUMIFS, so that month adds nothing. B8 says CHECK. Retype it as a number.

### B2 asks for the month
B1 is blank or holds text such as "Sept 2026". Type a real date, 09/01/2026.

## Template

[Download the .xlsx](templates/tidy-tabs-trailing-12-month-sales.xlsx)

The file has three tabs. **Trailing** holds the yellow input in **B1**. Gray cells are formulas: the total in **B2**, months found in **B3**, the previous 12 in **B4**, the changes in **B5** and **B6**, months in the previous 12 in **B7** and the check in **B8**. **Sales** holds the yellow month and sales columns in **A2:B200**, with the running trailing total in column C. **How to use** has short notes.

The shop and the sales are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Trailing tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
