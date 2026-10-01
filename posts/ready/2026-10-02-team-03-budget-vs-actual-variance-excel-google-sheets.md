---
title: How to Create a Budget vs Actual Template in Excel and Google Sheets
labels: excel-formulas, budgeting, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Build a budget vs actual template: dollar variance with =C2-B2, percent variance with IF to skip $0.00 budgets, and red over-budget rows. Free .xlsx.
---
Subtract budget from actual with `=C2-B2` for the dollar variance, and use `=IF(B2=0,"",(C2-B2)/B2)` for the percent. The IF leaves the percent blank when the budget is $0.00 instead of showing a divide-by-zero error.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every variance, percent, and total below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

## Steps

You need one row per spending category, with the budget and the actual amount side by side. Put the headers in row 1: Category, Budget, Actual, Variance ($), Variance (%), Status.

1. Type your categories in column A, starting in **A2**.
2. Type the amount you planned to spend in column B (Budget) and the amount you really spent in column C (Actual).
3. In **D2** (Variance $), type `=C2-B2`. Actual minus budget means a positive number is overspending and a negative number is money saved.
4. In **E2** (Variance %), type `=IF(B2=0,"",(C2-B2)/B2)`. This divides the dollar difference by the budget. When the budget is $0.00, the cell stays blank.
5. In **F2** (Status), type `=IF(C2>B2,"Over budget","On or under")`.
6. Fill D2:F2 down for every category row.
7. Format B, C, and D as currency with two decimals, and column E as a percentage with one decimal.
8. Put the totals beside the table, not under it, so you can add rows without moving them. In **H1** to **H4**, type Total budget, Total actual, Total variance ($), and Total variance (%).
9. In **I1**, type `=SUM(B2:B200)`. In **I2**, type `=SUM(C2:C200)`. In **I3**, type `=I2-I1`. In **I4**, type `=IF(I1=0,"",(I2-I1)/I1)`.
10. Select **A2:F200** and add a conditional formatting rule that turns a row red when actual is above budget. The formula is `=AND($C2<>"",$C2>$B2)`.

The `$C2<>""` part keeps empty rows from turning red. The dollar signs lock the columns so the whole row changes color, not just one cell.

### In Excel

1. Select **A2:F200**.
2. Go to **Home > Conditional Formatting > New Rule**.
3. Choose **Use a formula to determine which cells to format**.
4. Paste `=AND($C2<>"",$C2>$B2)`, click **Format**, pick a red fill, and click **OK**.

Microsoft documents the same dialog on its [conditional formatting help page](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-in-excel-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f).

### In Google Sheets

1. Select **A2:F200**.
2. Go to **Format > Conditional formatting**.
3. Under **Format cells if**, choose **Custom formula is**.
4. Paste `=AND($C2<>"",$C2>$B2)`, pick a red fill under **Formatting style**, and click **Done**.

Google describes these panels on its [conditional formatting help page](https://support.google.com/docs/answer/78413). The formulas in steps 3 to 9 work as typed in both apps.

## Example

Sample data below is fictional: one month of costs for a small craft booth business.

| Category | Budget | Actual | Variance ($) | Variance (%) | Status |
|---|---|---|---|---|---|
| Booth and market fees | $450.00 | $450.00 | $0.00 | 0.0% | On or under |
| Materials | $600.00 | $683.40 | $83.40 | 13.9% | Over budget |
| Shipping | $220.00 | $241.75 | $21.75 | 9.9% | Over budget |
| Packaging | $120.00 | $131.20 | $11.20 | 9.3% | Over budget |
| Software subscriptions | $85.00 | $85.00 | $0.00 | 0.0% | On or under |
| Advertising | $150.00 | $92.50 | -$57.50 | -38.3% | On or under |
| Payment processing fees | $90.00 | $87.35 | -$2.65 | -2.9% | On or under |
| Training (not budgeted) | $0.00 | $45.00 | $45.00 | blank | Over budget |

The totals block beside the table then shows:

| Total | Amount |
|---|---|
| Total budget | $1,715.00 |
| Total actual | $1,816.20 |
| Total variance ($) | $101.20 |
| Total variance (%) | 5.9% |

The Training row is the zero-budget case. The dollar variance is $45.00 and Status says Over budget, but the percent is blank because $45.00 divided by $0.00 has no answer. The four over-budget rows turn red. Booth fees and software land exactly on budget, so they stay plain.

## Troubleshooting

### Variance (%) shows #DIV/0!
The budget cell is 0 and the formula has no guard. Use `=IF(B2=0,"",(C2-B2)/B2)` so the percent stays blank for a $0.00 budget.

### The percent shows 0.139 instead of 13.9%
Column E is formatted as a number. Select it and change the format to Percent with one decimal place.

### No rows turn red
The rule has the wrong cell references. Select the range, open the rule, and check that the formula starts from the first row of the range (row 2) and has the dollar signs in `$C2>$B2`. Budget or Actual typed as text, such as an amount with a leading apostrophe, also blocks the comparison. Retype it as a plain number.

### Only one cell turns red, not the whole row
The formula is missing the dollar signs, so the comparison moves right along with each column. Use `$C2` and `$B2`, with the dollar sign before the column letter only.

### New categories are missing from the totals
The totals add rows 2 to 200. If you go past row 200, change `B2:B200` and `C2:C200` in **I1** and **I2** to a longer range.

## Template

[Download the .xlsx](templates/tidy-tabs-budget-vs-actual.xlsx)

The file has two tabs. **Budget vs actual** holds the eight sample categories with the formulas from the steps above. The variance, percent, and status columns are shaded as formula columns, rows where actual is above budget turn red, and the totals sit in **H1:I4** beside the table. **How to use** has short fill-in notes.

The categories and amounts are fictional. The file compares what you planned to what you spent; it does not decide where the money should go. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
