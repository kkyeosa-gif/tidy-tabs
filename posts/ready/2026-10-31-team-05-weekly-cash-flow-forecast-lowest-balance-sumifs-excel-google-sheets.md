---
title: How to Build a Weekly Cash Flow Forecast in Excel and Google Sheets
labels: excel-formulas, small-business, cash-flow
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model of the weekly totals, ending balances, lowest balance and check cell, in 28 workbooks (5,628 cell comparisons, 0 mismatches, no error values). Checked: items on the first and last day of the 13 weeks and one day after, an item before the first Monday, a start date that is not a Monday, In and Out typed in capitals, Income as a Type, a negative amount, text dates and amounts, a trailing space, no items, a negative opening balance, a tie for the lowest balance, times on dates and 100 items. The PDF export pages are US Letter. Excel and Google Sheets were not used.
search_description: Build a 13 week cash flow forecast with SUMIFS in Excel and Google Sheets. Free template totals money in and out per week and shows your lowest balance.
image_prompts: A paper desk calendar with several weeks circled in pencil next to a small stack of paper receipts held by a binder clip and a handheld calculator, soft daylight, no screens, no readable text, no logos
image_alt: Desk calendar with weeks circled in pencil beside clipped receipts and a handheld calculator
threads: Will the account run short before that client invoice lands?\nList each expected payment with a date, then SUMIFS totals money in and out for every week. MIN finds the lowest balance and INDEX/MATCH names the week.\nThe template covers 13 weeks and flags any week under a line you pick.
---
List each expected payment with a date, then total each week with `=SUMIFS(Items!$D$5:$D$104,Items!$C$5:$C$104,"In",Items!$A$5:$A$104,">="&A6,Items!$A$5:$A$104,"<"&A6+7)`. The ending balance is the opening balance plus the week's money in minus money out, and `=MIN(F6:F18)` shows the lowest point.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. SUMIFS, MIN, INDEX and MATCH are in every current version of Excel and Google Sheets, but neither app was opened for this post.

![Thirteen weeks with money in, money out, ending balance, a Below line flag and the lowest balance of $310.00](images/2026-10-31-team-05-weekly-cash-flow-forecast-lowest-balance-sumifs-excel-google-sheets/weekly-cash-flow-forecast-lowest-balance-template.png) *caption: the Forecast tab of the template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

The layout matches the template. **Items** has the date in column A, a description in B, the type (In or Out) in C and the amount in D, starting in row 5. **Forecast** has the opening balance in **B1**, the first Monday in **B2** and your low-balance line in **B3**. The 13 weeks run in rows 6 to 18.

The steps are the same in Excel and Google Sheets.

1. On **Forecast**, type your opening balance in **B1**, such as $2,400.00, and the first Monday of the forecast in **B2**, such as 11/02/2026.
2. In **B3**, type a low-balance line, such as $500.00. This is a number you choose, not advice.
3. On **Items**, list every payment you expect, one per row. Type the amount as a positive number and type In or Out in column C. Dates must be real dates, not text.
4. In **A6**, type `=IF(ISNUMBER($B$2),$B$2+7*(ROWS($A$6:A6)-1),"")`. In **B6**, type `=A6+6`.
5. In **C6**, type the money-in formula below. In **D6**, type the same formula with "Out" in place of "In".
6. In **E6**, type `=ROUND(C6-D6,2)`.
7. In **F6**, type the balance formula below.
8. Select **A6:F6** and drag the fill handle down to row 18.
9. Next to the table, type `=MIN(F6:F18)` for the lowest balance.

The money-in formula in **C6**:

`=SUMIFS(Items!$D$5:$D$104,Items!$C$5:$C$104,"In",Items!$A$5:$A$104,">="&A6,Items!$A$5:$A$104,"<"&A6+7)`

It adds the amounts typed In dated on or after the week start and before the next one, so every day belongs to exactly one week.

The balance formula in **F6**:

`=ROUND(IF(ROWS($A$6:A6)=1,$B$1,F5)+E6,2)`

The first row reads the opening balance in **B1** directly, so the header text above it is never used. Every later row starts from the row above.

To name the week of the lowest balance, use `=INDEX(A6:A18,MATCH(MIN(F6:F18),F6:F18,0))`. If two weeks tie, it returns the first one.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), [MIN function](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The Google page was not link-checked.

## Example

Sample data is fictional: a made up freelancer and shop with 30 expected payments from 11/02/2026 to 01/29/2027, an opening balance of $2,400.00 and a line of $500.00.

| Week starts | Money in | Money out | Ending balance | Flag |
|---|---|---|---|---|
| 11/02/2026 | $600.00 | $970.00 | $2,030.00 | |
| 11/16/2026 | $1,800.00 | $390.00 | $3,735.00 | |
| 11/30/2026 | $0.00 | $3,575.00 | $310.00 | Below line |
| 12/07/2026 | $190.00 | $45.00 | $455.00 | Below line |
| 12/14/2026 | $1,940.00 | $210.00 | $2,185.00 | |
| 01/11/2027 | $0.00 | $255.00 | $1,150.00 | |
| 01/25/2027 | $1,400.00 | $0.00 | $4,110.00 | |

The week of 11/30/2026 holds a $1,275.00 equipment repair, the December rent and a $1,650.00 supplies order, so the balance falls to $310.00 before the next payouts arrive. That is the lowest point, and the summary beside the table names the week and counts 2 weeks below the line.

Over 13 weeks, $9,020.00 comes in and $7,310.00 goes out, so week 13 ends at $4,110.00. The check cell says OK. The forecast is only as good as the dates and amounts you type, and it gives no financial, tax or lending advice.

## Troubleshooting

### A week shows $0.00 for money in, but an item exists
Check the **Row check** column on **Items**. In the tests, a Type of "Income" or "In " with a trailing space matched nothing and showed Check type. Capital letters do not matter: "in" and "OUT" were counted. Retype the Type as In or Out.

### An item is in no week at all
Its date is before the first Monday or after the 13th week. In the tests, 02/01/2027 fell one day outside a window that starts on 11/02/2026, and the item showed Outside weeks. Move the first Monday or fix the date.

### The week table is blank and the summary says n/a
The first Monday in **B2** is text, not a date. Retype it as a real date. The opening balance in **B1** has the same effect on the ending balances: typed as text, they stay blank.

### Weeks do not start on a Monday
The forecast runs 7 days from whatever date is in **B2**. Type a Monday for Monday weeks. **C2** shows the weekday.

### The check cell says Check
It says Check for a text date, a Type other than In or Out, a text or negative amount, an item outside the 13 weeks, a setting that is not a number, or data below row 104. The summary cells show which.

## Template

[Download the .xlsx](templates/tidy-tabs-weekly-cash-flow.xlsx)

The file has three tabs. **Items** has 30 fictional payments in yellow cells **A5:D104** (100 slots), with an In or Out dropdown and a Row check column. **Forecast** has the yellow opening balance, first Monday and low-balance line, the 13 weeks of totals and balances, and a summary of the lowest balance, its week, weeks below the line and a check. **How to use** has notes on each formula.

Limits: it holds 100 items and 13 weeks, and more was not tested. **Forecast** prints landscape and **Items** prints portrait, both on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
