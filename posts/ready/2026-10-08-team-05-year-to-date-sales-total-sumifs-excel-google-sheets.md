---
title: How to Calculate Year-to-Date Sales in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same SUMIFS, DATE and EOMONTH functions; the help links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Year-to-date sales with =SUMIFS(C:C,A:A,">="&DATE(YEAR(F1),1,1),A:A,"<="&F1). Month-to-date with EOMONTH, plus the text-date trap. Free .xlsx template.
image_prompts: Small craft business order desk with a stack of paper receipts held by a binder clip, a wall calendar with a few days circled in red marker, and a brass desk stamp, soft window light, all paper text blurred and unreadable, no screens, no logos
image_alt: Stack of paper receipts beside a wall calendar with circled days and a brass stamp
threads: Year-to-date sales without re-sorting your sheet every month.\nType a Report date in F1, then =SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),1,1),A2:A500,"<="&F1).\nChange F1 and the total moves. Dates saved as text get skipped with no error.
---
Put a report date in **F1** and use `=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),1,1),A2:A500,"<="&F1)`. It adds every amount dated from January 1 of that year through the report date.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Sales log with a report date of 08/31/2026, year-to-date sales of $9,596.75 and month-to-date sales of $1,655.50](images/2026-10-08-team-05-year-to-date-sales-total-sumifs-excel-google-sheets/year-to-date-sales-total-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Keep one row per sale, with a real date in column A and the amount in column C.

1. In row 1, type these headers in A to C: Sale date, Customer, Amount.
2. Type each sale below the headers. Use dates such as 08/04/2026, not words such as "Aug 4".
3. In **E1**, type Report date. In **F1**, type the date you want totals through, for example 08/31/2026.
4. In **E2**, type Year-to-date sales. In **F2**, type `=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),1,1),A2:A500,"<="&F1)`.
5. In **E3**, type Month-to-date sales. In **F3**, type `=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),MONTH(F1),1),A2:A500,"<="&F1)`.
6. Format **F1** as a date and **F2:F3** as currency with two decimals.
7. Change **F1** to another date. Both totals move with it.

The formula has two conditions. `">="&DATE(YEAR(F1),1,1)` builds January 1 of whatever year the report date is in. `"<="&F1` stops at the report date, so sales after it are left out. The `&` joins the operator to the date, which is how SUMIFS reads a comparison.

The sample uses a typed report date so the numbers stay put. You can type `=TODAY()` in **F1** instead for a total that updates every day. Sales dated after today are then skipped automatically.

To total a whole month, including days after the report date, swap the end condition for `"<="&EOMONTH(F1,0)`. EOMONTH returns the last day of the month. The template shows that in **F4**.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) and [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628) pages.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [SUMIFS function](https://support.google.com/docs/answer/3238496) page.

## Example

Sample data below is fictional. It is a small shop's sales log for 2026, with a report date of 08/31/2026.

| Sale date | Customer | Amount |
|---|---|---|
| 01/12/2026 | Harbor Coffee | $1,250.00 |
| 01/28/2026 | Maple Street Bakery | $480.00 |
| 02/10/2026 | Oak & Ember Candles | $920.50 |
| 02/24/2026 | Harbor Coffee | $1,100.00 |
| 04/07/2026 | Maple Street Bakery | $640.00 |
| 05/19/2026 | Harbor Coffee | $1,380.75 |
| 06/23/2026 | Oak & Ember Candles | $705.00 |
| 07/08/2026 | Lakeview Florist | $450.00 |
| 07/30/2026 | Harbor Coffee | $1,015.00 |
| 08/04/2026 | Maple Street Bakery | $520.50 |
| 08/18/2026 | Oak & Ember Candles | $860.00 |
| 08/31/2026 | Lakeview Florist | $275.00 |
| 09/12/2026 | Harbor Coffee | $1,190.00 |
| 10/03/2026 | Lakeview Florist | $330.00 |

With the report date at 08/31/2026, year-to-date sales are $9,596.75 and month-to-date sales are $1,655.50. The 08/31/2026 sale counts because the formula uses `<=`. The 09/12/2026 and 10/03/2026 rows do not.

Change the report date and the totals follow. These are the results from the recalculated file:

| Report date | Year-to-date | Month-to-date |
|---|---|---|
| 07/31/2026 | $7,941.25 | $1,465.00 |
| 08/15/2026 | $8,461.75 | $520.50 |
| 08/31/2026 | $9,596.75 | $1,655.50 |
| 12/31/2026 | $11,116.75 | $0.00 |

Check 07/31/2026 by hand: the first nine rows add up to $7,941.25, and the two July rows add up to $1,465.00. A report date of 12/31/2026 has no December sales, so month-to-date is $0.00.

## Troubleshooting

### One sale is missing from the total and there is no error
The date in that row is probably stored as text. In LibreOffice Calc, SUMIFS skipped a text "08/18/2026" without any warning: year-to-date dropped from $9,596.75 to $8,736.75, exactly the $860.00 on that row. Excel and Google Sheets were not tested for this. The **Text date demo** tab in the template reproduces it.

### How do I spot a text date?
Real dates usually align to the right and text aligns to the left. The template also has two check cells. `=COUNT(A2:A500)` counts real dates and `=COUNT(C2:C500)` counts amounts. In the demo tab they read 13 and 14, and that gap is the tell.

### Can I fix text dates without retyping them?
Try `=DATEVALUE(A2)` in a spare column and copy the results back as values. This was not tested here, and it depends on your app's date settings, so check one row against the original.

### The total is $0.00 or far too low after changing the report date
Check that **F1** holds a real date and not text. A text report date makes the `"<="&F1` condition match nothing useful. Retype it as 08/31/2026 and format the cell as a date.

### Sales after row 500 are ignored
The formulas read rows 2 to 500. If your log is longer, change 500 to a bigger number in every SUMIFS formula.

## Template

[Download the .xlsx](templates/tidy-tabs-year-to-date-sales-total.xlsx)

The file has three tabs. **Sales** holds 14 sample sales, the report date in **F1**, year-to-date, month-to-date and whole-month totals in **F2:F4**, helper dates in **F5:F7**, and two count checks. **Text date demo** is the same data with one date stored as text, so you can see the skipped row. **How to use** has short fill-in notes.

The customers and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
