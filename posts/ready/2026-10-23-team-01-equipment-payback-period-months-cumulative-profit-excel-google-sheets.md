---
title: How to Calculate Payback Period for Equipment in Excel and Google Sheets
labels: excel-formulas, small-business, equipment
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 26 edited variants (cost 0, negative, text or blank, cost equal to a month's running total, cost above the 24 month total, a blank profit month, a loss before the payback month, a loss after the cost was reached, text in a profit cell, decimal profits). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Calculate the payback period for equipment in Excel and Google Sheets. Free template with a running total, payback month and sample data.
image_prompts: A stainless steel milk pitcher, a small brass bell and a folded canvas apron on a wooden market cart counter, soft morning light, no screens, no readable text, no logos
image_alt: Steel milk pitcher, brass bell and folded apron on a wooden cart counter
threads: How many months until the new machine pays for itself?\nKeep a running total of its monthly profit with =SUM(B$6:B6), then count the months still below the cost.\nA $4,800.00 sample machine pays back in month 12. A dip after that can fool the count, so the sheet has a check cell.
---
Find the payback month with `=COUNTIF(C6:C29,"<"&B2)+1`, where C6:C29 is the running total of the equipment's monthly net profit and B2 is its cost. The result is the first month the running total reaches the cost, unless the total dips back below the cost afterward, which the template flags in G5.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Payback sheet with a $4,800.00 machine, a running total and month 12 highlighted as the payback month](images/2026-10-23-team-01-equipment-payback-period-months-cumulative-profit-excel-google-sheets/equipment-payback-period-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: add each month's profit to a running total, then count how many months the total is still below the cost. The next month is the payback month.

1. In **B2**, type the equipment cost, such as 4800.
2. In **A6:A29**, type the month numbers 1 to 24.
3. In **B6:B29**, type the net profit the equipment added each month. Net means sales from it minus its supplies and other running costs.
4. In **C6**, type `=SUM(B$6:B6)` and fill it down to **C29**. The first reference is locked at row 6, so each row adds one more month.
5. In **B3**, type `=COUNTIF(C6:C29,"<"&B2)+1`.

How the payback formula reads:

- `COUNTIF(C6:C29,"<"&B2)` counts the months where the running total is below the cost.
- `+1` moves to the first month that is not below it.

Both apps document the function on their help pages: [COUNTIF in Excel](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [COUNTIF in Google Sheets](https://support.google.com/docs/answer/3093480). Those pages describe the function, not our test.

### Payback with a fraction of a month

The template's **B4** gives a fractional answer. It takes the whole months before the payback month and adds the share of the payback month needed to cover what is left. For the sample, that is 11 + ($4,800.00 - $4,660.00) / $560.00 = 11.25 months.

**G1** rounds that up to whole months. Both assume profit arrives evenly through the month.

### What the template adds

The built B3 is longer than the bare formula. It shows "Type the cost in B2" when the cost is blank or text, "Cost must be above 0" for zero or a negative cost, and "Not paid back in this range" when no month in the 24 reaches the cost. The counting core is the same COUNTIF.

Column **D** says "Yes" in every month where the running total has reached the cost. **G4** returns the first of those months with MATCH, and **G5** compares it with B3.

### In Excel

The formulas work as typed. COUNTIF, SUM, SUMIF and MATCH are old functions, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed.

## Example

Sample data is fictional: a $4,800.00 espresso machine for a market cart, with made-up monthly profits that grow from $300.00 to $650.00.

| Month | Net profit | Running total | Reached the cost? |
|---|---|---|---|
| 1 | $300.00 | $300.00 | |
| 2 | $325.00 | $625.00 | |
| 9 | $500.00 | $3,600.00 | |
| 10 | $520.00 | $4,120.00 | |
| 11 | $560.00 | $4,660.00 | |
| 12 | $600.00 | $5,220.00 | Yes |
| 24 | $650.00 | $12,800.00 | Yes |

Months 1 to 11 are below $4,800.00, so COUNTIF returns 11 and the formula adds 1 for month 12. If the cost were exactly $4,660.00, the answer would be month 11, because the running total reaches the cost in that month.

The summary in **F1:G5** shows 24 months typed, 0 loss months, and OK in the check cell.

This is arithmetic on the profits you type. It does not discount for interest, and it does not tell you whether to buy anything.

## Troubleshooting

### The payback month looks too early or too late
A loss can move it, and it depends on when. A loss before the total first reaches the cost still gave the right month in our LibreOffice runs. A loss after the cost was reached can make the COUNTIF answer wrong. For example, with a $1,000.00 cost and profits of $600.00, $500.00, -$300.00 and $400.00, B3 says 3 but the total first reached the cost in month 2. G5 shows CHECK, and G4 holds the first month at or above the cost.

### G5 says CHECK
Either a profit cell holds text, such as "n/a", or the two methods disagree after a dip. SUM skips text, so that month would silently drop out of the total. Retype it as a plain number.

### It says Not paid back in this range
No month in the 24 reaches the cost. Check the cost in B2 for a typo. The template stops at row 29.

### The cost cell shows a message instead of a number
Cost 0, a negative cost, text and a blank cell each show a message in B3 and CHECK in G5. Type the cost as a positive number.

## Template

[Download the .xlsx](templates/tidy-tabs-equipment-payback.xlsx)

The file has two tabs. **Payback** holds the yellow inputs: the equipment name in **B1**, the cost in **B2** and the monthly net profit in **B6:B29**. Gray cells are formulas: the payback month in **B3**, the fraction in **B4**, the running total in **C6:C29**, the "Reached the cost?" column in **D**, and a summary with the check in **F1:G5**. The payback row turns green and G5 turns red when it is not OK. **How to use** has short notes.

The machine and the profits are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, and the green and red formatting was not seen in a spreadsheet window. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
