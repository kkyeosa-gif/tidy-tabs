---
threads_url: https://www.threads.com/@tin_ylab/post/DeaDB0WH8xQ
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-tiered-sales.html
title: How to Calculate Tiered Sales Commission in Excel and Google Sheets
labels: excel-formulas, sales, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math, including a boundary case, a one-cent case, a negative-sales case and a text-number case. Excel and Google Sheets were not opened; the help link is from Microsoft Support.
search_description: Pay each commission tier only on sales inside that tier with =SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4). Free .xlsx template.
image_prompts: Wooden desk with a small brass bell, a stack of paper sales receipts held by a binder clip, and a pencil resting on a plain notepad, soft window light, no screens, no readable text, no logos
image_alt: Brass bell beside clipped paper receipts and a pencil on a wooden desk
threads: Paying 3%, 5% and 8% tiers and your sheet pays 8% on everything?\nUse =SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4) so each rate only hits the sales inside its tier.\n$12,000.00 in sales comes out to $560.00.
---
Pay each tier's rate only on the sales inside that tier with `=SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4)`, where column G holds where each tier starts and column I holds how much each tier raises the rate. With 3%, 5% and 8% tiers, $12,000.00 in sales earns $560.00.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same function and were not opened in either app. This is arithmetic for a written commission plan, not payroll or tax advice.

![Tiered sales commission template with six reps, tier starts and rates, and a check column](images/2026-10-12-team-02-tiered-sales-commission-sumproduct-excel-google-sheets/tiered-sales-commission-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

A tiered plan pays a higher rate only on the part of the sales that falls above each threshold. The trick is to turn each rate into a raise: tier 1 pays 3% on everything, tier 2 adds 2% on sales above $5,000.00, and tier 3 adds 3% more on sales above $10,000.00.

This works when the tiers are in ascending order and sales are real numbers.

1. In **G1:I1**, type the headers Tier starts, Rate and Extra rate.
2. In **G2:G4**, type where each tier starts: `0`, `5000` and `10000`.
3. In **H2:H4**, type the rate for each tier: `3%`, `5%` and `8%`.
4. In **I2**, type `=H2`. In **I3**, type `=H3-H2`. In **I4**, type `=H4-H3`. These give 3.0%, 2.0% and 3.0%.
5. Type each rep's name in **A2:A7** and their sales in **B2:B7**.
6. In **C2**, type `=ROUND(SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4),2)` and fill it down to row 7.
7. In **B8** and **C8**, add `=SUM(B2:B7)` and `=SUM(C2:C7)` for the totals.

Here is what the formula does for $12,000.00. The first test, `B2>$G$2:$G$4`, is true for every tier that sales have passed. Then `B2-$G$2:$G$4` is how far sales went past each start. Multiplying by the extra rate gives $12,000.00 x 3% = $360.00, plus $7,000.00 x 2% = $140.00, plus $2,000.00 x 3% = $60.00. The total is $560.00.

The `$` signs lock the tier ranges. The reference to `B2` has no `$`, so it moves down a row as you fill.

### Check it a second way

The template also adds a second method in column D, so a typo in a tier shows up right away:

1. In **D2**, type `=ROUND(MIN(B2,$G$3)*$H$2+MAX(0,MIN(B2,$G$4)-$G$3)*$H$3+MAX(0,B2-$G$4)*$H$4,2)` and fill it down to row 7. This prices each slice of sales at its own rate.
2. In **E2**, type `=IF(ROUND(C2-D2,2)=0,"Yes","CHECK")` and fill it down.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) page. This was not opened in Excel for this post.

### In Google Sheets

The same formulas should work as typed. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. Tiers start at $0.00, $5,000.00 and $10,000.00 and pay 3.0%, 5.0% and 8.0%.

| Rep | Sales | Commission |
|---|---|---|
| Ana | $4,200.00 | $126.00 |
| Ben | $5,000.00 | $150.00 |
| Cy | $12,000.00 | $560.00 |
| Dee | $0.00 | $0.00 |
| Eli | $25,500.50 | $1,640.04 |
| Fay | $10,000.00 | $400.00 |
| Total | $56,700.50 | $2,876.04 |

Ana stays in tier 1, so $4,200.00 x 3% is $126.00. Eli's $25,500.50 is $150.00 from the first $5,000.00, $250.00 from the next $5,000.00 and $1,240.04 from the last $15,500.50.

In the recalculated file, column D matched column C on every row and column E read Yes for all six reps. Every value matched a Python check.

## Troubleshooting

### A rep with exactly $5,000.00 gets the lower rate on the last dollar
That is the plan as built. The test uses `>`, so sales that land exactly on a tier start earn nothing extra from that tier. Ben's $5,000.00 pays $150.00, and Fay's $10,000.00 pays $400.00. If your written plan says the boundary dollar belongs to the higher tier, check the wording before changing the formula.

### A sale of $10,000.01 barely changes the result
The extra cent earns $0.0008 at the top tier, so $10,000.01 still shows $400.00 in both methods. That is the rounding to cents, not a mistake.

### Negative sales make column E say CHECK
With -100 in **B5**, the SUMPRODUCT column gives $0.00 but the second method gives -$3.00, so the two disagree and E5 flags CHECK. Handle refunds before you use this sheet, for example by netting them against the month's sales.

### Sales typed as text give $0.00
A value such as "5000" stored as text made the SUMPRODUCT column return $0.00 while the second method returned $400.00 in the recalculated file, so E flagged it. Keep sales as real numbers: retype the cell or convert the column to a number format. How Excel and Google Sheets treat text numbers was not tested.

### You need a fourth tier
The template has three tiers. To add one, insert a row inside the tier table, copy the formulas in column I, and extend every `$G$2:$G$4` and `$I$2:$I$4` range to include it. Keep the tier starts in ascending order, or the results will be wrong.

### The sheet does not say what someone is owed
This is arithmetic for a written plan. It does not cover payroll, withholding or taxes, and it does not decide what a rep is owed. Check your commission agreement for caps, draws and when a sale counts.

## Template

[Download the .xlsx](templates/tidy-tabs-tiered-commission.xlsx)

The file has two tabs. **Commission** (printed landscape on US Letter) has the six sample reps and their sales in A2:B7, the SUMPRODUCT commission in C2:C7, the second method in D2:D7, the Yes or CHECK match in E2:E7, totals in row 8, and the yellow inputs in G2:H4 with the extra rates in I2:I4. **How to use** has short fill-in notes.

The names and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
