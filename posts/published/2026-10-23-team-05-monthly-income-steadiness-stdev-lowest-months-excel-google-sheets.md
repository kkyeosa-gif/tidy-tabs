---
threads_url: https://www.threads.com/@tin_ylab/post/De4aPqQEecV
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-measure-how-much-your-monthly.html
title: How to Measure How Much Your Monthly Income Varies With STDEV in Excel and Google Sheets
labels: excel-formulas, freelancers, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model (statistics.stdev and a Decimal average), including 16 edited variants: 12 identical months, two months, one month, no months, a blank month, text entries such as n/a, all zeros, zero months among twelve, negative months, large numbers and tied lowest months. No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Measure how much monthly income varies with STDEV in Excel and Google Sheets. Free template with average, spread and lowest months.
image_prompts: Twelve small kraft paper envelopes fanned out in a row on a wooden table beside a wooden ruler and a pencil, soft window light, no screens, no readable text, no logos
image_alt: Row of twelve kraft envelopes beside a wooden ruler and pencil on a table
threads: How much does your monthly income actually swing?\nDivide STDEV by AVERAGE. With 12 sample months averaging $2,650.00 and a STDEV of $577.61, you get 21.8%.\nLower means steadier months. It only describes past numbers, so the template has no advice in it.
---
Measure how much your monthly income swings with `=STDEV(Income!B2:B13)/AVERAGE(Income!B2:B13)`, which divides the standard deviation by the average. A result of 0.22 means the standard deviation is 22% of the average; a lower number means steadier months.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Steadiness sheet showing average month, standard deviation and variation percent for twelve months of income](images/2026-10-23-team-05-monthly-income-steadiness-stdev-lowest-months-excel-google-sheets/income-steadiness-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: find the average month, find how far months typically sit from it, then express that gap as a percent of the average.

1. On a tab named **Income**, type a month in **A2:A13** and that month's income in **B2:B13**. Twelve months fit.
2. On a tab named **Steadiness**, type `=AVERAGE(Income!$B$2:$B$13)` in **B2**.
3. In **B3**, type `=STDEV(Income!$B$2:$B$13)`. This is the standard deviation: roughly how far a typical month lands from the average.
4. In **B4**, type `=B3/B2` and format the cell as a percent with one decimal.
5. Optional: in **B5**, type `=MIN(Income!$B$2:$B$13)` for the lowest month.
6. Optional: in **B6**, type `=(SMALL(Income!$B$2:$B$13,1)+SMALL(Income!$B$2:$B$13,2)+SMALL(Income!$B$2:$B$13,3))/3` for the average of your three lowest months.

`STDEV` is the sample standard deviation, the version meant for a set of months you treat as a sample. The template uses `STDEV` rather than `STDEV.S`, so the file opens the same way in older Excel.

`SMALL(range,1)` is the smallest number, `SMALL(range,2)` the second smallest, and so on. Adding the first three and dividing by 3 gives the average of the lowest three.

### In Excel

Type the formulas as shown. Microsoft's help pages for [STDEV](https://support.microsoft.com/en-us/office/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0) and [SMALL](https://support.microsoft.com/en-us/office/small-function-17da8222-7c82-42b2-961b-14c45384df07) describe them. Those are the help pages, not our test.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [STDEV help page](https://support.google.com/docs/answer/3094054). That is the help page, not our test.

## Example

The numbers below are for a fictional freelance designer, October 2025 to September 2026.

| Month | Income |
|---|---|
| 10/01/2025 | $2,400.00 |
| 11/01/2025 | $3,100.00 |
| 12/01/2025 | $1,800.00 |
| 01/01/2026 | $2,900.00 |
| 02/01/2026 | $3,500.00 |
| 03/01/2026 | $2,200.00 |
| 04/01/2026 | $2,700.00 |
| 05/01/2026 | $3,000.00 |
| 06/01/2026 | $1,600.00 |
| 07/01/2026 | $2,800.00 |
| 08/01/2026 | $3,300.00 |
| 09/01/2026 | $2,500.00 |

The Steadiness sheet then shows:

| Measure | Cell | Result |
|---|---|---|
| Average month | B2 | $2,650.00 |
| Standard deviation | B3 | $577.61 |
| Variation (B3/B2) | B4 | 21.8% |
| Lowest month | B5 | $1,600.00 |
| Average of the 3 lowest months | B6 | $1,866.67 |
| Months below average minus one standard deviation | B7 | 2 |

The three lowest months are $1,600.00, $1,800.00 and $2,200.00. They add up to $5,600.00, and $5,600.00 / 3 is $1,866.67.

B7 uses `=COUNTIF(Income!$B$2:$B$13,"<"&(B2-B3))`. The average minus one standard deviation is $2,650.00 - $577.61 = $2,072.39, and two months fall below that line: $1,800.00 and $1,600.00.

This only describes the spread of numbers you already typed. It does not say how much to save or spend, and it does not set any line between steady and unsteady.

## Troubleshooting

### My average ignores a month
A blank cell is skipped by `AVERAGE` and `STDEV`, so the other months count as if that month did not exist. If a month truly had no income, type 0 so it counts as $0.00. In the template, B8 shows how many months were counted.

### The check cell says CHECK
An entry in Income B2:B13 is text, such as n/a, or a number typed with a dollar sign as text, or fewer than 3 months are filled. Text is skipped by `AVERAGE` and `STDEV`, so the results quietly use fewer months. Retype the entry as a plain number.

### B6 says Need 3 or more months
`SMALL(range,3)` needs at least three numbers. With only two months, the standard deviation still works: $2,000.00 and $3,000.00 give $707.11. With one month, B3 says Need 2 or more months.

### B4 says Average is 0 or below
Dividing by an average of zero is not possible, so the template shows words instead of #DIV/0!. Negative months do work in the arithmetic, but a percent of a zero or negative average has no useful meaning.

### I have more than 12 months
Only rows 2 to 13 are read. Change **B13** to your last row in every formula, including the check cell.

## Template

The download has three sheets: **Income** (yellow cells A2:B13 for your months), **Steadiness** (average, standard deviation, variation percent, lowest month, average of the 3 lowest, months below average minus one standard deviation, month count, highest month and a CHECK cell) and **How to use**. It prints on US Letter, with the Steadiness sheet in landscape. The sample months and income are fictional. Replace the yellow cells with yours.

[Download the .xlsx](templates/tidy-tabs-income-steadiness.xlsx)

To use it in Google Sheets, choose **File > Import > Upload**, then select the downloaded file. Import was not tested for this post.
