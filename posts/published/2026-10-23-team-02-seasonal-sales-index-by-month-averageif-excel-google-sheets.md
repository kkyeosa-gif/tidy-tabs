---
threads_url: https://www.threads.com/@tin_ylab/post/De2Xxkcmv2V
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-your-busiest-and-slowest.html
title: How to Find Your Busiest and Slowest Months With a Seasonal Index in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model, including 18 edited variants (a month with one year of data, a month with no data, no data at all, unequal years, a month number of 13, a month or amount typed as text, a blank or zero goal, all sales at 0, a negative month, a tie for busiest). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Find your busiest and slowest months with a seasonal index in Excel and Google Sheets. Free template with sample monthly sales.
image_prompts: A roll of red ribbon, three wrapped kraft paper gift boxes and a pair of scissors on a wooden gift shop counter, warm late afternoon light, no screens, no readable text, no logos
image_alt: Wrapped kraft gift boxes, red ribbon roll and scissors on a wooden shop counter
threads: Which month is your shop's real peak?\nAVERAGEIF one calendar month across every year, then divide by your average month.\nIn the sample, December comes out at 2.24 and February at 0.68. The shop is made up, so type your own sales.
---
Find each calendar month's index in two steps. Average that month across all your years with `=AVERAGEIF(Sales!$B$2:$B$200,A2,Sales!$C$2:$C$200)`, where A2 is the month number, then divide by the average of every month with `=C2/AVERAGE(Sales!$C$2:$C$200)`. An index of 2.24 means that month sells about 2.2 times the typical month; 0.68 means about 68% of it.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Index sheet listing twelve months with average sales and a seasonal index, December highest](images/2026-10-23-team-02-seasonal-sales-index-by-month-averageif-excel-google-sheets/seasonal-sales-index-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: put every month of sales in one list, tag each row with its month number, and let AVERAGEIF pull one calendar month out of it at a time.

1. Open a tab named **Sales**. Use three columns: **A** Year, **B** Month number (1 to 12), **C** Sales. Type one row per month, starting in row 2. The formulas read down to row 200.
2. Type the month as a number, not as text. A "1" typed as text is not counted.
3. Open a second tab named **Index**. In **A2:A13**, type the numbers 1 to 12.
4. In **C2**, type `=AVERAGEIF(Sales!$B$2:$B$200,A2,Sales!$C$2:$C$200)` and fill down to **C13**. This is the average sales for that calendar month across all years.
5. In **D2**, type `=C2/AVERAGE(Sales!$C$2:$C$200)` and fill down to **D13**. This divides by the average of every month on the Sales tab.
6. Format column D with two decimals.

`AVERAGEIF` averages only the rows where column B equals the month number in A2. `AVERAGE` covers every row and skips blanks. The dollar signs keep the ranges fixed as you fill down.

You can join the two steps into one formula, `=AVERAGEIF(...)/AVERAGE(...)`. The template splits them across C2 and D2 so you can see the monthly average before it becomes an index.

### What the template adds

The template's formulas are longer. C2 shows "No data" when a month has no rows, D2 shows "Check sales" instead of a divide by zero error, and column E counts how many years of data each month has. Column F spreads a yearly goal you type in **I1** across the months. The summary in **H1:I11** names the busiest and slowest months.

### In Excel

The formulas work as typed. AVERAGEIF and AVERAGE are old functions, and no newer function is needed. Microsoft's help pages describe them: [AVERAGEIF function](https://support.microsoft.com/en-us/office/averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642) and [AVERAGE function](https://support.microsoft.com/en-us/office/average-function-047bac88-d466-426c-a32b-8f33eb960cf6). They are help pages, not our test.

### In Google Sheets

The same formulas should work as typed. Google Sheets was not opened for this post.

## Example

Sample data is fictional: a gift shop with 24 months of sales, 2024 and 2025. The average month is $8,195.83 (total $196,700.00 over 24 months).

| Month | Average sales (C) | Index (D) |
|---|---|---|
| Jan | $6,400.00 | 0.78 |
| Feb | $5,550.00 | 0.68 |
| Mar | $5,950.00 | 0.73 |
| Apr | $6,300.00 | 0.77 |
| May | $7,200.00 | 0.88 |
| Jun | $6,700.00 | 0.82 |
| Jul | $6,050.00 | 0.74 |
| Aug | $6,450.00 | 0.79 |
| Sep | $6,950.00 | 0.85 |
| Oct | $8,500.00 | 1.04 |
| Nov | $13,950.00 | 1.70 |
| Dec | $18,350.00 | 2.24 |

December is $18,350.00 divided by $8,195.83, which is 2.24. February is $5,550.00 divided by $8,195.83, which is 0.68. The template shades months above 1.00 green.

The 12 indexes add up to 12.00 here because every month has two years of data. This is arithmetic on past sales, not a forecast or stocking advice.

### Spreading a yearly goal

Type a goal such as $90,000.00 in **I1**. Each month gets `ROUND($I$1*C2/SUM($C$2:$C$13),2)` in column F. Rounding each month to cents can leave the total off the goal. In the sample the 12 amounts add to $89,999.98, and **I10** shows the $0.02 gap. The gap can go either way, so check I10.

## Troubleshooting

### A month shows "No data"
No row on the Sales tab has that month number. Check that the month is typed as a number from 1 to 12 in column B. The goal is then shared among the other months.

### The indexes do not add up to 12.00
The months have different numbers of years. Say you add nine months of 2026 to the sample: the indexes then add to 12.73. Column E shows the years of data per month, so you can see which months have more.

### The check cell says CHECK
**I11** turns red when a month number is outside 1 to 12, a month is typed as text, a month cell is blank next to an amount, or an amount is typed as text such as "n/a". AVERAGE skips text, so a bad row can be ignored without any error.

## Template

[Download the .xlsx](templates/tidy-tabs-seasonal-sales-index.xlsx)

The file has three tabs. **Sales** holds the yellow inputs in **A2:C200** (year, month number, sales) with 24 sample months in rows 2 to 25. **Index** holds the month numbers in **A2:A13**, the formulas in **B2:F13**, a yearly goal in **I1** and the summary and check in **H1:I11**. **How to use** has short notes. Sales and How to use print portrait and Index prints landscape, all on US Letter (8.5 x 11 in).

The sales figures are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
