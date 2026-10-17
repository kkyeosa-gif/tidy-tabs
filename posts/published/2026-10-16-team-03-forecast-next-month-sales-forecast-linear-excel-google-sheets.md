---
threads_url: https://www.threads.com/@tin_ylab/post/DekxrpIoITv
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-forecast-next-months-sales-with.html
title: How to Forecast Next Month's Sales With FORECAST in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the forecasts and slope were checked against a Python least squares fit, including flat sales, falling sales, two months only, a month number typed as text, a blank cell and a mismatched range. Excel and Google Sheets were not opened; the help links are from Microsoft and Google.
search_description: Forecast next month's sales with =FORECAST(E2,$C$2:$C$25,$A$2:$A$25), a straight-line estimate from past months. Free .xlsx template.
image_prompts: Row of wooden market crates holding stacked jars of honey beside a chalkboard easel and a small cash tin on a farm stand table, soft daylight, no screens, no readable text, no logos
image_alt: Honey jars stacked in crates beside an easel and a small cash tin
threads: Eight months of sales, $1,200.00 up to $1,850.00, forecast $1,926.07 for Sep 2026.\n=FORECAST(E2,$C$2:$C$25,$A$2:$A$25) reads the next month off a straight line.\nThe line adds about $83.57 a month, and it needs at least 2 months of sales.
---
Estimate the next month with `=FORECAST(E2,$C$2:$C$25,$A$2:$A$25)`, where E2 is the next month number, C2:C25 holds past sales and A2:A25 holds the month numbers. It fits a straight line through your past months, so it is an estimate, not a prediction.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same function as typed. Neither Excel nor Google Sheets was opened for this post.

![Eight months of sales next to a three month forecast and trend per month](images/2026-10-16-team-03-forecast-next-month-sales-forecast-linear-excel-google-sheets/forecast-next-month-sales-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

FORECAST draws the straight line that best fits your past sales, then reads off where that line sits for a month you choose.

1. On a tab named **Sales**, type these headers in row 1: Month number, Month, Sales.
2. In column **A**, type 1, 2, 3 and so on, one number per month, with no gaps and in order.
3. In column **B**, type the first day of each month as MM/DD/YYYY, such as 01/01/2026. Format the cells as **mmm yyyy** if you want a short label.
4. In column **C**, type that month's sales as plain numbers.
5. In **E2**, type `=IF(COUNT($A$2:$A$25)=0,"",MAX($A$2:$A$25)+1)`. This is the next month number. Make **E3** add 2 and **E4** add 3 if you want three months ahead.
6. In **F2**, type `=IF(OR(E2="",COUNT($B$2:$B$25)=0),"",EDATE(MAX($B$2:$B$25),E2-MAX($A$2:$A$25)))`. This shows which calendar month the forecast is for.
7. In **G2**, type `=IF(OR(E2="",COUNT($C$2:$C$25)<2),"",FORECAST(E2,$C$2:$C$25,$A$2:$A$25))`. Copy it down to **G4**, which then reads E3 and E4.
8. In **G6**, type `=IF(COUNT($C$2:$C$25)<2,"",SLOPE($C$2:$C$25,$A$2:$A$25))`. This is how many dollars the line rises (or falls) each month.

FORECAST takes the new month number first, then the known sales, then the known month numbers. The ranges stop at row 25, which holds 24 months. You need at least 2 months of sales, and more is better, because two points always make a perfect line.

### In Excel

The formulas work as typed. Microsoft documents the function on its [FORECAST and FORECAST.LINEAR functions](https://support.microsoft.com/en-us/office/forecast-and-forecast-linear-functions-50ca49c9-7b40-4892-94e4-7ad38bbeda99) page. This post uses the older name FORECAST, which Microsoft keeps for compatibility. FORECAST.LINEAR is the newer name and was not run for this post. Format the results from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [FORECAST help page](https://support.google.com/docs/answer/3094000). Format the results with **Format > Number > Currency**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is one shop's sales for January through August 2026.

| Month number | Month | Sales |
|---|---|---|
| 1 | Jan 2026 | $1,200.00 |
| 2 | Feb 2026 | $1,350.00 |
| 3 | Mar 2026 | $1,500.00 |
| 4 | Apr 2026 | $1,480.00 |
| 5 | May 2026 | $1,620.00 |
| 6 | Jun 2026 | $1,710.00 |
| 7 | Jul 2026 | $1,690.00 |
| 8 | Aug 2026 | $1,850.00 |

| Result | Cell | Value |
|---|---|---|
| Next month number | E2 | 9 |
| Forecast for Sep 2026 | G2 | $1,926.07 |
| Forecast for Oct 2026 | G3 | $2,009.64 |
| Forecast for Nov 2026 | G4 | $2,093.21 |
| Slope per month | G6 | $83.57 |
| Sales months counted | G7 | 8 |
| Hand calculation check | G9 | OK |

The line adds about $83.57 each month, so each forecast is $83.57 above the one before it.

## Troubleshooting

### The forecast cells are blank
The formulas stay empty until there are at least 2 sales figures and a month number. Check that column C holds numbers, not text, and that column A has month numbers. With every input cleared, G7 shows 0.

### The check says CHECK
G8 repeats the first forecast by hand with SUMPRODUCT, SUMSQ and AVERAGE, and G9 compares the two. In the recalculated copy, a month number typed as text, a blank month number or a blank sales cell in the middle made G9 show CHECK. LibreOffice skipped that pair when fitting the line (month 9 came out at $1,928.39), while G7 still counted every sales cell. Fill the gap so every month is complete.

### FORECAST shows an error
The month numbers and the sales must cover the same rows. In a copy of the template, `=FORECAST(9,C2:C9,A2:A8)` has 8 sales cells but only 7 month numbers, and LibreOffice returned an error. Make both ranges start and end on the same rows.

### The forecast looks too high or too low
A straight line cannot see a holiday rush, a price change or a slow summer. It only continues the overall climb or fall. In the recalculated copy, falling sales from $2,000.00 down to $1,300.00 gave a slope of -$98.81 per month and forecasts of $1,217.86, $1,119.05 and $1,020.24. Flat sales of $1,000.00 gave a slope of $0.00 and forecasts of $1,000.00.

## Template

[Download the .xlsx](templates/tidy-tabs-forecast-next-month.xlsx)

The file has two tabs. **Sales** holds the 8 sample months in yellow input cells with room down to row 25, the three forecast rows in **E2** to **G4**, and the slope, month count and check in **G6** to **G9**. **How to use** has short fill-in notes.

The sales are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
