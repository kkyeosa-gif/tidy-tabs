---
title: How to Calculate Average Yearly Sales Growth (CAGR) in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 143 edited variants (other year pairs, flat and falling sales, a start or end of 0, negative and text inputs, reversed years, a gap year, 10 years of data). No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Find average yearly sales growth (CAGR) with =RRI(4,B2,B6) in Excel and Google Sheets. Free template with a year by year check of the rate.
image_prompts: A stack of small kraft mailer boxes beside a spool of natural twine and a wooden rubber stamp on a packing table, soft window light, no screens, no readable text, no logos
image_alt: Kraft mailer boxes, twine spool and wooden stamp on a small shop packing table
threads: Counting 5 years of sales as 5 periods?\nIt is 4. =RRI(4,B2,B6) gives 15.69% for $12,000.00 to $21,500.00, and using 5 gives 12.37%.\nThe real years ran from 11.97% to 20.13%, so one average hides the uneven ones.
---
Find the average yearly growth rate with `=RRI(4,B2,B6)`, where 4 is the number of years between the first and last sales figure, B2 is the first year's sales and B6 is the last year's. The same result comes from `=(B6/B2)^(1/4)-1`.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Growth sheet with yearly sales and the average growth rate result](images/2026-10-21-team-01-average-annual-sales-growth-cagr-rri-excel-google-sheets/annual-growth-rate-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

This rate is also called CAGR, the compound annual growth rate. It answers one question: what steady yearly percentage would take sales from the first year to the last?

1. In **A2:A6**, type the years: 2022 to 2026.
2. In **B2:B6**, type the sales for each year, for example $12,000.00 in B2 and $21,500.00 in B6.
3. Count the periods as the last year minus the first year: 2026 - 2022 = 4. Do not count the years listed, which is 5.
4. In any empty cell, type `=RRI(4,B2,B6)`.
5. Format that cell as a percentage with two decimals.

RRI takes the number of periods, the start value and the end value, in that order.

### Let the sheet find the numbers

The template does not type the 4 for you. It reads it from two year cells:

- **E1** holds the start year (2022) and **E2** the end year (2026). Type them as numbers.
- **E3** is the periods: `=E2-E1`.
- **E4** and **E5** look up the sales for those years with INDEX and MATCH.
- **E6** is the rate. It wraps RRI in guards, so it is not the bare `=RRI(E3,E4,E5)`. If the start is 0 or less, the years are the same or reversed, or a cell holds text, it shows "Check inputs".
- **E7** is the same rate by the power form, `=(E5/E4)^(1/E3)-1`.

Change E1 and E2 to measure any two years in the table, such as 2023 to 2025.

### In Excel

The formulas work as typed. Microsoft describes the arguments on its [RRI function](https://support.microsoft.com/en-us/office/rri-function-6f5822d8-7ef1-4233-944c-79e8172930f4) page.

### In Google Sheets

Google Sheets has RRI too, listed on Google's [function list](https://support.google.com/docs/table/25273). The same formulas should work as typed, but Sheets was not opened for this post.

## Example

Sample data is fictional: yearly sales for a small Etsy shop.

| Year | Sales | Change vs year before |
|---|---|---|
| 2022 | $12,000.00 | |
| 2023 | $14,200.00 | 18.33% |
| 2024 | $15,900.00 | 11.97% |
| 2025 | $19,100.00 | 20.13% |
| 2026 | $21,500.00 | 12.57% |

The template results for 2022 to 2026:

| Cell | Result |
|---|---|
| E3 Periods | 4 |
| E6 Average yearly rate | 15.69% |
| E8 Total growth | 79.17% |
| E9 Start sales grown at 15.69% for 4 years | $21,500.00 |
| E10 Check | OK |

E9 rebuilds the end value as `=E4*(1+E6)^E3`. E10 shows OK when that matches the real end sales to the cent.

The Check tab also shows what sales would have been at a steady 15.69%: $13,883.39 for 2023 against a real $14,200.00, and $16,062.38 for 2024 against a real $15,900.00. Real years ran from 11.97% to 20.13%, so the single average hides uneven years. This is arithmetic on past sales, not a forecast and not a judgment of what a good rate is.

## Troubleshooting

### The rate looks too small or too big
You probably counted the years listed instead of the gaps between them. With 5 years of data there are 4 periods. Using 5 gives 12.37% instead of 15.69%.

### The cell says Check inputs
The start sales are 0 or less, the start and end years are the same or reversed, a year is not typed in column A, or a year or sales cell holds text such as "n/a" or "$12,000" typed as text. Fix the cell and the rate returns.

### The rate shows -100.00%
The end year's sales are $0.00. Growing to nothing is a -100.00% rate, and E9 rebuilds $0.00. Check that the sales cell is not simply blank.

### The average rate does not match any single year
That is expected. One average cannot match uneven years. Compare it with the Change column on the Check tab, where the best year was 20.13% and the weakest 11.97%.

## Template

[Download the .xlsx](templates/tidy-tabs-annual-growth-rate.xlsx)

The file has three tabs. **Growth** holds the yellow inputs in **A2:B11** (up to 10 years) and **E1:E2**, and the results in **E3:E10**. **Check** shows the change versus the year before, sales at the average rate, and the best and weakest year. **How to use** has short notes. All tabs print portrait on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
