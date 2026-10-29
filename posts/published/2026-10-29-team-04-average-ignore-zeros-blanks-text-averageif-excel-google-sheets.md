---
threads_url: https://www.threads.com/@tin_ylab/post/DfF0jXVH26V
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-average-range-and-ignore-zeros.html
title: How to Average a Range and Ignore Zeros, Blanks and Text in Excel and Google Sheets
labels: excel-formulas, small-business, averages
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model. Checked: all zero, all blank and all text ranges, one value, negatives, text in a number cell, numbers stored as text, the text "0", an #N/A or #DIV/0! typed into the range, and two random mixed lists. Excel and Google Sheets were not used.
search_description: Average a range and ignore zeros, blanks and text in Excel and Google Sheets with AVERAGEIF. Free template with 28 sample market days, a days counted line and an all zero guard.
image_prompts: A paper cash tally sheet with a few stacks of dollar bills and coins beside a closed cash box and a pencil on a market stall table, soft morning light, no screens, no readable text, no logos
image_alt: Stacks of bills and coins on a stall table beside a closed cash box
threads: To average only the days that had sales: =AVERAGEIF(B5:B32,">0")\nBlanks and text get skipped, zeros fall out through the ">0" rule.\nIf every day is zero it throws #DIV/0!, so the template wraps it in an IF that shows "No days" instead.
---
Average only the days above zero with `=AVERAGEIF(B5:B32,">0")`, where B5:B32 holds the daily sales. Blanks and text are skipped automatically, and zero days are skipped by the `">0"` rule.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same older functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Twenty-eight market days of sales beside a plain average, a zero-skipping average and days counted](images/2026-10-29-team-04-average-ignore-zeros-blanks-text-averageif-excel-google-sheets/average-ignore-zeros-daily-sales-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

`AVERAGE` counts a $0.00 day as a day, which pulls the result down. `AVERAGEIF` with `">0"` leaves those days out, and it also ignores blank cells and text.

A real $0.00 day may belong in your average, for example a day you were open and sold nothing. Whether to count it is your call. The template shows both numbers side by side.

The layout below matches the template. Day numbers go in column A and daily sales go in **B5:B32**.

### In Excel

1. In **B5:B32**, type the daily sales. Leave a day empty if you have no figure for it.
2. In **E5**, type `=AVERAGE(B5:B32)`. This counts zero days.
3. In **E6**, type `=AVERAGEIF(B5:B32,">0")`. This is the average of days above zero.
4. In **E9**, type `=COUNTIF(B5:B32,">0")`. This is how many days fed the average.
5. In **E10**, type `=COUNTIF(B5:B32,0)`. This is how many days sit at zero.
6. In **E8**, type `=IF(COUNTIF(B5:B32,">0")=0,"No days",AVERAGEIF(B5:B32,">0"))`. This shows No days instead of #DIV/0! when nothing is above zero.

Microsoft's page: [AVERAGEIF function](https://support.microsoft.com/en-us/office/averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642). That page describes the function, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 6 above, typed exactly as written.

Google's help page: [AVERAGEIF function](https://support.google.com/docs/answer/3256529). That page describes the function, not our test. None of these formulas needs an array formula.

### The same average without AVERAGEIF

The template's **E7** is `=IF(E9=0,"No days",SUMIF(B5:B32,">0")/E9)`. It adds only the days above zero and divides by the days counted, so it matches E6. Using `SUMIF` here matters: a plain `SUM` would also add negative values.

## Example

Sample data is fictional: 28 made up market days with US dollar sales.

| Cell | What it shows | Result |
|---|---|---|
| E5 | `AVERAGE`, zeros counted | $157.29 |
| E6 | `AVERAGEIF(B5:B32,">0")` | $204.48 |
| E7 | `SUMIF` divided by days counted | $204.48 |
| E8 | Guarded average | $204.48 |
| E9 | Days counted | 20 |
| E10 | Days at zero | 6 |
| E12 | Text entries | 1 |
| E13 | Blank cells | 1 |
| E14 | Check | OK |

The plain average is $47.19 lower than the other. Of the 28 days, 20 had sales, 6 were $0.00, 1 held the text rain and 1 was blank. 20 + 6 + 1 + 1 = 28, and **E14** says OK when that adds up.

Edge cases from the test, all in LibreOffice Calc:

- With sales of 10, -5, 20, -30 and 0, `AVERAGE` gave -1 and `AVERAGEIF(">0")` gave 15. Negative values are skipped by `">0"` and counted separately in **E11**.
- With one value of $50.50, every average gave $50.50.
- If every day is $0.00, `AVERAGEIF` returns #DIV/0!. The guard in E8 shows No days.

## Troubleshooting

### E6 shows #DIV/0!
Nothing in the range is above zero, so there is nothing to average. The range may be all zero, all blank or all text. Use the guarded formula in **E8**, which shows No days.

### A number is missing from the average
It may be stored as text. In the test, "45" and "50" typed as text were skipped by `AVERAGE`, `AVERAGEIF` and `COUNTIF`, and counted as text in **E12**. Convert them to numbers and they join the average.

### E5 shows #N/A but E6 still shows a number
There is an error value inside the range. In LibreOffice Calc, `AVERAGEIF(">0")` skipped the error cell and still returned a number, so only E5 and the check cell E14 gave it away. The template keeps a small error demo in **G5:G9**: `AVERAGE` gave #N/A and `AVERAGEIF(">0")` gave $41.83 from 40, 55.5 and 30. That is LibreOffice's result only. We did not check how Excel or Google Sheets treat errors inside `AVERAGEIF`, so fix the error cell instead of relying on it.

### E14 says CHECK
The counts do not add up to the number of rows, or the range holds an error value. Look for an error in the sales cells first.

## Template

[Download the .xlsx](templates/tidy-tabs-average-ignore-zeros.xlsx)

The file has two tabs. **Daily sales** holds yellow sales cells in **B5:B32**, with 28 fictional days filled in, and the yellow error demo in **G5:G9**. Gray cells are formulas in **E5:E14** and **E17:E18**. **How to use** has short notes.

The template does not decide whether a $0.00 day belongs in your average. It was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
