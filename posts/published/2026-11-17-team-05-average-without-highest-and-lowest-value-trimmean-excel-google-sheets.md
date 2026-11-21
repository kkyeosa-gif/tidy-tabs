---
threads_url: https://www.threads.com/@tin_ylab/post/DgA1HOCieoX
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-average-without-highest-and.html
title: How to Average Without the Highest and Lowest Value in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 28 workbooks and 3,192 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Drop the highest and lowest value with =(SUM(r)-MAX(r)-MIN(r))/(COUNT(r)-2) or =TRIMMEAN(r,2/COUNT(r)). Why TRIMMEAN 0.1 drops nothing. Free template.
image_prompts: A rain-soaked folded canopy tent leaning against a garage wall after a wet market day, two wooden crates of unsold apples stacked beside it, and a canvas zip bank pouch lying on top of the crates, gray overcast light, realistic photo, no screens, no readable text, no logos
image_alt: Wet folded canopy in a garage beside crates of unsold apples and a bank pouch
threads: 10 made-up Saturdays, TRIMMEAN at 0.1, 0 days dropped.\nIn LibreOffice Calc the drop count rounds down to an even number, so it returned the plain average, $674.10.\n=TRIMMEAN(r,2/COUNT(r)) drops one high and one low day: $641.63.
---
Drop the single highest and lowest value and average the rest with `=(SUM(r)-MAX(r)-MIN(r))/(COUNT(r)-2)`, or use `=TRIMMEAN(r,2/COUNT(r))`, which drops the same two values. In the made-up log of 10 Saturdays with one $1,480.00 holiday and one $128.00 rain day, the plain average is $674.10 and the average without the highest and lowest is $641.63, while `TRIMMEAN(r,0.1)` on the same list returns $674.10 and drops nothing.

This is also called a trimmed mean. It answers "what does a typical day look like once the one best and one worst day are set aside", which is a different question from the plain average or the median. None of the three is a forecast.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (28 workbooks, 3,192 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (SUM, MAX, MIN, COUNT, AVERAGE, MEDIAN, TRIMMEAN) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Ten market days of sales beside plain average, median, trimmed average and TRIMMEAN results](images/2026-11-17-team-05-average-without-highest-and-lowest-value-trimmean-excel-google-sheets/average-without-highest-and-lowest-template.png) *Ten fictional market Saturdays with a $1,480.00 holiday rush and a $128.00 rain day: plain average $674.10, without the highest and lowest $641.63, and TRIMMEAN at 0.1 still $674.10 (the Status column and check rows are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Sales** tab, one market day per row starting in row 5.

1. Type each market date in **A5** and down (MM/DD/YYYY, such as 09/05/2026).
2. Type that day's sales as a plain number in **B5** and down: 612, not "$612.00" typed as text. Format the column as currency if you want the dollar sign.
3. Leave a closed day blank. A 0 is a real number, so it would count as the lowest day and get dropped in place of your real worst day.
4. In **G9**, type `=(SUM(B5:B104)-MAX(B5:B104)-MIN(B5:B104))/(COUNT(B5:B104)-2)`. SUM adds every day, MAX and MIN take out the best and worst, and COUNT-2 divides by the days that are left.
5. In **G10**, type `=TRIMMEAN(B5:B104,2/COUNT(B5:B104))` as a check. TRIMMEAN takes a percent of all values to drop, split between the top and the bottom. Two values out of the count is the percent that drops exactly one from each end, and `2/COUNT` works it out again as you add days.
6. Compare **G9** and **G10**. They should match. The template shows OK in **G11** when they do.
7. To drop the top 2 and bottom 2 instead, use `=TRIMMEAN(B5:B104,4/COUNT(B5:B104))` once you have 5 or more days (the template has it in **G12**).

Why not just type a percent? In LibreOffice Calc, `TRIMMEAN(B5:B104,0.1)` on 10 days returned the plain average, $674.10. Ten percent of 10 values is 1, and the number of values dropped was rounded down to an even number, so 0 were dropped. Microsoft's [TRIMMEAN function](https://support.microsoft.com/en-us/office/trimmean-function-d90c9878-a119-4746-88fa-63d988f511d3) page describes the same rounding for Excel: the number of excluded points is rounded down to the nearest multiple of 2. With 15 days, 0.2 is 3 values, rounded down to 2, so LibreOffice dropped one from each end. The `2/COUNT` version matched the SUM, MAX and MIN version for every count from 3 to 100 values in LibreOffice.

For the middle value instead, see Microsoft's [MEDIAN function](https://support.microsoft.com/en-us/office/median-function-d0916313-4753-414c-8537-ce85bdd967d2) page and the Tidy Tabs post on [finding the median order value](https://tidytabs.blogspot.com/2026/10/how-to-find-median-order-value-in-excel.html). To flag every unusual day rather than drop one at each end, see [finding unusually high or low order amounts](https://tidytabs.blogspot.com/2026/10/how-to-find-unusually-high-or-low-order.html).

## Example

The market days and sales are fictional.

| Market date | Sales | Note |
|---|---|---|
| 09/05/2026 | $612.00 | |
| 09/12/2026 | $655.00 | |
| 09/19/2026 | $590.00 | |
| 09/26/2026 | $1,480.00 | Holiday rush |
| 10/03/2026 | $640.00 | |
| 10/10/2026 | $128.00 | Rained out |
| 10/17/2026 | $700.00 | |
| 10/24/2026 | $668.00 | |
| 10/31/2026 | $623.00 | |
| 11/07/2026 | $645.00 | |

The summary in **F4:G17** shows:

| Summary | Value |
|---|---|
| Days counted | 10 |
| Plain average | $674.10 |
| Median | $642.50 |
| Highest day / lowest day | $1,480.00 / $128.00 |
| Average without highest and lowest | $641.63 |
| TRIMMEAN check, 2/COUNT | $641.63 (The two match? OK) |
| Drop top and bottom 2 | $640.50 |
| TRIMMEAN(B5:B104,0.1), for comparison | $674.10 |

The trimmed average holds 641.625 in the cell. Rows skipped, Rows showing Check and Typed below row 104 all show 0, and CHECK shows OK.

Ties count once. With 500, 500, 600, 700, 700, only one 500 and one 700 are dropped, and the result is 600.

## Troubleshooting

### TRIMMEAN gives the plain average
The percent is too small for the count. `TRIMMEAN(r,0.1)` on 10 values dropped nothing in LibreOffice and returned $674.10. Use `2/COUNT(r)` as the percent.

### The trimmed average dropped a closed day, not the real low day
A 0 was typed for a day the market was closed. In a copy of the template with a 0 added, the lowest day became $0.00, the rained-out $128.00 stayed in, and Status showed Check: 0 or less. Clear the 0 and leave the cell blank.

### Days counted is lower than the rows you typed
A sales amount was typed as text, such as "$612.00" stored as text. COUNT, SUM and TRIMMEAN skip it without a warning, so Days counted showed 9 and Status showed Skipped: text. Retype it as a plain number.

### #DIV/0! or #VALUE! with only two days
With two values there is nothing left after dropping the highest and lowest. In LibreOffice the SUM version gave #DIV/0! and TRIMMEAN gave #VALUE!. The template shows Need 3+ days in **G9** until there are at least 3 days.

## Template

[Download the .xlsx](templates/tidy-tabs-average-without-highest-lowest.xlsx)

The file has two tabs. **Sales** has rows 5 to 104 with Market date, Sales and Note (inputs in **A:C**) and a Status in **D**: OK, Skipped: blank, Skipped: text or Check: 0 or less. The summary in **F4:G17** holds Days counted, Plain average, Median, Highest day, Lowest day, Average without highest and lowest, the TRIMMEAN check, The two match?, Drop top and bottom 2, TRIMMEAN with 0.1 for comparison, Rows skipped, Rows showing Check, entries typed below row 104 and the CHECK. **How to use** explains each formula and the 0.1 trap. The Sales tab prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Skipped rows turn orange, and Check: 0 or less and a CHECK that is not OK turn red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
