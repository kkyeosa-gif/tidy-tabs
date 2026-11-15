---
title: How to Check if Temperature and Market Sales Move Together With CORREL in Excel and Google Sheets
labels: excel-formulas, small-business, sales-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 31 workbooks and 9,734 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Measure how two columns move together with =CORREL(temps,sales), then add RSQ and SLOPE for dollars per degree. Free farmers market log template.
image_prompts: An outdoor farmers market booth under a white canopy tent on a hot summer morning, a round analog dial thermometer clipped to the tent leg, a gray metal cash box open on the table edge, and a galvanized bucket of cut sunflowers in the shade below, realistic photo, no screens, no readable text, no logos
image_alt: Dial thermometer clipped to a market tent leg above a cash box and sunflowers
threads: CORREL said 0.84. The first 9 rows alone said 0.96.\nIn a made-up log of 12 market Saturdays, sales climbed with the heat up to 79 F, then the two hottest days fell back to $860 and $815.\nOne number can blend two stretches. Split the log at a cutoff and look again.
---
Measure how closely two columns move together with `=CORREL(temps,sales)`, which returns one number from -1 to 1, and add `=SLOPE(sales,temps)` for the dollars per degree. In the made-up log of 12 Saturdays, CORREL is 0.84 and SLOPE is $6.12 per degree, while the first 9 rows alone (48 F to 79 F) give 0.96, a sign that sales drop again on the hottest days. The download counts the pairs used, so a skipped blank or "n/a" temperature shows up.

A high CORREL does not mean the heat brought in the sales. Crowds, events and the season can push both up at the same time, and 12 Saturdays is a small sample. Treat the number as a pattern in your log, not a cause and not a forecast.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (31 workbooks, 9,734 cell comparisons, 0 mismatches), plus the edge cases below. CORREL, RSQ, SLOPE and INTERCEPT also exist in Excel and Google Sheets, but neither app was opened for this post.

![Twelve Saturday market rows of high temperature and sales beside CORREL, RSQ and SLOPE](images/2026-11-16-team-05-check-if-temperature-and-market-sales-move-together-correl-rsq-excel-google-sheets/temperature-vs-market-sales-correl-template.png) *Twelve fictional Saturdays from 48 F to 90 F give CORREL 0.84 and SLOPE $6.12 per degree, and the days at or below 79 F alone give 0.96 (the Note and helper columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use one row per market day on the **Log** tab, starting in row 5.

1. Type the market date in **A5** (04/25/2026), the high temperature in F in **B5** (48) and the sales in **C5** (610), then fill your other days down to row 104 at most. Leave a temperature blank if you have none.
2. In **J4**, type `=SUMPRODUCT(ISNUMBER(B5:B104)*ISNUMBER(C5:C104))` to count the rows where both the temperature and the sales are numbers. These are the pairs CORREL actually uses.
3. In **J6**, type `=CORREL(B5:B104,C5:C104)`. Near 1 means sales tend to be higher on warmer days, near -1 lower, and near 0 no straight-line pattern. Words like strong or weak are rough labels, not rules.
4. In **J7**, type `=RSQ(C5:C104,B5:B104)`. It is the share of the ups and downs in sales that a straight line on temperature accounts for (0.70 in the sample, which is CORREL squared).
5. In **J8**, type `=SLOPE(C5:C104,B5:B104)` for the dollars of sales per 1 F along that line, and in **J9** type `=INTERCEPT(C5:C104,B5:B104)`. Put the sales range first in both. CORREL gives the same answer either way round, but SLOPE does not: with the ranges swapped, LibreOffice returned 0.1142 (degrees per dollar) instead of 6.1178.

The template wraps J6:J9 in `=IF(J4<3,"Need 3+ pairs",IFERROR(...,"Check"))`, so a log with fewer than 3 pairs or a column of equal values shows a word instead of an error.

To check whether one number hides a pattern, type a temperature in **J13** (79 in the file). Gray helper columns **F** and **G** copy the temperature and sales of each OK row at or below that cutoff, and **J14:J19** show the pairs, CORREL, SLOPE and average sales at or below it, plus the count and average sales of the days above it. No array formulas are needed.

Help pages: Microsoft's [CORREL function](https://support.microsoft.com/en-us/office/correl-function-995dcef7-0c0a-4bed-a3fb-239d7b68ca92) and [SLOPE function](https://support.microsoft.com/en-us/office/slope-function-11fb8f97-3117-4813-98aa-61d7e01276b9), and Google's [CORREL](https://support.google.com/docs/answer/3093990).

## Example

The dates, temperatures and sales are fictional.

| Market date | High temp (F) | Sales | Status |
|---|---|---|---|
| 04/25/2026 | 48 | $610.00 | OK |
| 05/02/2026 | 52 | $655.00 | OK |
| 05/09/2026 | 55 | $700.00 | OK |
| 05/16/2026 | 61 | $780.00 | OK |
| 05/23/2026 | 64 | $760.00 | OK |
| 05/30/2026 | 68 | $845.00 | OK |
| 06/06/2026 | 72 | $880.00 | OK |
| 06/13/2026 | 75 | $905.00 | OK |
| 06/20/2026 | 79 | $870.00 | OK |
| 06/27/2026 | 83 | $920.00 | OK |
| 07/04/2026 | 86 | $860.00 | OK |
| 07/11/2026 | 90 | $815.00 | OK |

The summary shows 12 pairs used, 0 rows skipped, 0 rows showing Check, and CHECK OK.

| Result | All 12 days | At or below 79 F |
|---|---|---|
| Pairs | 12 | 9 |
| CORREL | 0.8358 | 0.9648 |
| RSQ | 0.6986 | |
| SLOPE | $6.12 per 1 F | $9.53 per 1 F |
| INTERCEPT | $375.33 | |
| Average sales | | $778.33 |

The 3 days above 79 F averaged $865.00. Sales climb with the temperature up to about 79 F, then the two hottest Saturdays fall back to $860.00 and $815.00. CORREL on all 12 days blends both stretches into 0.84.

## Troubleshooting

### Pairs used is lower than the number of rows
A temperature or sales cell is blank or holds text, so that pair was left out of CORREL. In the LibreOffice test, sales typed as the text "$815.00" dropped the log to 11 pairs and changed CORREL to 0.9234 with no error. Look for Skipped: blank or Skipped: text in Status, and retype the value as a plain number.

### CHECK shows Check but CORREL looks fine
A temperature typed as "n/a" is text. CORREL skipped it and stayed at 0.8358 in the test, but the row shows Skipped: text and CHECK flags it. Clear the cell instead: a blank temperature on a rained-out day is skipped without turning CHECK to Check.

### CORREL shows #DIV/0! or Check
Every temperature or every sales value is the same, so there is no spread to compare. A bare CORREL gives #DIV/0!, and Microsoft's CORREL page says Excel does the same when the standard deviation of either set is zero. The template shows Check instead. Log more days with different values.

### Status shows Check: temp
The temperature is below -40 or above 130 F, often a date pasted into the wrong column. In the test, a date in the temperature column pulled CORREL down to 0.0477. Move the date to column A and type the temperature in B.

### The results show Need 3+ pairs
Fewer than 3 rows have both a temperature and sales as numbers. Add more market days. Even 12 is a small sample: a few unusual Saturdays can move the result a lot.

## Template

[Download the .xlsx](templates/tidy-tabs-correlation-sales-vs-temperature.xlsx)

The file has two tabs. **Log** holds rows 5 to 104: Market date, High temp (F), Sales and Note (yellow inputs in **A:D**), Status in **E**, and the gray cutoff helpers in **F:G**. The summary sits in **I4:J19**: pairs used, rows skipped, CORREL, RSQ, SLOPE, INTERCEPT, rows showing Check and CHECK in **J4:J11**, then the cutoff input in **J13** and the at-or-below and above results in **J14:J19**. **How to use** explains each part. The Log tab prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Status turns red for Check and orange for Skipped, and CHECK turns red when it is not OK. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
