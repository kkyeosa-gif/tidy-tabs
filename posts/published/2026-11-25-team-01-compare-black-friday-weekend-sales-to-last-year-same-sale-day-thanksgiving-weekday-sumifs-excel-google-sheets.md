---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-compare-black-friday-sales-to.html
title: How to Compare Black Friday Sales to Last Year in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 360 workbooks and 1,501,037 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. Thanksgiving dates were checked against Python's datetime for every year from 1990 to 2045.
search_description: Line up Black Friday, Cyber Monday and the days around them with last year's by Thanksgiving, not by date: DATE, WEEKDAY and MOD, then SUMIFS per sale day.
image_prompts: A small gift shop counter during a holiday sale weekend with a tall stack of kraft paper wrapped gift boxes tied with red twine, a half unrolled roll of brown wrapping paper with scissors on it and a paper shopping bag stuffed with white tissue, warm indoor light, realistic photo, no screens, no readable text, no logos
image_alt: Stack of twine-tied gift boxes, wrapping paper roll and a tissue-filled shopping bag
threads: Comparing Black Friday to 11/28 last year? In 2024 that was Thanksgiving.\nLine up by days from Thanksgiving: =DATE(year,11,1)+MOD(5-WEEKDAY(DATE(year,11,1)),7)+21 finds it, and +1 is Black Friday.\nIn the sample shop, Black Friday comes to $705.21 vs $579.27, up 21.7%.
---
Find each year's Thanksgiving with `=DATE(year,11,1)+MOD(5-WEEKDAY(DATE(year,11,1)),7)+21`, add an offset (Black Friday is +1, Cyber Monday +4), and total each sale day with `=SUMIFS(log_amount,log_date,thanksgiving+offset)`. That lines up Black Friday 11/28/2025 with Black Friday 11/29/2024 instead of an ordinary 11/28/2024. In the made-up shop log, Black Friday comes to $705.21 against $579.27, up 21.7%.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (360 workbooks, 1,501,037 cell comparisons, 0 mismatches), plus the edge cases below. Thanksgiving dates were checked against Python's datetime for every year from 1990 to 2045. Excel and Google Sheets were not opened for this post.

![Eight Thanksgiving week sale days lined up with last year's matching days and sales](images/2026-11-25-team-01-compare-black-friday-weekend-sales-to-last-year-same-sale-day-thanksgiving-weekday-sumifs-excel-google-sheets/black-friday-vs-last-year-template.png) *Compare tab of a fictional candle shop log for 2025 vs 2024: each sale day is matched by Thanksgiving, so Black Friday 11/28/2025 ($705.21) sits next to Black Friday 11/29/2024 ($579.27), +21.7%, Sunday shows No entries this year, and the side block shows date minus 364 landing on the right Friday in this pair of years while EDATE lands on a Thursday. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The Thanksgiving rule here is the US one: the fourth Thursday of November. The sheet compares sales only, not profit, so ad spend and discounts are not in it.

The steps are the same in Excel and Google Sheets.

1. On the **Sales log** tab, type one sale per row from row 5: Date in **A**, Amount in **B**, optional Channel in **C**. Several rows per day are fine; a refund is a negative amount. Column **E** flags bad rows.
2. On the **Compare** tab, type the year to look at in **B1**. **B2** holds `=B1-1` (last year), and you can type over it with any other year.
3. **B3** finds Thanksgiving with `=IF($L$6<>"OK","",DATE(B1,11,1)+MOD(5-WEEKDAY(DATE(B1,11,1)),7)+21)`. `MOD(5-WEEKDAY(...),7)` counts the days from November 1 to the first Thursday, and +21 moves three weeks on to the fourth. **B4** does the same for B2. See the [WEEKDAY help from Microsoft](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a).
4. In rows 7 to 14, column **B** holds the days from Thanksgiving: -3 for the Monday before through 4 for Cyber Monday. Column **A** is only a label, so "Small Business Saturday" is just a name. Add rows down to row 26 to widen the window.
5. **C7** gives this year's date with `=IF(OR($L$6<>"OK",NOT(ISNUMBER(B7))),"",IF(B7<>INT(B7),"",$B$3+B7))`, and **D7** the compare year's date with `=IF(C7="","",$B$4+B7)`.
6. **E7** totals that day's sales with `=IF(C7="","",SUMIFS('Sales log'!$B$5:$B$2004,'Sales log'!$A$5:$A$2004,C7))`. **F7** is the same with D7. See the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496).
7. **G7** gives Change $ with `=IF(OR(C7="",LEFT(I7,2)="No"),"",ROUND(E7-F7,2))` and **H7** Change % with `=IF(OR(C7="",LEFT(I7,2)="No"),"",IF(F7=0,"",E7/F7-1))`.
8. **I7** shows OK, a Check message, or "No entries this year" (or compare year, or either year) when the log has no row for that date.
9. Read the summary in **K6:L15**: both totals, the change, days with no entries (counted as $0.00) and a final CHECK cell.

For 2026, type 2026 in **B1** once this year's sales are in the log. Thanksgiving 2026 is 11/26/2026, so Black Friday is 11/27/2026 and lines up with 11/28/2025.

For calendar months, see [How to Compare This Year's Sales to Last Year Month by Month](https://tidytabs.blogspot.com/2026/10/how-to-compare-this-years-sales-to-last.html). The Thanksgiving formula uses the method from [How to Find the Date of Any Nth Weekday of a Month](https://tidytabs.blogspot.com/2026/10/how-to-find-date-of-2nd-saturday-or-any.html).

## Example

The sales below are made up: a candle shop's Etsy sales for 11/15 to 12/06 of 2023, 2024 and 2025, plus a market booth row on each Black Friday. Sunday 11/30/2025 was left out of the log on purpose. **B1** is 2025 and **B2** is 2024, so Thanksgiving is 11/27/2025 and 11/28/2024.

| Sale day | Offset | 2025 date | 2024 date | 2025 sales | 2024 sales | Change $ | Change % | Status |
|---|---|---|---|---|---|---|---|---|
| Monday before Thanksgiving | -3 | 11/24/2025 | 11/25/2024 | $274.45 | $181.60 | $92.85 | 51.1% | OK |
| Tuesday | -2 | 11/25/2025 | 11/26/2024 | $183.96 | $211.91 | -$27.95 | -13.2% | OK |
| Wednesday | -1 | 11/26/2025 | 11/27/2024 | $259.84 | $239.53 | $20.31 | 8.5% | OK |
| Thanksgiving | 0 | 11/27/2025 | 11/28/2024 | $246.15 | $230.71 | $15.44 | 6.7% | OK |
| Black Friday | 1 | 11/28/2025 | 11/29/2024 | $705.21 | $579.27 | $125.94 | 21.7% | OK |
| Small Business Saturday | 2 | 11/29/2025 | 11/30/2024 | $277.73 | $219.67 | $58.06 | 26.4% | OK |
| Sunday | 3 | 11/30/2025 | 12/01/2024 | $0.00 | $245.93 | | | No entries this year |
| Cyber Monday | 4 | 12/01/2025 | 12/02/2024 | $187.71 | $223.35 | -$35.64 | -16.0% | OK |

The summary shows 8 days compared, $2,135.05 this year against $2,131.97, up $3.08 (0.1%), with 1 day with no entries and CHECK OK.

The trap block in **K17:N22** shows why the shortcuts slip. With **B1** set to 2024, Black Friday 11/29/2024 minus 364 days is 12/01/2023, a Friday one week late, which pulls $180.24 instead of Black Friday 2023's $222.00. `EDATE(11/29/2024,-12)` gives 11/29/2023, a Wednesday, with $272.77.

## Troubleshooting

### Black Friday is compared with Thanksgiving
The formula matches 11/28 to 11/28 or uses `EDATE(date,-12)`. With 2025 against 2024, that puts Black Friday 2025 ($705.21) next to Thanksgiving 2024 ($230.71). In Python, EDATE one year back from Black Friday landed on a Wednesday or Thursday every year from 2000 to 2040. Line up by days from Thanksgiving instead.

### Date minus 364 is a week off
Minus 364 keeps the weekday but not the sale week. In Python it gave the wrong Friday in 2002, 2008, 2013, 2019, 2024, 2030 and 2036 (2000 to 2040); in the template, 2024's Black Friday minus 364 is 12/01/2023, $180.24 instead of $222.00.

### A day shows $0.00
The log has no row for that date, so Status says "No entries this year" and Change $ and Change % stay blank instead of showing a -100% drop. In the sample that is 11/30/2025. Log the day, or ignore it in the comparison.

### A sale is missing from the total
Its date is text or has a time, and SUMIFS matches whole dates only. The Sales log Status shows "Check: date blank or text" or "Check: date has a time". Retype it as a plain date.

## Template

[Download the .xlsx](templates/tidy-tabs-black-friday-vs-last-year.xlsx)

**Compare** holds the years, both Thanksgiving dates, up to 20 sale days (rows 7 to 26), the summary and the trap block. **Sales log** takes 2,000 rows with its own CHECK cell, which also catches rows typed below row 2004. **How to use** explains each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. The Black Friday row is light green, "No entries" flags are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
