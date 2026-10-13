---
threads_url: https://www.threads.com/@tin_ylab/post/Decn1fpgdPk
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-total-sales-by-day-of-week-in.html
title: How to Total Sales by Day of the Week in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python math, including a date stored as text, a $0.00 day and an empty sheet. Excel and Google Sheets were not opened; the one help link is from Microsoft Support.
search_description: Name each weekday with =TEXT(A2,"ddd"), then =AVERAGEIF($C$2:$C$29,"Sat",$B$2:$B$29). Sample month: Saturdays average $159.19, Sundays $229.25. Free .xlsx.
image_prompts: Small farmers market stall table with a metal cash box, a paper bag of goods and a wall calendar page leaning behind, soft morning light, no screens, no readable text, no logos
image_alt: Market stall table with a cash box and a paper bag in front of a wall calendar page
threads: Which day of the week actually pays your bills?\nName each date with =TEXT(A2,"ddd"), then =AVERAGEIF($C$2:$C$29,"Sat",$B$2:$B$29).\nIn the sample month Saturdays average $159.19 and Sundays $229.25. Compare averages, not totals.
---
Name each date's weekday with `=TEXT(A2,"ddd")`, then average one day with `=AVERAGEIF($C$2:$C$29,"Sat",$B$2:$B$29)`. In the sample month, Saturdays average $159.19 and Sundays $229.25.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Weekday summary table with days counted, total sales, average per day and share of sales for each day](images/2026-10-13-team-02-sales-by-day-of-week-sumif-excel-google-sheets/sales-by-day-of-week-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: add a column that turns each date into Mon, Tue and so on, then sum and average the sales for each of those names.

1. In row 1 of a tab named Daily sales, type these headers in A to C: Date, Sales, Weekday.
2. In **A2:A29**, type one real date per row, such as 09/01/2026. Use one row per day.
3. In **B2:B29**, type that day's sales, such as $42.00.
4. In **C2**, type `=IF(A2="","",TEXT(A2,"ddd"))` and fill it down. It shows Tue for 09/01/2026 and stays blank when the date is empty.
5. On a second tab named By weekday, type Mon through Sun in **A2:A8**, one per row. Add the headers Weekday, Days counted, Total sales, Average per day, Share of sales in A1:E1.
6. In **B2**, type `=COUNTIF('Daily sales'!$C$2:$C$200,A2)` to count the days.
7. In **C2**, type `=SUMIF('Daily sales'!$C$2:$C$200,A2,'Daily sales'!$B$2:$B$200)` to add that weekday's sales.
8. In **D2**, type `=IF(B2=0,"",C2/B2)`. This is the average per day.
9. In **E2**, type `=IF($C$9=0,"",C2/$C$9)` for each weekday's share of the total.
10. Fill **B2:E2** down to row 8. In **B9** and **C9**, type `=SUM(B2:B8)` and `=SUM(C2:C8)`.

Compare the average column, not the total. If your data covers 5 Saturdays and 4 Mondays, the Saturday total is bigger even if every Saturday was slow.

For one weekday, skip the table and use the short form from the top of this post. Change `"Sat"` to any other three-letter name.

Add two check cells so a bad row does not slip by. `=IF(B9=COUNT('Daily sales'!A2:A200),"OK","CHECK")` confirms every real date was counted. `=IF(ROUND(C9-SUM('Daily sales'!B2:B200),2)=0,"OK","CHECK")` confirms the weekday totals add up to all sales.

### In Excel

The formulas work as typed. Microsoft documents the format code on its [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c) page, where `"ddd"` gives the short day name. Format the sales cells from **Home > Number Format > Currency**.

### In Google Sheets

The same formulas should work as typed. Format the sales cells with **Format > Number > Currency**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is 28 days of sales from a small market stall, from 09/01/2026 (a Tuesday) to 09/28/2026, so each weekday appears 4 times.

| Weekday | Days counted | Total sales | Average per day | Share of sales |
|---|---|---|---|---|
| Mon | 4 | $378.50 | $94.63 | 13.9% |
| Tue | 4 | $165.50 | $41.38 | 6.1% |
| Wed | 4 | $150.75 | $37.69 | 5.5% |
| Thu | 4 | $202.75 | $50.69 | 7.5% |
| Fri | 4 | $265.25 | $66.31 | 9.8% |
| Sat | 4 | $636.75 | $159.19 | 23.4% |
| Sun | 4 | $917.00 | $229.25 | 33.8% |
| **Total** | **28** | **$2,716.50** | | |

Sunday is the best day at $229.25 on average. Sunday and Saturday together bring in 57.2% of the month's sales. Both checks showed OK.

The averages are shown to cents. The unrounded Saturday value is 159.1875, and the results matched a Python check.

## Troubleshooting

### A weekday shows more days than the others
A date stored as text can still get a weekday name. In the recalculated file, "09/29/2026" typed as text with $100.00 made Tue show 5 days and $265.50. The days counted (30) no longer matched the count of real dates (29), so the first check showed CHECK. Re-enter the date as a real date.

### A day with $0.00 pulls the average down
A row with a date and $0.00 counts as a day. In the recalculated file, adding 09/30/2026 with $0.00 gave Wed a fifth day and lowered its average. That is correct if the shop was open. Delete the row if it was closed.

### The weekday table shows 0 and blanks
The Weekday column is empty or the names do not match. In the recalculated file, clearing every row gave counts and totals of 0, blank averages and shares, and no Best day. Fill down the formula in column C and check that A2:A8 reads Mon to Sun.

### The weekday names do not match A2:A8
`"ddd"` returns day names in the language of the file. This template assumes a US-English file with Mon to Sun. A different locale returns other names, and COUNTIF finds nothing. Type the names your file returns in A2:A8.

### Rows past 200 are ignored
The ranges stop at row 200, and the template's helper formulas are filled to row 60. For a longer list, fill column C further down and extend the ranges.

## Template

[Download the .xlsx](templates/tidy-tabs-sales-by-weekday.xlsx)

The file has three tabs. **Daily sales** holds the 28 fictional days in A2:B29 and the weekday formula in C2:C60. **By weekday** has the table above, a Best day marker in column F, and the two checks in F11 and F12. **How to use** has short fill-in notes. Each tab is set to US Letter, fit to one page wide.

The dates and sales are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.

It assumes one row per date and covers sales totals only. It does not tell you why a day is slow.
