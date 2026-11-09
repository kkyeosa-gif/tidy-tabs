---
threads_url: https://www.threads.com/@tin_ylab/post/DfhaCSGm8QQ
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-find-date-you-hit-sales-goal-in.html
title: How to Find the Date You Hit a Sales Goal in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 21 workbooks and 1,232 comparisons with 0 mismatches. Checked: the sample, a goal equal to a running total, a goal of 0, a blank, text and negative goal, a goal above the total, a first day already over the goal, refunds that pull the total back under the goal, blank days, a text amount, unsorted dates, an empty sheet, 366 days with and without reaching the goal, and five random sets with refunds. Excel and Google Sheets were not used. Not run: TODAY(), a non-US language setting, more rows than the 366 slots, date serials below 61.
search_description: Find the first date a running total reaches a sales goal in Excel and Google Sheets with SUM, a Yes flag and INDEX/MATCH. Free template included.
image_prompts: A tall glass jar half full of coins and folded bills, a red marker and a poster board with a hand-drawn blank thermometer, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Glass jar of coins beside a poster board with a hand-drawn thermometer
threads: Find the day your running total passed the goal with a Yes flag.\n=IF(C4>=$B$1,"Yes","") on each row, then MATCH("Yes",$D$4:$D$369,0) finds the first one. In the sample, a $10,000 goal is passed on 09/12/2026.\nThe COUNTIF shortcut goes wrong when refunds pull the total back under.
---
Add a running total with `=SUM($B$4:B4)`, mark the first row that reaches the goal with `=IF(C4>=$B$1,"Yes","")`, and return its date with `=IFERROR(INDEX($A$4:$A$369,MATCH("Yes",$D$4:$D$369,0)),"Not yet")`. With a $10,000.00 goal, the sample total first passes it on 09/12/2026, day 12, at $10,612.00 (fictional sales).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (21 workbooks, 1,232 comparisons). The formulas use only SUM, IF, INDEX, MATCH, IFERROR and COUNTA, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Daily sales with running total and the first day passing a ten thousand dollar goal](images/2026-11-08-team-05-date-you-hit-sales-goal-cumulative-total-match-excel-google-sheets/date-you-hit-sales-goal-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional.*

## Steps

The steps are the same in Excel and Google Sheets. On **Daily sales**, the goal is in **B1**, a goal check is in **B2**, the headers are in row 3 and the data starts in row 4: date in **A**, sales for that day in **B**.

1. In **B1**, type your goal, such as 10000.
2. From row 4, enter one row per day in date order, oldest first. A day with no sales can stay blank or hold 0.
3. In **C4**, type `=SUM($B$4:B4)`. Fill it down. The first cell of the range is locked and the last one moves, so each row adds everything so far.
4. In **D4**, type `=IF(C4>=$B$1,"Yes","")`. Every row from the goal date on says Yes, and that is fine. Only the first one matters.
5. In any cell, type `=IFERROR(INDEX($A$4:$A$369,MATCH("Yes",$D$4:$D$369,0)),"Not yet")`. `MATCH` with 0 finds the first Yes, `INDEX` returns the date on that row, and `IFERROR` shows Not yet if the goal was never reached. Format the cell as a date.
6. For the day count, type `=H3-$A$4+1` where H3 holds the date. The first row is day 1, so this assumes row 4 is the start of your period.

The template wraps these in a few more checks so a blank or text goal shows Check instead of Not yet.

Help pages: Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [SUM function](https://support.microsoft.com/en-us/office/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89), and Google's [MATCH function](https://support.google.com/docs/answer/3093378). They describe the functions, not our test.

### Why the Yes flag instead of a one-cell test

A common shortcut is `MATCH(TRUE,INDEX(C4:C369>=B1,0),0)`. We did not use it. Empty rows hold a blank text result, and in Excel text counts as larger than any number, so empty rows would read as TRUE. We did not run that formula, so we make no claim about it. Matching a Yes flag avoids the problem and is easier to read.

### A shorter formula that breaks with refunds

You may see `=INDEX(A4:A369,COUNTIF(C4:C369,"<"&B1)+1)`. It counts how many running totals are below the goal and takes the next row. That works only if the total never goes down.

With refunds (negative amounts) the total can dip back under the goal. In our test, sales of 1000, 700, -600, 300 and 900 with a goal of 1500 first reach it on day 2, but the COUNTIF version returned day 4. A second test (2500, -1000, 400, 2000 with a goal of 2000) gave day 1 for the Yes flag and day 3 for COUNTIF. If you record refunds as negative sales, use the Yes flag.

## Example

Sample data is fictional: 14 days from 09/01/2026 to 09/14/2026, goal $10,000.00. Four of the rows:

| Date | Sales | Running total | Goal reached? |
|---|---|---|---|
| 09/10/2026 | $640.75 | $8,376.50 | |
| 09/11/2026 | $1,510.00 | $9,886.50 | |
| 09/12/2026 | $725.50 | $10,612.00 | Yes |
| 09/14/2026 | $1,175.00 | $12,767.00 | Yes |

The summary block shows:

- Date you hit the goal: 09/12/2026
- Days it took: 12
- Running total that day: $10,612.00
- Amount over: $612.00
- Total of all rows: $12,767.00

If the goal equals a running total exactly, that day counts, because the test is `>=`. A goal of $6,415.75 returns 09/08/2026. A goal of 99,999 returns Not yet and shows $87,232.00 still to go.

## Troubleshooting

### The goal cell says Check
The goal in **B1** is blank, text or negative. Type a plain number such as 10000, without a dollar sign in quotes. While it says Check, the date cell says Check too, not Not yet.

### A row says Check in the row check column
The date is not a real date or the sales amount is text, such as "1,250.00" typed with quotes or pasted as text. `SUM` ignores text, so that day would count as zero and the total would be too low. Retype the value as a number.

### A row says Out of order
The date is earlier than the row above it. The first Yes is still found by position, but the Days it took figure is no longer reliable. Sort the rows by date, oldest first.

### It says Not yet but you think the goal was met
The goal may be higher than the total of all rows, or the rows past the last one are not filled in. Check the total of all rows in the summary. The template holds 366 days, so a longer period needs more rows.

### The date is a five digit number
The result cell is formatted as a number. Format it as a date.

## Template

[Download the .xlsx](templates/tidy-tabs-date-hit-sales-goal.xlsx)

The file has two tabs. **Daily sales** has the goal in **B1**, a goal check in **B2**, 366 day rows from row 4 with 14 fictional sample days, and a summary in **G3:H10**: the date you hit the goal, days it took, running total that day, amount over, total of all rows, still to go, rows needing attention and a **CHECK** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The file uses typed dates only, with no `TODAY()`. Printed, it runs 14 US Letter pages because of the 366 slots.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
