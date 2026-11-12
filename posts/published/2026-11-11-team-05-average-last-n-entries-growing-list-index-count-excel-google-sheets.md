---
threads_url: https://www.threads.com/@tin_ylab/post/DfpIbN_m89N
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-average-last-5-entries-in.html
title: How to Average the Last 5 Entries in a Growing List in Excel and Google Sheets
labels: excel-formulas, small-business, freelancers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 33 workbooks and 6,831 comparisons. Checked: 0, 1, 4, 5, 6, 14 and 200 entries with N = 5, N of 1, 2, 5, 13, 14, 15, 200, 201 and 1000000, invalid settings (0, -1, 2.5, text, "5" typed as text, blank), a text cell and blank cells inside the list, entries that start lower on the sheet, a text cell in the last row, and 8 random lists. Excel and Google Sheets were not used. Not run: a non-US language setting, more rows than the slots.
search_description: Average the latest 5 entries of a growing list with INDEX, MAX and LOOKUP in Excel and Google Sheets, no OFFSET needed. Free template with sample sales.
image_prompts: A row of weekly paper receipts clipped to a corkboard beside a pencil and a small stack of unclipped receipts, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Weekly paper receipts clipped in a row to a corkboard beside a pencil
threads: Average of your last 5 sales, and the list grows every week?\n=AVERAGE(INDEX(B4:B203,MAX(1,F5-3-$B$1+1)):INDEX(B4:B203,F5-3)) follows the newest row, no OFFSET.\nIn the sample, 14 weeks average $1,439.44 but the last 5 average $1,632.93 (made-up sales).
---
Average the latest N entries with `=AVERAGE(INDEX(B4:B203,MAX(1,F5-3-$B$1+1)):INDEX(B4:B203,F5-3))`, where B1 holds N and F5 holds the row of the last number. With 14 sample weekly sales and N = 5, it averages entries 10 to 14 and shows $1,632.93, next to $1,439.44 for all 14. This describes numbers already typed. It does not forecast.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (33 workbooks, 6,831 comparisons). The formulas use only IF, AND, OR, NOT, ISNUMBER, COUNT, COUNTIF, LOOKUP, INDEX, MAX, ROW, INT and AVERAGE, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Fourteen weekly sales amounts with an average of the latest five beside each one](images/2026-11-11-team-05-average-last-n-entries-growing-list-index-count-excel-google-sheets/average-last-5-entries-template.png) *Fourteen weekly sales with the average of the last 5 on every row and the summary at the right. PDF export from LibreOffice Calc 24.2. Sample data is fictional.*

## Steps

The steps are the same in Excel and Google Sheets. Put the dates in column **A** and the amounts in column **B**, with headers in row 3 and the first entry in row 4. Type new entries in the next empty row, top to bottom.

1. In **B1**, type how many latest entries to average, such as `5`. It must be a whole number of 1 or more.
2. In **F5**, find the row of the newest number: `=LOOKUP(2,1/ISNUMBER(B4:B203),ROW(B4:B203))`. `1/ISNUMBER(...)` gives 1 for a number and an error for anything else. LOOKUP skips the errors and returns the row of the last number.
3. In **F6**, type `=AVERAGE(INDEX(B4:B203,MAX(1,F5-3-$B$1+1)):INDEX(B4:B203,F5-3))`.
4. Read F6 from the inside out. `F5-3` turns a sheet row into a position in B4:B203, because the list starts below header row 3. Subtracting `$B$1-1` more gives the position N entries back, and MAX keeps it at 1 when the list is shorter than N.
5. The two INDEX calls return the first and last cell of the window. The colon joins them into a range, and AVERAGE does the rest.
6. For an average on every row, type this in **C4** and fill it down: `=AVERAGE(INDEX($B$4:$B$203,MAX(1,ROW(B4)-3-$B$1+1)):B4)`. The template wraps it so empty rows stay blank.

The 3 in these formulas is the header row offset. The template writes it as `ROW($B$3)`, so the formulas still work if you insert rows above the list.

Many guides use OFFSET for this job. Microsoft documents [OFFSET](https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66) as volatile, which means it recalculates on every change anywhere in the workbook. `INDEX(...):INDEX(...)` is not volatile. See Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd), [LOOKUP function](https://support.microsoft.com/en-us/office/lookup-function-446d94af-663b-451d-8251-369d5e3864cb) and [AVERAGE function](https://support.microsoft.com/en-us/office/average-function-047bac88-d466-426c-a32b-8f33eb960cf6). They describe the functions, not our test.

## Example

Sample data is fictional: 14 weekly sales from 08/03/2026 to 11/02/2026, with N = 5.

| Date | Amount | Average of last 5 |
|---|---|---|
| 08/03/2026 | $1,180.00 | $1,180.00 |
| 08/10/2026 | $1,245.50 | $1,212.75 |
| 08/17/2026 | $1,310.25 | $1,245.25 |
| 09/21/2026 | $1,510.00 | $1,355.23 |
| 10/19/2026 | $1,702.50 | $1,541.27 |
| 11/02/2026 | $1,788.90 | $1,632.93 |

The summary for the whole sample:

| Item | Value |
|---|---|
| Entries | 14 |
| Row of the latest entry | 17 |
| Average of last 5 | $1,632.93 |
| Average of all entries | $1,439.44 |
| Numbers in the last 5 rows | 5 |
| Cells in column B that are not numbers | 0 |
| CHECK | OK |

The latest entry is in row 17, so the window runs from entry 10 ($1,620.00 on 10/05/2026) to entry 14 ($1,788.90 on 11/02/2026). Those five amounts add up to $8,164.65, and $8,164.65 divided by 5 is $1,632.93.

The first rows average what exists so far, so the first row shows its own amount. With N larger than 14, the average of the last N equals the average of all entries.

## Troubleshooting

### Every row and F6 show Check
B1 is not usable. In the test, a blank, text, 0, a negative number and 2.5 all gave Check. So did "5" typed as text. Retype it as a plain whole number.

### The window holds fewer than N numbers
The window is the last N rows ending at the newest number, not the last N numbers. A blank or text cell inside it is skipped by AVERAGE. F8 counts the numbers actually used, so a 4 there with N = 5 means one cell in the window is empty or text.

### CHECK shows Check but N is fine
Something in column B is not a number. F9 counts those cells. In the test, a text cell in the last row did this. Retype the amount as a plain number, without stray spaces or text.

### New entries are ignored
The template reads rows 4 to 203. An entry in row 204 or lower is outside the range. Extend the ranges in F5, F6 and the other formulas if you need more rows.

### A gap row in the middle
A blank row between entries is fine. It is skipped in the average, but it still takes one of the N rows in the window.

## Template

[Download the .xlsx](templates/tidy-tabs-average-last-n.xlsx)

The file has two tabs. **Log** has the number of latest entries in **B1**, and 200 slots from row 4 to row 203 with 14 fictional weekly sales. Column **A** is the date, column **B** the amount, and column **C** the average of the last N on every row. A summary in **E3:F10** shows the entries, the row of the latest entry, the average of the last N, the average of all entries, the numbers in the window, cells in column B that are not numbers, and a CHECK cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The template describes past numbers only, and it does not predict next week.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
