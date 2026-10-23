---
threads_url: https://www.threads.com/@tin_ylab/post/De0Ocl4juqa
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-convert-yyyymmdd-numbers-like.html
title: How to Convert YYYYMMDD Numbers Like 20261021 to Real Dates in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python date model, including 16 edited variants of the date file (impossible dates, leap days, spaces, time zone suffixes, odd numbers). No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Convert YYYYMMDD numbers, text and ISO timestamps to real dates with DATE in Excel and Google Sheets. Free template flags impossible dates.
image_prompts: A wooden tray of small cardboard shipping boxes with blank paper tags tied on with string beside a rubber date stamp and an ink pad, soft window light, no screens, no readable text, no logos
image_alt: Shipping boxes with blank tags, a date stamp and ink pad on a wooden tray
threads: Export gave you 20261021 instead of a date?\n=DATE(INT(A2/10000),MOD(INT(A2/100),100),MOD(A2,100)) turns it into 10/21/2026.\nIf it is text, swap in LEFT, MID and RIGHT. The template flags impossible dates too.
---
Convert a number like 20261021 into a real date with `=DATE(INT(A2/10000),MOD(INT(A2/100),100),MOD(A2,100))`, then format the result as MM/DD/YYYY to get 10/21/2026. If the value is text, use `=DATE(LEFT(B2,4),MID(B2,5,2),RIGHT(B2,2))` instead.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Convert sheet with numbers and text like 20261021 beside the real dates they became](images/2026-10-22-team-03-convert-yyyymmdd-number-to-date-date-left-mid-excel-google-sheets/yyyymmdd-to-date-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

An export that gives 20261021 is a plain number or plain text, not a date. You cannot sort it by month or total it by month until it becomes a real date. DATE builds one from a year, a month and a day.

### If the value is a number

1. Type the numbers in **A2:A101**, for example 20261021 in A2.
2. In **D2**, type `=DATE(INT(A2/10000),MOD(INT(A2/100),100),MOD(A2,100))`.
3. Format **D2** as a date: **Format > Number > Date** in Sheets, or **Home > Number Format** in Excel. Use MM/DD/YYYY.
4. Copy the formula down.

INT(A2/10000) drops the last four digits and leaves the year, 2026. MOD(INT(A2/100),100) leaves the month, 10. MOD(A2,100) leaves the day, 21.

### If the value is text

A value that is left aligned, or that has leading zeros you must keep, is probably text. Use LEFT, MID and RIGHT to cut it into pieces.

1. Type the text values in **B2:B101**, in cells formatted as Plain text, for example "20261021" in B2.
2. In **E2**, type `=DATE(LEFT(B2,4),MID(B2,5,2),RIGHT(B2,2))`.
3. Format **E2** as MM/DD/YYYY and copy it down.

LEFT(B2,4) is the year, MID(B2,5,2) starts at character 5 and takes two for the month, and RIGHT(B2,2) is the day. In LibreOffice the same formula wrapped in VALUE, as `=DATE(VALUE(LEFT(B2,4)),VALUE(MID(B2,5,2)),VALUE(RIGHT(B2,2)))`, gave the same date.

### If the value is an ISO timestamp

A timestamp such as 2026-10-21T14:35:00 has hyphens, so the month starts at character 6 and the day at character 9. The time starts at character 12.

1. Type the timestamps in **C2:C101**, in Plain text cells, for example "2026-10-21T14:35:00" in C2.
2. In **F2**, type `=DATE(LEFT(C2,4),MID(C2,6,2),MID(C2,9,2))+TIME(MID(C2,12,2),MID(C2,15,2),0)`.
3. Format **F2** as MM/DD/YYYY h:mm AM/PM.

The seconds are dropped. A time zone ending such as Z or -05:00 is ignored, and no time zone conversion is done: the date and time come out as written.

### Flag impossible dates

DATE does not stop on a bad date. It rolls over to the next valid one. The template adds a check column next to each result, **G**, **H** and **I**, that turns the result back into text and compares it with the original. If they differ, the cell says "Check date" and turns red.

The template cell that does the conversion, **D2**, also wraps the formula in blank, text and IFERROR tests, so it is longer than the one above. The core is the same.

### In Excel

The formulas work as typed. Microsoft describes the arguments on its [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) page.

### In Google Sheets

Google documents DATE on its [DATE help page](https://support.google.com/docs/answer/3092969). The same formulas should work as typed, but Sheets was not opened for this post.

## Example

Sample data is fictional. The template has 20 rows, each with one value in all three shapes. Row 2:

| Cell | Typed value | Result |
|---|---|---|
| A2 number | 20261021 | D2 = 10/21/2026 |
| B2 text | "20261021" | E2 = 10/21/2026 |
| C2 ISO text | "2026-10-21T14:35:00" | F2 = 10/21/2026 2:35 PM |

Row 5 holds the leap day 20240229, which gives 02/29/2024. All 60 sample values converted and every check cell says OK. The summary shows 60 values, 60 OK, 0 flagged, and the CHECK cell **L4** says OK.

The **By month** tab counts the converted dates per month for the year in **H1** (2026). For 2026 it counts 19 dates from the numbers, 20 from the text and 20 from the ISO text. One real date falls in another year: the 02/29/2024 leap day.

What happens with impossible dates, in LibreOffice Calc 24.2:

| Typed value | Date you get | Check |
|---|---|---|
| 20261341 (month 13) | 02/10/2027 | Check date |
| 20260230 (Feb 30) | 03/02/2026 | Check date |
| 20230229 (no leap day in 2023) | 03/01/2023 | Check date |
| 20240229 | 02/29/2024 | OK |

## Troubleshooting

### A month 13 or Feb 30 turned into a different date with no error
DATE rolls extra months and days forward instead of failing. Without the check column, 20261341 silently becomes 02/10/2027. Keep the check column, or compare your own result with the original, and fix the source value for any "Check date" row.

### The text version gives the wrong day, such as 10/01/2026
The text has a space at the end, so RIGHT(B2,2) is "1 " instead of "21". In LibreOffice, "20261021 " gave 10/01/2026, and the check flagged it. Clean the text with TRIM first, or remove the space.

### A cell says Check date and the date looks fine
The check compares the result with the original, so any mismatch is flagged. Causes seen in testing: an impossible date, a 6 digit value like 261021 (YYMMDD), text typed in the number column A (put it in column B), a time past 23:59 in ISO text such as T25:10, a date with no time such as 2026-10-21, an ISO value with a space instead of the T, or a space inside the text. A flagged row should not be used, even if it shows a date.

### The time is wrong after a time zone suffix
The suffix is ignored. 2026-10-21T14:35:00Z and 2026-10-21T14:35:00-05:00 both give 2:35 PM. If your export is in UTC and you need local time, add or subtract the hours yourself.

### The result shows a plain number instead of a date
The formula is right but the cell is formatted as a number. Format the cell as MM/DD/YYYY. For dates that flip between month and day when you type or paste them, see the existing post on fixing dates that flip from MM/DD to DD/MM.

## Template

[Download the .xlsx](templates/tidy-tabs-yyyymmdd-to-date.xlsx)

The file has three tabs. **Convert** holds the yellow inputs in **A2:C101** (number, text and ISO text), the three results in **D:F**, the three check columns in **G:I** and a summary in **K1:L4**. **By month** counts the converted dates by month for the year in **H1**. **How to use** has short notes. The Convert tab prints landscape and the other tabs portrait, all on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
