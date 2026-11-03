---
threads_url: https://www.threads.com/@tin_ylab/post/DfSREvSmo5b
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-best-by-date-from-made.html
title: How to Calculate a Best-By Date From the Made-On Date and Shelf Life in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 14 workbooks and 1,750 comparisons with 0 mismatches. Checked: an empty list, 3, 12 and 40 batches, shelf lives from 1 to 36 months, text in a date and a blank months cell, a text or blank as-of date, a window of 0 days, and 01/31 and 02/29 dates. Excel and Google Sheets were not used.
search_description: Calculate a best-by date from the made-on date and shelf life in months with EDATE in Excel and Google Sheets. Free template with status flags.
image_prompts: Three small glass jars of homemade lip balm with blank kraft labels and a pencil lying on a worn wooden shelf beside a roll of twine, soft daylight, no screens, no readable text, no logos
image_alt: Three small jars with blank kraft labels on a wooden shelf beside twine and a pencil
threads: Made it on 03/14, shelf life is 24 months. What is the best-by date?\n=EDATE(B6,C6) adds whole months and keeps the day. 01/31 plus 1 month lands on 02/28, not an error.\nSubtract your as-of date from it and you know how many days are left.
---
Get a best-by date with `=EDATE(B6,C6)`, where B6 is the made-on date and C6 is the shelf life in months. Subtract an as-of date from the result, as in `=D6-$B$2`, to see how many days are left.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. EDATE exists in current Excel and Google Sheets, but neither app was opened for this post.

![Batch list with made-on date, shelf life in months, best-by date, days left and status](images/2026-11-03-team-02-best-by-date-from-made-on-date-shelf-life-edate-excel-google-sheets/best-by-date-edate-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The template layout: **Product** in column A, **Made on** in B, **Shelf life (months)** in C, starting in row 6. The as-of date is in **B2**.

1. In **B2**, type the date to check against, such as 11/02/2026. Type `=TODAY()` instead if you want it to follow the calendar.
2. In **A6:C6**, enter a product, the date it was made and its shelf life in months.
3. In **D6**, type `=EDATE(B6,C6)`. Format the cell as a date (MM/DD/YYYY).
4. In **E6**, type `=D6-$B$2`. Format it as a number. A negative number means the date has passed.
5. In **F6**, type `=IF(E6<0,"Past date",IF(E6<=$B$3,"Use soon","OK"))`, with the use-soon window (days) in **B3**.
6. Fill **D6:F6** down.

EDATE moves a date forward by whole months and keeps the day of the month. If the target month is shorter, it uses that month's last day. In the test, 01/31/2026 plus 1 month gave 02/28/2026, and 02/29/2024 plus 12 months gave 02/28/2025.

A best-by date equal to the as-of date has 0 days left, and the template calls that Use soon, not Past date.

This is date arithmetic on the shelf life you type. The sheet does not say how long any product stays safe or sellable.

Help pages: Microsoft's [EDATE function](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/02/2026; the pages were not read for this post.

## Example

Sample data is fictional: ten batches from a small soap and candle shop, checked on 11/02/2026 with a 30 day window.

| Product | Made on | Shelf life (months) | Best by | Days left | Status |
|---|---|---|---|---|---|
| Lavender soy candle | 03/14/2026 | 24 | 03/14/2028 | 498 | OK |
| Oatmeal goat milk soap | 01/31/2026 | 12 | 01/31/2027 | 90 | OK |
| Bath bomb set | 05/29/2026 | 6 | 11/29/2026 | 27 | Use soon |
| Eucalyptus shower steamer | 09/01/2026 | 2 | 11/01/2026 | -1 | Past date |
| Vanilla room spray | 02/28/2026 | 8 | 10/28/2026 | -5 | Past date |

The full list has 7 rows marked OK, 1 Use soon and 2 Past date.

## Troubleshooting

### The best-by date shows a five-digit number
The cell is formatted as General. Format column D as a date.

### The row says Check
The made-on date or the shelf life is text or blank. A date typed as text, such as `March 14`, is not a date. Retype it as 03/14/2026. A shelf life of `6 months` typed into the cell is text too. Type just 6.

### Every row says Check
The as-of date in **B2** is blank or text. In the test, a text as-of date made every row Check instead of showing wrong day counts.

### The date landed on the 28th, not the 31st
This is how EDATE works for a short month. If you need the same day count instead, add days with `=B6+C6` and put the number of days in column C.

### The days left do not change from day to day
**B2** is a typed date. Replace it with `=TODAY()` to follow the calendar. That version was not run in the tests.

## Template

[Download the .xlsx](templates/tidy-tabs-best-by-dates.xlsx)

The file has two tabs. **Batches** has the as-of date, the use-soon window, room for 40 batches (10 fictional ones filled in), the best-by date, days left and status, plus counts of Past date, Use soon, OK and rows to fix and a **CHECK** cell, **I9**. **How to use** has notes on each formula.

**CHECK** says OK when both settings are numbers and no row needs fixing. Fix every Check before you trust a status.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
