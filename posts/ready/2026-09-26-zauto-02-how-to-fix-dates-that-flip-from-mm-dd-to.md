---
title: How to Fix Dates That Flip From MM/DD to DD/MM After a CSV Import
labels: data-cleanup, data-entry
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small business desk with a laptop showing a spreadsheet full of date columns, a printed invoice with a date circled in pen, screens and labels unreadable
image_alt: Desk with laptop spreadsheet and printed invoice showing a circled date
search_description: Fix dates that switch between MM/DD and DD/MM after importing a CSV into Excel or Google Sheets, with the exact import settings to lock the format.
threads: CSV import turning 03/04/2026 into April 3rd instead of March 4th?\nIt's not random. Excel and Sheets guess the date order from your system settings, not the file.\nFix it by importing the date column as Text first, then converting it with DATEVALUE once you know the real order.
---
Dates flip between MM/DD and DD/MM after a CSV import because Excel and Google Sheets guess the date format from your locale settings instead of reading it from the file. Stop the guessing by importing the date column as **Text** first, then converting it on purpose with a formula once you know which order the file actually uses.

> Works in: Excel for Microsoft 365 (Windows) and Google Sheets, based on Microsoft and Google help pages; not hands-on tested.

## Steps

### In Excel

1. Open a blank workbook first. Don't double-click the CSV file, that skips the import options.
2. Go to **Data > From Text/CSV** (or **Data > Get Data > From File > From Text/CSV**).
3. Select your CSV file and click **Import**.
4. In the preview window, click the date column header, then set its data type to **Text** instead of letting Excel auto-detect.
5. Click **Load**. The dates now sit in the sheet exactly as written in the file, like `03/04/2026`, with no reordering.
6. In a new column, convert them on purpose with `=DATEVALUE(A2)` if the file uses MM/DD/YYYY, or rearrange the text first if it uses DD/MM/YYYY (for example `=DATEVALUE(MID(A2,4,2)&"/"&LEFT(A2,2)&"/"&RIGHT(A2,4))`).
7. Format the new column as a date with **Home > Number Format > Short Date**.

### In Google Sheets

1. Go to **File > Import > Upload** and select the CSV.
2. In the import dialog, choose **Replace spreadsheet** or **Insert new sheet**, then set **Separator type** and click **Import data**. Google Sheets auto-converts anything that looks like a date, which is the problem.
3. Before importing, if you can edit the CSV in a text editor, add a space or apostrophe in front of each date so Sheets treats the whole column as text on import. Remove it after.
4. Once the raw text is safely in a column, use `=DATE(RIGHT(A2,4), LEFT(A2,2), MID(A2,4,2))` for MM/DD/YYYY source data, swapping the LEFT and MID arguments if the source is DD/MM/YYYY.
5. Format the result column with **Format > Number > Date**.

## Example

Sample data below is fictional, made up for this post.

| Raw text from CSV | Actual meaning | Wrong auto-import result | Correct after DATEVALUE/DATE fix |
|---|---|---|---|
| 03/04/2026 | March 4, 2026 | April 3, 2026 | 03/04/2026 |
| 11/22/2026 | November 22, 2026 | Import error or blank | 11/22/2026 |
| 01/09/2026 | January 9, 2026 | September 1, 2026 | 01/09/2026 |

Notice row two: `11/22/2026` has no valid day-first reading, since no month is "22." That's actually a useful clue for figuring out which order your source file uses. If every date in the column could be read either way, you have to check the source system (Etsy export, QuickBooks, a client's invoicing tool) to know for sure.

## Troubleshooting

### All dates shifted by the same pattern (day and month always swapped)
Your system locale (or the exporting app's locale) doesn't match the file. Re-import as text, then convert with the formulas above once you confirm the real order from a date like `11/22/2026` that only works one way.

### Some dates converted, others show as text or errors
Days over 12 can't be read as a month, so Excel or Sheets leaves those cells alone while converting the rest, giving you a mixed column. Import the whole column as text first, then convert every cell with one formula so the result is consistent.

### Dates look right on your screen but wrong after you share the file
The reader's Excel or Sheets locale reformats the date's display, not the file. Keep the underlying value as a real date (not text) so it displays correctly under any locale, or add a separate text column showing the format spelled out, like "04 Mar 2026".

### Leading zeros disappear from dates like 01/09/2026
If the column got auto-detected as a number instead of a date, Excel may drop the leading zero and show `1/9/2026` or convert it to a serial number. Reset the column type during import instead of fixing it after the fact.

### Whole column shows five-digit numbers like 46090
That's a date serial number, not an error. Select the column and apply **Format > Number > Date** (Google Sheets) or **Home > Number Format > Short Date** (Excel) to display it properly.

## Copy-paste setup

Use these formulas once your raw dates are safely imported as text.

**If the source file uses MM/DD/YYYY:**
```
=DATEVALUE(A2)
```

**If the source file uses DD/MM/YYYY, in Excel:**
```
=DATEVALUE(MID(A2,4,2)&"/"&LEFT(A2,2)&"/"&RIGHT(A2,4))
```

**If the source file uses DD/MM/YYYY, in Google Sheets:**
```
=DATE(RIGHT(A2,4), MID(A2,4,2), LEFT(A2,2))
```

Column headers to set up before you import:
- `Raw Date Text` — the imported column, formatted as Text
- `Fixed Date` — the formula column, formatted as a real date

Learn more from the source pages this post is based on:
- [Import or export text (.txt or .csv) files](https://support.microsoft.com/en-us/office/import-or-export-text-txt-or-csv-files-5250ac4c-663c-47ce-937b-1de6f4edb610)
- [DATEVALUE function](https://support.microsoft.com/en-us/office/datevalue-function-df8b07d4-7761-4a93-bc33-b7471bb70a89)
- [Import spreadsheet data in Google Sheets](https://support.google.com/docs/answer/40608)
