---
threads_url: https://www.threads.com/@tin_ylab/post/DdtwtpYm3gC
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-keep-leading-zeros-in-excel-zip.html
title: How to Keep Leading Zeros in Excel ZIP Codes
labels: excel-basics, data-entry
tested_in: LibreOffice Calc 24.2.7 (Linux), 09/25/2026; Excel and Google Sheets steps from Microsoft/Google help pages
image_prompts: A stack of outgoing US mail envelopes with printed address labels on a home-office desk beside a laptop with a spreadsheet open, labels and screen out of focus
image_alt: Addressed envelopes next to a laptop with a spreadsheet
search_description: Excel dropping the 0 from ZIP codes like 02108? Format the column as Text before typing, or rebuild it with =TEXT(A2,"00000").
threads: Excel turned 02108 into 2108 again?\nFormat the column as Text before you type or paste, not after.\nAlready broken? =TEXT(A2,"00000") brings the zero back.
---
Keep ZIP codes intact by formatting the cells as text before you type or paste them: **Text** in Excel, **Plain text** in Google Sheets. This preserves leading zeros in codes such as 02108.

> Works in: Excel for Microsoft 365 on Windows, Google Sheets in a desktop browser. Tested in: LibreOffice Calc 24.2.7 only. The Excel and Google Sheets steps come from Microsoft and Google help pages.

![ZIP codes stored as numbers next to the same codes repaired with the TEXT formula](images/2026-09-28-a-keep-leading-zeros-zip-codes-excel/zip-codes-leading-zeros-fixed.png) *ZIPs typed as numbers (column C) and repaired with =TEXT() (columns D and E). PDF export from LibreOffice Calc 24.2.*

## Steps

### In Excel

1. Select the ZIP code column by clicking its column letter.
2. Press **Ctrl+1** to open **Format Cells**.
3. On the **Number** tab, select **Text** and click **OK**.
4. Type or paste your ZIP codes. 02108 now stays 02108.
5. For a single cell, type an apostrophe first: `'02108`. Excel stores the entry as text and the apostrophe does not print.

Microsoft covers these methods in [Keeping leading zeros and large numbers](https://support.microsoft.com/en-us/excel/keeping-leading-zeros-and-large-numbers).

**Stop Excel from removing zeros in the first place.** Excel for Microsoft 365 has an **Automatic Data Conversion** setting. Go to **File > Options > Data**, find **Automatic Data Conversion**, and clear **Remove leading zeros and convert to number**. Microsoft says this applies when you type, paste, or open a .csv or .txt file ([Set automatic data conversions](https://support.microsoft.com/en-us/excel/set-automatic-data-conversions)).

**Import a CSV without losing zeros.** Double-clicking a .csv file lets Excel guess the column types. Import it instead:

1. Click **Data > From Text/CSV** and pick the file.
2. In the preview window, click **Transform Data** to open the Power Query Editor. Older versions label this button **Edit**.
3. Click the ZIP column header. On the **Home** tab, in the **Transform** group, set **Data Type** to **Text**.
4. If asked, click **Replace Current**.
5. Click **Close & Load**.

### In Google Sheets

1. Select the ZIP code column.
2. Click **Format > Number** and choose **Plain text**.
3. Type or paste your ZIP codes.
4. For a single cell, type an apostrophe first: `'02108`.

**Import a CSV.** Click **File > Import**, upload the file, and pick an import location ([Import data sets & spreadsheets](https://support.google.com/docs/answer/40608)). If the import dialog shows an option to convert text to numbers, dates, and formulas, uncheck it before you click **Import**. Google's help page does not document that option, so if you don't see it, repair the column with the formula below.

### Repair ZIP codes that already lost their zeros

Formatting a column as Text after the fact does not bring zeros back. The cell still holds 2108. Rebuild it with a formula:

1. In an empty column next to your ZIPs, type `=TEXT(A2,"00000")`, using the cell that holds the damaged ZIP. 2108 becomes 02108.
2. Fill the formula down the column.
3. Copy the new column and paste it back as values. Excel: **Home > Paste > Values**. Google Sheets: **Edit > Paste special > Values only**.
4. Delete the old column and the helper formulas.

For ZIP+4, use `=TEXT(A2,"00000-0000")`. This works only if the cell holds all 9 digits as one number, such as 21081522, which becomes 02108-1522.

### Why a custom 00000 format is risky

You can also select the column, open **Format Cells > Number > Custom**, and type `00000`. The cell then shows 02108, but the value is still the number 2108. Mail merge, CSV export, and lookups against a text ZIP list can read the underlying 2108 and drop the zero. Microsoft suggests this format only when the data stays inside the workbook. For anything that leaves the file, store ZIPs as text.

## Example

Here is what happens to four real ZIP codes when they land in a General (number) cell, and what they look like after the fix:

| City | State | ZIP typed | Stored as number | After fix |
| --- | --- | --- | --- | --- |
| Boston | MA | 02108 | 2108 | 02108 |
| Hoboken | NJ | 07030 | 7030 | 07030 |
| Holtsville | NY | 00501 | 501 | 00501 |
| Burlington | VT | 05401 | 5401 | 05401 |

We checked this in LibreOffice Calc 24.2. Opening the sample CSV turned 02108 into 2108 and 00501 into 501. The hyphenated ZIP+4 values, such as 02108-1522, stayed as text because the hyphen stops them from reading as numbers. `=TEXT(C2,"00000")` brought 2108 back to 02108 and 501 back to 00501.

## Troubleshooting

### The zeros disappear again after I paste

A normal paste can carry the source cell's format and replace your Text format. Paste values only, or reapply Text to the column and paste again. In Excel for Microsoft 365, turning off **Remove leading zeros and convert to number** also covers pastes.

### Excel shows a green triangle on every ZIP

That is Excel flagging a number stored as text. For ZIP codes, that's what you want. Leave the triangle or ignore the error; do not choose **Convert to Number**, or the zeros go away.

### The sheet shows 02108 but my mail merge prints 2108

The column most likely uses the custom `00000` format, which changes only the display. Rebuild the column with `=TEXT(A2,"00000")` and paste it back as values.

### My ZIP+4 came out as 00000-2108

The formula `=TEXT(A2,"00000-0000")` ran on a 5-digit ZIP. Use the ZIP+4 pattern only on cells holding all 9 digits. If your data has only 5 digits, use `"00000"`.

## Template

[Download the ZIP code practice workbook (.xlsx)](templates/tidy-tabs-zip-codes-leading-zeros.xlsx)

The workbook has three tabs:

- **Fixed (text)**: six ZIP codes and ZIP+4 codes in columns formatted as Text before entry, so the zeros stay.
- **Broken (numbers)**: the same ZIPs stored as numbers, with `=TEXT()` repair formulas in columns D (5-digit) and E (ZIP+4).
- **How to use**: short notes on each tab.

To practice an import, grab [the sample CSV](templates/tidy-tabs-zip-codes-sample.csv) with the same six cities.

The sample data is fictional practice data: the city names and 5-digit ZIP codes are real, but the +4 extensions are made up.

To use the workbook in Google Sheets, open a blank sheet and click **File > Import > Upload**, then choose the .xlsx file.
