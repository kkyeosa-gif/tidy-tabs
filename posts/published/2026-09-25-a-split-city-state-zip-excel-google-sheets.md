---
threads_url: https://www.threads.com/@tin_ylab/post/Ddt8ap6GlDW
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-split-city-state-and-zip-into.html
title: How to Split City, State, and ZIP Into Separate Columns in Excel and Google Sheets
labels: excel-basics, data-cleanup
tested_in: LibreOffice Calc 24.2.7 (Linux), 09/25/2026 (Example results and every Troubleshooting case). Excel and Google Sheets steps from Microsoft/Google help pages.
image_prompts: A small business owner's wooden desk with a stack of printed mailing labels and addressed envelopes next to an open laptop; the laptop shows a spreadsheet with three narrow columns, slightly out of focus | Close-up of hands sorting addressed envelopes into three neat piles on a desk, the address text blurred
image_alt: Desk with mailing labels, envelopes, and a laptop showing a spreadsheet | Hands sorting addressed envelopes into three piles
search_description: Split "Boston, MA 02108" into City, State, and ZIP with three formulas that work in Excel and Google Sheets and keep the ZIP's leading zero.
threads: Got a column of "Boston, MA 02108" and need City, State, ZIP in separate columns?\nThree formulas, same in Excel and Google Sheets:\nCity =LEFT(B2,FIND(",",B2)-1)\nState =MID(B2,FIND(",",B2)+2,2)\nZIP =RIGHT(B2,5)\nRIGHT returns text, so 02108 keeps its 0.
---
To split "Boston, MA 02108" into three columns, use `=LEFT(B2,FIND(",",B2)-1)` for the city, `=MID(B2,FIND(",",B2)+2,2)` for the state, and `=RIGHT(B2,5)` for the ZIP. These formulas work the same in Excel and Google Sheets, and the ZIP comes back as text, so 02108 keeps its leading zero.

> Works in: Excel for Microsoft 365 and Google Sheets (standard LEFT, MID, RIGHT, FIND functions). Tested in: LibreOffice Calc 24.2.7 only, with the eight addresses in the Example below.

## Steps

### Split with formulas (Excel or Google Sheets)
1. Put your addresses in column B, one per row, in the form `City, ST 12345`.
2. Type the headers **City**, **State**, and **ZIP** in C1, D1, and E1.
3. In C2, type `=LEFT(B2,FIND(",",B2)-1)`. This keeps everything before the comma.
4. In D2, type `=MID(B2,FIND(",",B2)+2,2)`. This skips the comma and the space, then takes 2 letters.
5. In E2, type `=RIGHT(B2,5)`. This takes the last 5 characters.
6. Select C2:E2 and drag the fill handle down to the last address.
7. To keep only the results, copy C:E and paste them back as values. In Excel, use **Home > Paste > Values**. In Google Sheets, use **Edit > Paste special > Values only**.

Microsoft walks through the same LEFT/MID/RIGHT approach on its [split text with functions](https://support.microsoft.com/en-us/office/split-text-into-different-columns-with-functions-49ec57f9-3d5a-44b2-82da-50dded6e4a68) page.

### Why not Text to Columns?
Excel's **Data > Text to Columns** wizard splits on delimiters you pick, such as **Comma** and **Space** ([Microsoft's steps](https://support.microsoft.com/en-us/office/split-text-into-different-columns-with-the-convert-text-to-columns-wizard-30b14928-5550-41f5-97ca-7a3e9c363ed7)). That works for "Boston, MA 02108". It breaks on "Salt Lake City, UT 84101", because every space becomes a split, and the city lands in three columns.

Google Sheets' [SPLIT function](https://support.google.com/docs/answer/3094136) has the same catch. By default it splits on each character in the delimiter, so `=SPLIT(B2,", ")` also cuts multi-word cities apart.

The formulas above only look at the comma, so a city can have as many words as it wants.

## Example

Eight addresses, including two multi-word cities and a ZIP that starts with two zeros. We ran the formulas in LibreOffice Calc 24.2 and got these results:

| City, ST ZIP | City | State | ZIP |
| --- | --- | --- | --- |
| Boston, MA 02108 | Boston | MA | 02108 |
| Hoboken, NJ 07030 | Hoboken | NJ | 07030 |
| Holtsville, NY 00501 | Holtsville | NY | 00501 |
| Burlington, VT 05401 | Burlington | VT | 05401 |
| Chicago, IL 60601 | Chicago | IL | 60601 |
| Las Vegas, NV 89101 | Las Vegas | NV | 89101 |
| Salt Lake City, UT 84101 | Salt Lake City | UT | 84101 |
| St. Louis, MO 63101 | St. Louis | MO | 63101 |

All eight ZIPs came back as text, including 00501. That matters if you later export to CSV or run a mail merge. If your ZIPs have already lost their zeros somewhere else, see [How to Keep Leading Zeros in Excel ZIP Codes](https://tidytabs.blogspot.com/2026/09/how-to-keep-leading-zeros-in-excel-zip.html).

![Split tab with addresses in column B and City, State, ZIP filled in by formula](images/2026-09-25-a-split-city-state-zip-excel-google-sheets/split-tab-libreoffice.png) *The Split tab with the sample addresses. PDF export from LibreOffice Calc 24.2.*

## Troubleshooting

### City and State show #VALUE!
The address has no comma ("Boston MA 02108"), so FIND can't find one. ZIP still works. Add the comma or fix that row by hand.

### State shows one letter, like "A "
There's no space after the comma ("Boston,MA 02108"). Change the `+2` in the State formula to `+1`, or add the space back.

### ZIP lost its first digit, like "2108 "
The address ends with a trailing space, so RIGHT picks up the space. Wrap the address in TRIM: `=RIGHT(TRIM(B2),5)`.

### ZIP shows "-1522"
That row has a ZIP+4 ("02108-1522"), which is 10 characters. Use `=RIGHT(B2,10)` for the full ZIP+4, or `=LEFT(RIGHT(B2,10),5)` for just the 5-digit ZIP.

### The results disappear when I delete column B
They're formulas that read column B. Paste them as values first (step 7), then delete the original column.

## Template

[Download the City, State, ZIP splitter (.xlsx)](templates/tidy-tabs-split-city-state-zip.xlsx)

The workbook has two tabs. **Split** holds eight sample addresses in column B, with the City, State, and ZIP formulas already filled in (gray cells). **How to use** explains each formula. Paste your own addresses over column B and copy the gray cells down for more rows. Customer names are fictional; the city, state, and ZIP combinations are real.

To open it in Google Sheets, click **File > Import** and upload the .xlsx.
