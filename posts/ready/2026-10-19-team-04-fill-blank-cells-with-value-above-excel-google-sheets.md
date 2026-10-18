---
title: How to Fill Blank Cells With the Value Above in Excel and Google Sheets
labels: excel-formulas, small-business, data-cleanup
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 15 sample rows were checked against a Python forward fill and Decimal sums, plus 9 edited variants (a blank first row, a cell holding only spaces, padded names, numbers as the key, a blank separator row, rows with no vendor after empty rows, an amount typed as text, an empty report). Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Fill blank cells with the value above using =IF(Report!A3="",A2,Report!A3) in Excel and Google Sheets. Free template with totals by vendor.
image_prompts: Printed vendor invoice sheet with every other row left empty, a ruler laid under one gap and a highlighter beside a stack of sample candle boxes, soft natural window light, no screens, no readable text, no logos
image_alt: Printed sheet with empty gaps, a ruler and highlighter beside small candle boxes
threads: Exported report with the vendor name only on the first row of each group?\n=IF(Report!A3="",A2,Report!A3) copied down fills every blank with the value above, no macro. A TRIM version handles blanks holding a stray space.
---
Fill each blank with the cell above it using `=IF(Report!A3="",A2,Report!A3)`, copied down the Filled sheet under a first row of `=Report!A2`. If some blanks hold a stray space, wrap the test in TRIM: `=IF(TRIM(Report!A3)="",A2,Report!A3)`.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Vendor report with blank vendor cells beside the same rows filled down](images/2026-10-19-team-04-fill-blank-cells-with-value-above-excel-google-sheets/fill-blanks-vendor-report-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

Many exports type the vendor on the first row of each group only. That breaks sorting, filters and SUMIFS. The fix lives on a second tab, so your report stays untouched.

1. Make a tab named **Report**. Put Vendor in **A1**, Product in **B1** and Amount in **C1**, then paste your data from row 2 down.
2. Make a second tab named **Filled**. In **A1**, type the header Vendor (fill down).
3. In **A2**, type `=IF(AND(Report!A2="",Report!B2="",Report!C2=""),"",IF(Report!A2="","",Report!A2))`. The first row has no row above it, so a bare `=Report!A2` would show 0 when the cell is empty. The `IF` returns a blank instead.
4. In **A3**, type `=IF(AND(Report!A3="",Report!B3="",Report!C3=""),"",IF(Report!A3="",A2,Report!A3))`. The part that does the work is `IF(Report!A3="",A2,Report!A3)`: if the Report cell is blank, take the Filled cell above, otherwise keep the Report value.
5. In **B1**, type Vendor (fill down, spaces too). In **B2**, type `=IF(AND(Report!A2="",Report!B2="",Report!C2=""),"",TRIM(Report!A2))`.
6. In **B3**, type `=IF(AND(Report!A3="",Report!B3="",Report!C3=""),"",IF(TRIM(Report!A3)="",B2,TRIM(Report!A3)))`. This one treats a cell holding only spaces as blank and trims padded names.
7. In **C2**, type `=IF(Report!B2="","",Report!B2)`. In **D2**, type `=IF(Report!C2="","",Report!C2)`. These copy Product and Amount across.
8. Select **A3:B3**, then **C2:D2**, and drag each fill handle down to row 60.
9. To total by vendor, type vendor names in **F2:F8**. In **H2**, type `=IF(F2="","",SUMIFS($D$2:$D$60,$A$2:$A$60,F2))` and fill down.

The `AND` test around everything means a row that is empty on the Report tab stays empty here. Without it, the last vendor would repeat down to row 60.

### In Excel

The formulas work as typed. Microsoft describes the functions on its [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) and [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) pages.

### In Google Sheets

The same formulas should work as typed. Drag the small square at the corner of the selected cell to fill down. Google describes [TRIM](https://support.google.com/docs/answer/3094140) in its help. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small market vendor report where the vendor name appears only on the first row of each group.

| Vendor on Report | Product | Amount | Vendor on Filled |
|---|---|---|---|
| Pine Hollow Candles | Cedar candle | $54.00 | Pine Hollow Candles |
| (blank) | Balsam candle | $48.00 | Pine Hollow Candles |
| (blank) | Wick refill pack | $18.50 | Pine Hollow Candles |
| Salt Marsh Soap | Sea salt bar | $30.00 | Salt Marsh Soap |
| (blank) | Oat bar | $27.50 | Salt Marsh Soap |
| (blank) | Gift set | $66.00 | Salt Marsh Soap |
| (blank) | Soap dish | $21.00 | Salt Marsh Soap |
| Willow & Wick | Birch candle | $40.00 | Willow & Wick |
| (blank) | Match box | $12.00 | Willow & Wick |

The full sample has 15 rows and five vendors. A SUMIFS on the Report tab finds only the amount beside each name: Pine Hollow Candles shows $54.00 and Salt Marsh Soap shows $30.00. On the Filled tab they total $120.50 and $144.50. All five vendors add up to $632.00 both before and after.

**I15** shows OK when no used row is still blank and the two amount totals match.

## Troubleshooting

### The first vendor is still blank, and I15 says CHECK
The first data row on the Report tab has no vendor, so there is nothing above it to copy. Type the vendor in **Report!A2**. With that cell cleared, Filled A2 to A4 stayed blank and Pine Hollow Candles totaled $0.00.

### A vendor total is $0.00 but the name is there
A cell may hold a single space, which `Report!A3=""` does not treat as blank. Use column **B**, the TRIM version. In a test with one space where Salt Marsh Soap starts, column A kept the space and that vendor totaled $0.00, while column B filled from the row above.

### SUMIFS misses a name with extra spaces
A padded name such as " Salt Marsh Soap " does not equal "Salt Marsh Soap". The plain column keeps the padding, so a SUMIFS on the clean name misses it. The TRIM column returns the trimmed text, so total from column **I**, not **H**.

### Vendors after a blank row are still blank
A fully blank row ends a group, and the formula returns blank for it. The next row then has a blank vendor above it, so I11 and I12 count it and I15 says CHECK. Type the vendor name on the row after the break.

## Limits

Formulas only, no macros. The Filled tab reads Report rows 2 to 60.

## Template

[Download the .xlsx](templates/tidy-tabs-fill-blanks-above.xlsx)

The file has three tabs. **Report** holds the 15 sample rows in yellow input cells **A2:C60**. **Filled** has the fill down columns **A** and **B**, the Product and Amount columns **C** and **D**, the vendor totals in **F:I** and the checks in **I10:I15**. **How to use** has short notes. All tabs print portrait on US Letter, with the first 40 rows set as the print area.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
