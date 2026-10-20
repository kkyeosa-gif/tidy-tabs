---
threads_url: https://www.threads.com/@tin_ylab/post/DeuHD9ElCc6
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-make-mailing-labels-from-address.html
title: How to Make Mailing Labels From an Address List in Excel and Google Sheets
labels: excel-formulas, small-business, mailing
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and all 18 label cells were compared with a Python layout of the addresses in reading order, plus a list of 12 addresses, a blank address row, a 19th to 21st address, an empty list, a blank name and a ZIP typed as a number. The Labels tab exported to one US Letter portrait PDF page. Nothing was printed on paper or label sheets, whether a very long street name wraps to a fourth line was not inspected, and Excel and Google Sheets were not opened.
search_description: Make mailing labels from an address list with INDEX, ROW and COLUMN in Excel and Google Sheets, no mail merge. Free 3 across label template for US Letter.
image_prompts: Sheet of blank self-adhesive address labels partly peeled, a stack of plain white envelopes and a roll of stamps on a home office table, soft natural window light, no screens, no readable text, no logos
image_alt: Partly peeled blank address label sheet with envelopes and a stamp roll
threads: 18 addresses in one column, 3 across a label sheet.\nOne INDEX formula with (ROW()-1)*3+COLUMN() lays them out, and CHAR(10) puts name, street and city on separate lines. Not tested on real label sheets.
---
Lay one column of addresses out as a 3 across label sheet with a single INDEX formula copied across the grid, such as `=IFERROR(INDEX(Addresses!$A$2:$A$19,(ROW()-1)*3+COLUMN()),"")` for the name line. The full label formula joins name, street, and city, state and ZIP with `CHAR(10)` line breaks.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post, and nothing was printed on paper or on label sheets.

![Three across, six down grid of eighteen mailing labels with names and sample addresses](images/2026-10-19-team-05-mailing-labels-3-across-from-address-list-index-row-column-excel-google-sheets/mailing-labels-3-across-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

You need two tabs: **Addresses** holds one person per row, and **Labels** holds the grid. This is layout only, not a brand label template.

1. On a tab named **Addresses**, type the headers Name, Street, City, State and ZIP in **A1:E1**. Type one person per row in rows 2 to 19.
2. Format column **E** as Plain text before you type or paste ZIP codes, so 02108 keeps its leading zero.
3. Add a tab named **Labels**. Make columns **A:C** about 32 characters wide and rows 1 to 6 about 100 points tall, which fits one US Letter page.
4. Select **A1:C6** and turn on Wrap Text. The line breaks from `CHAR(10)` only show when wrapping is on.
5. Work out which address each cell needs. `(ROW()-1)*3+COLUMN()` gives 1 in **A1**, 2 in **B1**, 3 in **C1** and 4 in **A2**. It counts across each row, then moves down. The grid starts at **A1**, so there is no offset to subtract.
6. In **A1**, type this formula as one line:

   `=IFERROR(IF(INDEX(Addresses!$A$2:$A$19,((ROW()-1)*3+COLUMN()))&INDEX(Addresses!$B$2:$B$19,((ROW()-1)*3+COLUMN()))&INDEX(Addresses!$C$2:$C$19,((ROW()-1)*3+COLUMN()))="","",INDEX(Addresses!$A$2:$A$19,((ROW()-1)*3+COLUMN()))&CHAR(10)&INDEX(Addresses!$B$2:$B$19,((ROW()-1)*3+COLUMN()))&CHAR(10)&INDEX(Addresses!$C$2:$C$19,((ROW()-1)*3+COLUMN()))&", "&INDEX(Addresses!$D$2:$D$19,((ROW()-1)*3+COLUMN()))&" "&INDEX(Addresses!$E$2:$E$19,((ROW()-1)*3+COLUMN()))),"")`

7. Copy **A1**, select **A1:C6** and paste. Every cell now holds the same formula, and ROW and COLUMN make each one pick a different address.
8. Set the page to US Letter, portrait, with 0.5 in margins and centered horizontally. Set the print area to **A1:C6**.

The `$` signs lock each INDEX range to rows 2 to 19. The IF test returns a blank label when name, street and city are all empty.

### Check that every address got a label

Put these outside the print area, in **F2:F5** on the Labels tab:

- **F2**: `=SUMPRODUCT(--(Addresses!A2:A19&Addresses!B2:B19&Addresses!C2:C19<>""))` counts the addresses listed.
- **F3**: `=COUNTIF(A1:C6,"?*")` counts the label cells that are filled.
- **F4**: the same count as F2 for Addresses rows 20 to 200.
- **F5**: `=IF(AND(F2=F3,F4=0),"OK","CHECK")`.

### In Excel

The formulas work as typed. Microsoft describes the function on its [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) page.

### In Google Sheets

The same formulas should work as typed. Google describes [INDEX](https://support.google.com/docs/answer/3098242) in its help. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional. These are the first six of the 18 customers, in the order they land on the page.

| Address # | Label cell | Line 1 | Line 2 | Line 3 |
|---|---|---|---|---|
| 1 | A1 | Maya Alvarez | 14 Beacon Row | Boston, MA 02108 |
| 2 | B1 | Tom Brennan | 220 Willow Ave, Apt 3 | Hoboken, NJ 07030 |
| 3 | C1 | Priya Shah | 9 Orchard Lane | Portland, ME 04101 |
| 4 | A2 | Jordan Lee | 1800 Maple Street | Austin, TX 78701 |
| 5 | B2 | Dana Whitfield | 75 Harbor View Road | Providence, RI 02903 |
| 6 | C2 | Carlos Ortega | 412 Elm Court | Tucson, AZ 85701 |

The last label, **C6**, is address 18: Victor Reyes, 905 Canal Street, New Orleans, LA 70130. With 18 addresses, **F2** and **F3** both show 18, **F4** shows 0 and **F5** shows OK. All 18 cells matched the Python layout.

Print on plain paper first and hold it over the label sheet against a window. This grid is not an Avery or other brand template. Adjust column widths and row heights on the Labels tab to fit.

## Troubleshooting

### F5 says CHECK and the last names are missing
You have more than 18 addresses. Only Addresses rows 2 to 19 reach the grid. With addresses in rows 20 to 22, **F2** and **F3** stayed at 18, **F4** showed 3 and **F5** showed CHECK. Print the first 18, then move the next 18 up into rows 2 to 19.

### Each label is one long line
Wrap Text is off, so the `CHAR(10)` line breaks do not show. Turn on wrapping for **A1:C6**.

### The ZIP lost its leading zero
A ZIP typed as the number 7030 stays 7030. Cells formatted as text keep what you type, so format the ZIP column as Plain text first, then retype the codes that already lost the zero.

### The labels shifted or turned into errors after I inserted a row or column
The formula uses `ROW()` and `COLUMN()`, so inserting rows above the grid or columns to its left changes the numbers. Copy a good label formula back over **A1:C6** to repair it.

## Template

[Download the .xlsx](templates/tidy-tabs-mailing-labels.xlsx)

The file has three tabs. **Addresses** holds the headers and 18 sample customers in the yellow cells **A2:E40**, with the ZIP column formatted as text. **Labels** holds the 3 by 6 grid in **A1:C6** and the counts and check in **F2:F5**. **How to use** has short notes. The Labels tab prints on one US Letter portrait page.

The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets, and not printed on label paper. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
