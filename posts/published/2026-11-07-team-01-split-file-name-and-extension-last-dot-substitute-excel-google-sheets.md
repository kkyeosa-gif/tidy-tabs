---
threads_url: https://www.threads.com/@tin_ylab/post/DfckPbVm5JZ
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-split-file-name-into-name-and.html
title: How to Split a File Name Into Name and Extension in Excel and Google Sheets
labels: excel-formulas, text-formulas, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 11 workbooks and 6,754 comparisons with 0 mismatches. Checked: the sample, no names, a.b, a, a., a.., .., a lone dot, .gitignore, .tar.gz, spaces around and doubled inside a name, capital and lowercase extensions, a space after the last dot, a plain number typed as a name (123), accents, a blank gap, whitespace-only cells, a 300-character name, 100 rows, the count table with and without a leading dot, and six random sets of up to 100 names. Excel and Google Sheets were not used. Not run: non-breaking spaces, names with a folder path, names containing CHAR(1), a decimal such as 12.5 typed as a name.
search_description: Split a file name into name and extension in Excel and Google Sheets by finding the last dot with SUBSTITUTE, FIND, MID and LEFT. Free template.
image_prompts: A small stack of blank manila folders with tabs, a brass label holder and a short pencil on a gray metal filing cabinet top, soft side light, realistic photo, no screens, no readable text, no logos
image_alt: Blank manila folders with tabs beside a brass label holder and a pencil on a cabinet
threads: Invoice.2026.11.docx has three dots, and only the last one starts the extension.\nSUBSTITUTE swaps that dot for CHAR(1), FIND locates it, MID takes docx and LEFT takes Invoice.2026.11.\nThe free template flags names like "Dr. Smith flyer" instead of guessing, and counts files by extension.
---
Find the last dot with `=FIND(CHAR(1),SUBSTITUTE(B7,".",CHAR(1),LEN(B7)-LEN(SUBSTITUTE(B7,".",""))))`, then take the extension with `=MID(B7,C7+1,500)` and the name with `=LEFT(B7,C7-1)`. For "Invoice.2026.11.docx" the last dot is at position 16, the extension is docx and the name is Invoice.2026.11.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model (11 workbooks, 6,754 comparisons). SUBSTITUTE, FIND, MID, LEFT, LEN, TRIM and COUNTIF exist in Microsoft 365, Excel and Google Sheets, but neither Excel nor Google Sheets was opened for this post.

![Table splitting twelve fictional file names into name, extension and status with a count by extension](images/2026-11-07-team-01-split-file-name-and-extension-last-dot-substitute-excel-google-sheets/split-file-name-extension-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. On the **Files** tab, type or paste one file name per row in column **A**, starting at row 5. Paste names only, not folder paths: a dot in a folder name would be taken as the last dot (not tested).

1. In **B5**, type `=IF(A5="","",TRIM(A5))` to remove extra spaces. TRIM removes ordinary spaces only; a non-breaking space copied from a web page stays (not tested).
2. In **C5**, count the dots and find the last one: `=IF(B5="","",IF((LEN(B5)-LEN(SUBSTITUTE(B5,".","")))=0,0,FIND(CHAR(1),SUBSTITUTE(B5,".",CHAR(1),(LEN(B5)-LEN(SUBSTITUTE(B5,".","")))))))`. A result of 0 means the name has no dot.
3. In **D5**, type `=IF(B5="","",IF(OR(C5=0,C5=1,C5=LEN(B5)),"",MID(B5,C5+1,500)))` for the extension. A name with no dot, a leading dot or a final dot gets an empty extension. 500 characters after the dot is the cap.
4. In **E5**, type `=IF(B5="","",IF(OR(C5=0,AND(C5=1,LEN(B5)>1)),B5,LEFT(B5,C5-1)))` for the name without the extension.
5. In **F5**, type `=IF(B5="","",IF(C5=0,"No extension",IF(C5=LEN(B5),"Ends with dot",IF(C5=1,"Starts with dot",IF(ISNUMBER(FIND(" ",D5)),"Check extension","OK")))))`.
6. In **G5**, type `=IF(B5="","",LOWER(D5))` so PNG and png count together.
7. Fill **B5:G5** down to row 104.
8. To count files by type, type extensions in **I13:I20** (with or without a leading dot) and use `=IF(SUBSTITUTE(TRIM(I13),".","")="","",COUNTIF($G$5:$G$104,SUBSTITUTE(TRIM(I13),".","")))` in **J13**.

How the last dot works: `LEN(B7)-LEN(SUBSTITUTE(B7,".",""))` is the number of dots. SUBSTITUTE with a fourth argument replaces only that numbered dot, so using the dot count replaces the last one with CHAR(1), an invisible character. FIND then returns where it sits. A name that already contains CHAR(1) would break the position (not tested).

The extension is whatever follows the last dot. That is why "Dr. Smith flyer" is flagged **Check extension** instead of being split into "Dr" and " Smith flyer".

Help pages: Microsoft's [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) and [FIND function](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 12 file names in rows 5 to 16.

| File name | Last dot at position | Extension | Name without extension | Status |
|---|---|---|---|---|
| Mug-Wrap-Design_v2.png | 19 | png | Mug-Wrap-Design_v2 | OK |
| Invoice.2026.11.docx | 16 | docx | Invoice.2026.11 | OK |
| wedding-sign.final.SVG | 19 | SVG | wedding-sign.final | OK |
| Logo.tar.gz | 9 | gz | Logo.tar | OK |
| Packing list | 0 | (blank) | Packing list | No extension |
| Dr. Smith flyer | 3 | " Smith flyer" | Dr | Check extension |

The summary in **J4:J9** shows 12 names listed, 10 OK, 1 No extension, 0 starting or ending with a dot, 1 Check extension and **CHECK** OK. The count table shows png 2, pdf 2, svg 1, jpg 1 and zip 1. The lowercase column is why "wedding-sign.final.SVG" counts as svg.

## Troubleshooting

### The extension is blank
The name has no dot, starts with a dot (such as .gitignore) or ends with one (such as report.). **F** says which. Fix the name or leave it.

### The extension has a space in it
The text after the last dot is a sentence, as in "Dr. Smith flyer". **F** shows Check extension. The formula does not guess, so rename the file or ignore the row.

### Logo.tar.gz gives gz, not tar.gz
By design: the extension is what follows the last dot. The name column keeps Logo.tar.

### A count in the table is too low
The count table compares lowercase extensions typed without spaces. A leading dot in **I13** is fine. Check that the extension is spelled the same way as in column **D**.

### A number typed as a name shows no extension
A plain number such as 123 is treated as text with no dot. A decimal such as 12.5 was not tested.

## Template

[Download the .xlsx](templates/tidy-tabs-split-file-name-extension.xlsx)

The file has two tabs. **Files** has 100 rows from row 5 with 12 fictional sample names and yellow input cells in column **A**, gray formula cells in **B:G**, a summary in **J4:J9** and the extension count table in **I12:J20**. **How to use** has notes on each formula.

The template holds 100 names and no more. Excel and Google Sheets were not used. In Google Sheets, go to **File > Import > Upload**, select the file, and check the results against the Example table before relying on them.
