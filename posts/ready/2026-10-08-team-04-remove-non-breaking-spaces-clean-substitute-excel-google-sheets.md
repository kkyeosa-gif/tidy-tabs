---
title: How to Remove Non-Breaking Spaces in Excel and Google Sheets
labels: excel-formulas, data-cleaning, text-functions
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same functions; the TRIM, SUBSTITUTE and CLEAN links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: TRIM leaves the non-breaking space (code 160). Use =TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," "))) so lookups and COUNTIF match. Free .xlsx with 8 sample names.
image_prompts: A stack of printed customer order sheets from a small online shop on a packing table, a roll of brown kraft tape, and a pen lying across the top sheet, soft daylight, all paper text blurred and unreadable, no screens, no logos
image_alt: Stack of printed order sheets with a tape roll and pen on a packing table
threads: TRIM did nothing and the lookup still says not found?\nText pasted from a web page or PDF can hold a non-breaking space (code 160), and TRIM only removes code 32.\n=TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," "))) fixes it. =CODE(MID(A2,5,1)) returning 160 confirms it.
---
Use `=TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," ")))`. SUBSTITUTE swaps the non-breaking space (character 160) for a normal space, CLEAN removes line breaks and tabs, and TRIM removes the extra spaces that are left.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Pasted name list with LEN before and after, the character code at position 5, and COUNTIF results before and after cleaning](images/2026-10-08-team-04-remove-non-breaking-spaces-clean-substitute-excel-google-sheets/remove-non-breaking-spaces-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

A non-breaking space looks like a normal space but has character code 160, not 32. Web pages and PDFs often use it, and it comes along when you paste.

Microsoft's [TRIM page](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) says TRIM was designed to trim the 7-bit ASCII space character (value 32), and that by itself it does not remove the nonbreaking space (value 160). In LibreOffice Calc, TRIM left it in place too: a 9-character name stayed at 9 characters.

1. Paste your names into column A, starting in **A2**.
2. In **B2**, type `=LEN(A2)`. A name like Dana Ruiz that shows 9 instead of 8 has an invisible extra character.
3. To confirm the cause, type `=CODE(MID(A2,5,1))` in **C2**. It returns the code of the 5th character. 32 is a normal space. 160 is a non-breaking space. Change the 5 to point at the character you suspect.
4. In **E2**, type `=TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," ")))`.
5. In **F2**, type `=LEN(E2)` to see the length after cleaning.
6. Fill **B2:F2** down for every row.
7. To keep the clean names, copy column E, then use **Paste Special > Values only** and delete the formula columns.

CLEAN deletes a line break or tab. It does not replace it with a space. If one sits between two words, swap it first: `SUBSTITUTE(A2,CHAR(10)," ")`.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUBSTITUTE](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) and [CLEAN](https://support.microsoft.com/en-us/office/clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) pages.

### In Google Sheets

The same formulas should work as typed. Google Sheets was not tested for this post, so check one row with `=LEN` before and after. Google documents the arguments on its [TRIM](https://support.google.com/docs/answer/3094140) and [SUBSTITUTE](https://support.google.com/docs/answer/3094215) pages.

## Example

Sample data below is fictional. It is eight customer names pasted from a web order form and a PDF, checked against a customer list that holds the clean names.

| Pasted name | LEN before | Code at position 5 | Cleaned | LEN after | COUNTIF before | COUNTIF after |
|---|---|---|---|---|---|---|
| Dana Ruiz | 9 | 160 | Dana Ruiz | 9 | 0 | 1 |
| Emma Lee | 8 | 160 | Emma Lee | 8 | 0 | 1 |
| Sara Kim | 9 | 160 | Sara Kim | 8 | 0 | 1 |
| Mark Diaz | 10 | 32 | Mark Diaz | 9 | 0 | 1 |
| Lena Ortiz | 11 | 160 | Lena Ortiz | 10 | 0 | 1 |
| Josh Park | 10 | 160 | Josh Park | 9 | 0 | 1 |
| Anna Cho | 8 | 160 | Anna Cho | 8 | 0 | 1 |
| Mike Stone | 10 | 32 | Mike Stone | 10 | 1 | 1 |

Sara Kim had a non-breaking space at the end and drops from 9 to 8. Lena Ortiz had one plus a trailing tab, and drops from 11 to 10. Josh Park had two in a row, which TRIM collapses to one space, so 10 becomes 9.

Mark Diaz is a different case. Character 5 is a normal space (32), but the cell ends with a line break, so LEN is 10. CLEAN removes it. Mike Stone is the control: a normal name that matches before and after.

The lookup is what breaks in real work. `=COUNTIF($L$2:$L$9,A2)` returns 0 for seven of the eight pasted names, because the text does not equal the clean name in the list. `=COUNTIF($L$2:$L$9,E2)` returns 1 for all eight. `=IFERROR(VLOOKUP(A2,$L$2:$M$9,2,FALSE),"not found")` says not found for Dana Ruiz, but with E2 in place of A2 it returns $120.00.

## Troubleshooting

### LEN is still too high after cleaning
Some other invisible character is left. Use `=CODE(MID(E2,n,1))`, with n set to each position, to find it.

### CHAR(160) does not match, or CODE shows a number other than 160
In LibreOffice Calc, CHAR and CODE returned different numbers in two setups on the same computer (160 in one, 194 in another), apparently depending on locale settings. The template uses `UNICHAR(160)` and `UNICODE` instead, which are tied to the Unicode character itself. Both are in Excel 2013 and later and in Google Sheets, per the [UNICHAR page](https://support.microsoft.com/en-us/office/unichar-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8). Neither was tested in Excel or Google Sheets.

### Two words run together after cleaning
A line break or tab sat between them, and CLEAN deleted it. Replace it with a space first: `=TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A2,CHAR(160)," "),CHAR(10)," ")))`.

### COUNTIF or VLOOKUP still returns 0 or not found
The lookup list may have the same problem, not just your pasted column. Run the same cleanup on both sides. Also check for a trailing space in the list, and for a name spelled differently.

## Template

[Download the .xlsx](templates/tidy-tabs-remove-non-breaking-spaces.xlsx)

The file has two tabs. **Pasted names** holds the eight sample names, the LEN, character code, TRIM-only and cleaned columns, the COUNTIF and VLOOKUP checks against a small customer list in columns L and M, and a count of names found before and after. **How to use** has short fill-in notes.

The names and prices are fictional. You can paste your own names over column A. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
