---
threads_url: https://www.threads.com/@tin_ylab/post/DfVDmPtFgJ0
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-make-initials-from-full-name-in.html
title: How to Make Initials From a Full Name in Excel and Google Sheets
labels: excel-formulas, small-business, text
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 10 workbooks and 4,040 comparisons with 0 mismatches. Checked: an empty list, a one-word name, extra spaces, lower case, a name of five words, a hyphenated last name, an apostrophe, a suffix of Jr., a name typed Last, First, a name starting with a digit, a name with accented letters, and five random lists of up to 100 names. Excel and Google Sheets were not used.
search_description: Make initials from a full name in Excel and Google Sheets with UPPER, LEFT and SUBSTITUTE. Free template for monogram and engraving order lists.
image_prompts: A blank leather keychain tag and a small metal stamping hammer beside a bare wooden coaster on a workbench, a pencil and a few wood shavings nearby, soft daylight, no screens, no readable text, no logos
image_alt: Blank leather tag and small hammer beside a bare wooden coaster on a workbench
threads: Engraving 40 orders and need two-letter initials from the customer names?\nFirst letter: =UPPER(LEFT(TRIM(A5),1)). Last name's letter comes from the last word, found with SUBSTITUTE, REPT and RIGHT.\nMaria Alvarez gives MA. Dana Marie Thompson gives DT. A one-word name shows Check.
---
Get the first initial with `=UPPER(LEFT(TRIM(A5),1))` and the last initial from the last word of the name, then join them with `&`. Maria Alvarez becomes MA. The last-word formula is in step 2 below.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. UPPER, LEFT, RIGHT, TRIM, SUBSTITUTE and REPT exist in current Excel and Google Sheets, but neither app was opened for this post.

![Name list with first initial, last initial, initials and initials with periods](images/2026-11-04-team-03-initials-from-full-name-monogram-left-substitute-excel-google-sheets/initials-from-full-names-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. Put the names as First Last in column **A** from row 5.

1. In **B5**, type `=UPPER(LEFT(TRIM(A5),1))`. TRIM removes extra spaces, LEFT keeps the first character and UPPER makes it a capital.
2. In **C5**, type `=UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(A5)," ",REPT(" ",100)),100)),1))`.
3. In **D5**, type `=B5&C5` for the initials.
4. In **E5**, type `=B5&"."&C5&"."` for initials with periods.
5. Select **B5:E5** and fill down.

Step 2 works from the inside out. SUBSTITUTE replaces each single space with 100 spaces, so the words end up far apart. RIGHT takes the last 100 characters, which is the last word with a lot of spaces in front. TRIM removes those spaces, and LEFT with UPPER takes the first letter of what is left.

It looks at the last word, so middle names are skipped: Dana Marie Thompson gives DT.

In the template, each formula starts with a blank test and a one-word name shows Check instead of an error. The names are read in the order typed.

Help pages: Microsoft's [LEFT function](https://support.microsoft.com/en-us/office/left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c), [RIGHT function](https://support.microsoft.com/en-us/office/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f), [UPPER function](https://support.microsoft.com/en-us/office/upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6), [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332), [REPT function](https://support.microsoft.com/en-us/office/rept-function-04c4d778-e712-43b4-9c15-d656582bb061) and [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/03/2026; the pages were not read for this post.

## Example

Sample data is fictional: names for a run of engraved coasters.

| Full name | Initials | With periods |
|---|---|---|
| Maria Alvarez | MA | M.A. |
| kevin o'brien | KO | K.O. |
| Luis de la Cruz | LC | L.C. |
| `  Linh   Nguyen ` (extra spaces) | LN | L.N. |
| Dana Marie Thompson | DT | D.T. |
| Ruth Ann Johnson-Lee | RJ | R.J. |

Lower case, extra spaces before, between and after the words, and a last name with a particle (de la Cruz) all came out right in the test.

## Troubleshooting

### A one-word name shows Check
There is no last word to take a letter from. In the test, `Prince` alone gave Check in the last initial, initials and with periods columns, and the **CHECK** cell said Check until the row was fixed.

### The letters are in the wrong order
The name was typed Last, First. In the test, `Alvarez, Maria` gave AM, the first letter of Alvarez and then of Maria. Put the first name first before you use the sheet.

### A suffix is used as the last name
Jr. counts as the last word. In the test, `Maria Alvarez Jr.` gave MJ. Delete the suffix from the name cell, or type the name without it.

### A name starting with a digit or an accent
The formula takes whatever character comes first. In the test, `3M Company` gave 3C, and a name with accented letters kept them.

### I want the middle initial too
This template skips middle names. A name with a middle name still gives two letters.

## Template

[Download the .xlsx](templates/tidy-tabs-initials-from-names.xlsx)

The file has two tabs. **Names** has 100 rows from row 5 (yellow names in **A**, formulas in **B:E**) and a summary in **G4:H7**: names typed, initials made, rows showing Check and a **CHECK** cell. **How to use** has notes on each formula.

**CHECK** says OK when every typed name gave initials. Not tested: non-breaking spaces copied from a web page.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
