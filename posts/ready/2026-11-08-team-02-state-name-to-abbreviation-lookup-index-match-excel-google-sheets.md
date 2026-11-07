---
title: How to Look Up a State Abbreviation From the Full State Name in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 9 workbooks and 1,800 comparisons with 0 mismatches. Checked: the sample, no customers, all 51 names typed as is, in capitals, in lowercase and padded with spaces, non-breaking spaces, doubled inner spaces, tx/Tx/tX, valid and invalid codes, Calif., CA., blank and spaces-only cells, numbers, Puerto Rico, PR, Guam, the characters T*, *, ??, ? and ~, and random sets of 200 rows. Excel and Google Sheets were not used. Not run: UNICHAR in Excel 2010 or earlier, more than 200 rows.
search_description: Turn Texas into TX with INDEX and MATCH in Excel and Google Sheets, even with stray spaces or capitals. Free template marks unknown states Check.
image_prompts: A folded paper road map, a stack of blank padded mailers and a tape gun on a packing bench, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Folded paper road map beside padded mailers and a tape gun on a packing bench
threads: Free typed state names are a mess to clean up.\nTRIM plus SUBSTITUTE(C2,UNICHAR(160)," ") strips stray and non-breaking spaces, then INDEX/MATCH returns the code. " ohio " gives OH.\nA name not on the list, like "Calif.", shows Check instead of a blank.
---
Look the code up with `=INDEX(States!$B$2:$B$52,MATCH(D2,States!$A$2:$A$52,0))` after cleaning the typed text in D2 with `=TRIM(SUBSTITUTE(C2,UNICHAR(160)," "))`. The typed text " ohio " returns OH, and a name that is not on the list, such as "Calif.", returns Check instead of a blank.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every result below was recalculated and compared with a Python model (9 workbooks, 1,800 comparisons). The formulas use INDEX, MATCH, TRIM, SUBSTITUTE, UNICHAR and IFERROR, which also exist in Excel (UNICHAR needs Excel 2013 or later) and Google Sheets, but neither was opened for this post.

![Customer list with typed state names turned into two-letter codes, three rows flagged](images/2026-11-08-team-02-state-name-to-abbreviation-lookup-index-match-excel-google-sheets/state-name-to-abbreviation-lookup-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional.*

## Steps

The steps are the same in Excel and Google Sheets, apart from the notes below. Build two tabs. **States** holds the list: state name in **A2:A52**, two-letter code in **B2:B52**. That is 50 states plus the District of Columbia, in alphabetical order. **Customers** holds your orders: name in **A**, city in **B**, the state as typed in **C**, starting in row 2.

1. On **States**, type or paste the 51 names and codes. The codes come from the USPS Publication 28 list of two-letter abbreviations. Territories and military codes are left out on purpose.
2. On **Customers**, in **D2**, type `=TRIM(SUBSTITUTE(C2,UNICHAR(160)," "))`. This swaps non-breaking spaces (common in pasted web text) for normal ones, then trims the ends and collapses doubled spaces.
3. In **E2**, type `=IFERROR(INDEX(States!$B$2:$B$52,MATCH(D2,States!$A$2:$A$52,0)),"Check")`. `MATCH(...,0)` finds the exact name and ignores capitals. `INDEX` returns the code from the same row.
4. Some customers already typed a code, such as tx or OR. Wrap the fallback so a valid code is accepted: replace `"Check"` with `IF(AND(LEN(D2)=2,SUMPRODUCT(--(States!$B$2:$B$52=D2))>0),UPPER(D2),"Check")`.
5. In **F2**, type `=IF(E2="Check","Check",IF(UPPER(D2)=E2,"Already a code","From name"))` to show where each code came from.
6. Fill **D2:F2** down to the last customer.

The short formulas above are the teaching version. The built template adds one more layer: it puts a `~` in front of any `~`, `*` or `?` in the typed text before `MATCH`, because `MATCH` with 0 reads `*` and `?` as wildcards. Without that, a typed "T*" would match Texas. The simplified version was not the one tested.

### In Excel

Type the formulas as shown. `UNICHAR` needs Excel 2013 or later. Microsoft describes the lookup functions on its [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) page and the cleanup on its [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) page.

### In Google Sheets

Type the same formulas. Google documents [MATCH](https://support.google.com/docs/answer/3093378) and [INDEX](https://support.google.com/docs/answer/3098242) in its Docs Editors Help. Those pages describe the functions, not our test.

## Example

Sample data is fictional: 12 customers, typed the way real order exports often look.

| Typed | Cleaned | Code | Status |
|---|---|---|---|
| Texas | Texas | TX | From name |
| " ohio " | ohio | OH | From name |
| NEW YORK | NEW YORK | NY | From name |
| tx | tx | TX | Already a code |
| Calif. | Calif. | Check | Check |
| (empty) | | Check | Check |
| nevada | nevada | NV | From name |
| OR | OR | OR | Already a code |
| ZZ | ZZ | Check | Check |
| Georgia | Georgia | GA | From name |

The summary shows 12 customers listed, 7 from a name, 2 already a code and 3 Check, and the **Rows add up** cell says OK. Florida and Colorado also come out as FL and CO, and the table leaves them out for length.

## Troubleshooting

### A row says Check but the state looks right
The text is probably an abbreviation with a period, such as "Calif." or "CA.". The template does not guess. Retype it as California or CA. A state typed as a number, or a name that is not on the list, such as Puerto Rico or Guam, also shows Check.

### The formula fails on pasted web text
The gap between words may be a non-breaking space, which `TRIM` alone does not remove. Use `UNICHAR(160)`, not `CHAR(160)`. In LibreOffice, `CHAR(160)` follows the system locale and did not match a real non-breaking space.

### A blank state row shows Check
That is intended. A completely empty row stays blank, but a row with a name and no state shows Check so you can find the gap.

### New York with two spaces still works
`TRIM` also collapses doubled spaces inside a name, so "New  York" returns NY. If you do not want that, remove `TRIM` and expect more Check rows.

### Wildcard characters match the wrong state
`MATCH` with 0 treats `*` and `?` as wildcards. If you skip the `~` escaping, "T*" returns TX. In the template, "T*", "??" and "?" all return Check.

## Template

[Download the .xlsx](templates/tidy-tabs-state-abbreviations.xlsx)

The file has three tabs. **States** holds the 51 names and codes, with a note that the source is the USPS Publication 28 list. **Customers** has 200 rows with 12 fictional sample customers, the cleaned state, the code and a status column, plus a summary in **H1:I5**: customers listed, from a name, already a code, Check rows, and a **Rows add up** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The template only converts spelling. It does not check whether an address, a city or a ZIP code is real.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
