---
threads_url: https://www.threads.com/@tin_ylab/post/DfnFKjjFlK0
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-check-zip-codes-for-right-format.html
title: How to Check ZIP Codes for the Right Format (5 Digits or ZIP+4) in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with independent Python models over 7 workbooks and 5,069 cell comparisons. Checked: the 15-row sample, 48 fixed strings (leading zeros lost, a letter O for zero, 1,234, $1234, 1E100, +1234, hyphens in the wrong place, spaces, tabs, full width digits, blanks), ZIPs stored as numbers, and 6 random sets of 200 rows. Excel and Google Sheets were not used. Not run: non-breaking spaces, a non-US language setting.
search_description: Check ZIP codes for 5-digit or ZIP+4 format with LEN, MID, FIND and SUMPRODUCT in Excel and Google Sheets. Free template flags missing leading zeros.
image_prompts: A small stack of addressed kraft envelopes with one address label peeled and torn, beside a rubber stamp and an ink pad on a wooden packing table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Stack of addressed envelopes with one torn label beside a rubber stamp
threads: 2134 in your ZIP column? That is 02134 with the 0 eaten.\nTRIM it, then test each character with FIND(MID(C4,{1,2,3,4,5},1),"0123456789").\nMy 15 made-up customers: 4 OK, 2 ZIP+4, 4 missing a zero, 4 Check, 1 blank. Format check only, not a ZIP lookup.
---
A ZIP code has the right shape when `=AND(LEN(C4)=5,SUMPRODUCT(--ISNUMBER(FIND(MID(C4,{1,2,3,4,5},1),"0123456789")))=5)` returns TRUE, where C4 holds `=TRIM(B4)`. In the sample, 02134 passes, 2134 is flagged as missing a leading zero with the fix 02134, and 12O45 (a letter O) is flagged Check (made-up customers). This checks the shape only. It does not tell you whether a ZIP exists.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every result below was recalculated and compared with Python models (7 workbooks, 5,069 comparisons). The formulas use only IF, AND, TRIM, LEN, MID, FIND, ISNUMBER and SUMPRODUCT, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Table of fifteen sample ZIP codes with result and one-zero fix suggestion columns](images/2026-11-11-team-02-check-zip-code-format-5-digit-zip-plus-4-len-isnumber-excel-google-sheets/check-zip-code-format-template.png) *Fifteen fictional ZIPs with the result and the one-zero fix suggestion. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Put one customer per row from row 4, with headers in row 3. **A** is the customer and **B** is the ZIP as typed.

1. In **C4**, type `=TRIM(B4)`. This removes spaces at the ends, so " 07030 " becomes 07030.
2. In **D4**, type `=IF(C4="","",LEN(C4))` for the length.
3. To test for 5 digits, use `=AND(LEN(C4)=5,SUMPRODUCT(--ISNUMBER(FIND(MID(C4,{1,2,3,4,5},1),"0123456789")))=5)`. MID pulls out characters 1 to 5 one at a time. FIND looks each one up in the string of digits. ISNUMBER turns each hit into TRUE, and SUMPRODUCT counts them. All 5 must be digits.
4. To test for ZIP+4 (such as 90210-1234), require `LEN(C4)=10` and `MID(C4,6,1)="-"`. Then count digits at positions `{1,2,3,4,5,7,8,9,10}` and expect 9.
5. To catch a lost zero, test for a length of 4 with 4 digits, or a length of 9 with a hyphen at position 5 and 8 digits. These are the ZIPs that Excel and Sheets shorten when a column is formatted as a number.
6. In the fix column, type `="0"&C4` for the rows flagged as missing a zero. It only adds one 0.
7. Fill the formulas down as far as your list goes.

Why test each character? A shorter idea is `ISNUMBER(--C4)`. It accepts text such as "1,234", "$1234", "+1234" and "1E100" as numbers, so a 5-character bad ZIP can pass. The per-character test does not.

The template combines all of these in one formula in **E4** and returns one of: OK, OK (ZIP+4), Missing leading zero, Check or Blank ZIP.

Help pages: Microsoft's [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) and [FIND function](https://support.microsoft.com/en-us/office/find-find-b-function-c7912941-af2a-4bdf-a553-d0d89b0a0628). They describe the functions, not our test.

## Example

Sample data is fictional: 15 customers with made-up ZIPs.

| ZIP typed | Result | Fix |
|---|---|---|
| 02134 | OK | |
| 2134 | Missing leading zero | 02134 |
| 90210-1234 | OK (ZIP+4) | |
| 9021 | Missing leading zero | 09021 |
| 1234567 | Check | |
| 12O45 (letter O) | Check | |
| " 07030 " | OK | |
| 60614 | OK | |
| 10001-0001 | OK (ZIP+4) | |
| 2134-5678 | Missing leading zero | 02134-5678 |
| 8101 | Missing leading zero | 08101 |
| (blank, customer listed) | Blank ZIP | |
| 331 | Check | |
| 30301-12 | Check | |
| 55401 | OK | |

The summary block counts 15 ZIPs listed, 4 OK, 2 OK (ZIP+4), 4 Missing leading zero, 4 Check and 1 Blank ZIP. The CHECK cell says OK when those counts add up to 15.

2134 and 8101 are 4-digit ZIPs, so the formula suggests 02134 and 08101. A 3-digit entry such as 331 shows Check, because more than one zero may be gone and a guess would be wrong.

## Troubleshooting

### 00000 passes
The test checks the shape, not the place. 00000 is five digits, so it shows OK. To find out whether a ZIP is real for an address, use the [USPS ZIP Code lookup](https://tools.usps.com/zip-code-lookup.htm?byaddress).

### The fix is wrong
The fix adds one 0 and only suggests. The original may have lost more than one zero, or may be a different ZIP. Check the address before you overwrite anything.

### My ZIPs are numbers, and the zeros are already gone
A ZIP stored as the number 2134 is already damaged, and the sheet shows Missing leading zero for it. That is the expected result. To stop it happening again, format column **B** as Plain text before you paste. Our post on [keeping leading zeros in ZIP codes](https://tidytabs.blogspot.com/2026/09/how-to-keep-leading-zeros-in-excel-zip.html) has the steps for Excel and Google Sheets.

### A ZIP that looks fine shows Check
Look for a letter O in place of a zero, a letter after the digits, a comma, a $ sign, an exponent such as 1E100, or full width digits. All of these show Check. Spaces inside the number also show Check. Non-breaking spaces and tabs were not tested.

### The CHECK cell shows Check
The summary counts do not add up to the number of ZIPs listed. Scroll for a row that has a customer name but a formula that was overwritten or not filled down.

## Template

[Download the .xlsx](templates/tidy-tabs-check-zip-format.xlsx)

The file has two tabs. **ZIPs** has 200 slots from row 4 to row 203 and 15 fictional sample customers. Columns **A** and **B** hold the customer and the ZIP (column B is formatted as text). Columns **C** to **F** hold the cleaned ZIP, Length, Result and Fix suggestion. A summary block at the right counts ZIPs listed, OK 5 digits, OK ZIP+4, Missing leading zero, Check and Blank ZIP, plus an overall CHECK. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
