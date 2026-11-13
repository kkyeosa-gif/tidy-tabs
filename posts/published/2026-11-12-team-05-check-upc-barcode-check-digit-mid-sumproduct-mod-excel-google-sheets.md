---
threads_url: https://www.threads.com/@tin_ylab/post/DfrtN0FFEhg
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-check-upc-barcode-numbers-check.html
title: How to Check a UPC Barcode Number's Check Digit in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 8 workbooks and 8,056 cell comparisons with 0 mismatches. Checked: a valid code, its first 11 and last 11 digits, the code stored as a number, 10, 13 and 14 digits, a wrong last digit, X as the last digit, a letter O for zero, a neighbor swap, spaces and hyphens in printed groups, double hyphens, an underscore, a decimal, 000000000000, 999999999993, a full width digit, a leading +, blank rows with and without a product, and 6 random sets of 200 rows. Excel and Google Sheets were not used. Not run: a non-US language setting, more rows than the template holds.
search_description: Check the 12th digit of a UPC-A code with MID, SUMPRODUCT and MOD, and flag wrong check digits and lost leading zeros. Free Excel template.
image_prompts: A roll of blank barcode sticker labels beside a stack of kraft product boxes and a hand-held barcode scanner lying on a wooden packing table, any bars blurred, soft daylight from a workshop window, realistic photo, no screens, no readable text, no logos
image_alt: Roll of blank sticker labels next to kraft boxes and a handheld scanner on a packing table
threads: Swap two neighboring digits in a UPC and the check digit can still come out right.\nIn the made-up sample, 084512300162 and 084512300612 both pass.\nThe math: triple the digits in odd places, add the even ones, and the 12th digit tops the total up to the next multiple of 10. It checks arithmetic, not who owns the code.
---
Work out the 12th digit of a UPC-A code with `=MOD(10-MOD(SUMPRODUCT(--MID(C4,{1,2,3,4,5,6,7,8,9,10,11},1),{3,1,3,1,3,1,3,1,3,1,3}),10),10)`, where C4 holds the code as text, and compare it with the last digit you have. In the made-up sample, 084512300179 passes and 084512300328 does not (the first 11 digits give 2).

This checks the arithmetic only. A matching check digit does not mean a code is registered, assigned to your company or tied to any product. Company prefixes for real UPCs come from [GS1 US](https://www.gs1us.org); do not make up codes for products you sell.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every result below was recalculated and compared with a Python model (8 workbooks, 8,056 comparisons). The formulas use only IF, OR, ISNUMBER, FIND, MID, RIGHT, LEN, TRIM, SUBSTITUTE, SUMPRODUCT, MOD and COUNTIF, which also exist in Excel and Google Sheets, but neither was opened for this post. The open question there is SUMPRODUCT with the `{1,2,...}` array constants inside MID.

![Twelve fictional product codes with computed check digits, results and notes](images/2026-11-12-team-05-check-upc-barcode-check-digit-mid-sumproduct-mod-excel-google-sheets/upc-check-digit-template.png) *Twelve fictional product codes with the computed check digit, the result and a note. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Barcodes** tab and a **How to use** tab.

1. Format column **B** as text before you paste. A UPC stored as a number loses its leading 0, the same problem as in [How to Keep Leading Zeros in Excel ZIP Codes](https://tidytabs.blogspot.com/2026/09/how-to-keep-leading-zeros-in-excel-zip.html). Column B in the template is already set to text.
2. Paste product names into **A4** and UPCs into **B4** and down. The template holds rows 4 to 203.
3. In **C4**, type `=SUBSTITUTE(SUBSTITUTE(TRIM(B4)," ",""),"-","")`. This removes spaces and hyphens, so a code copied from a printed label such as `0 84512 30018 6` or `0-84512-30025-4` reads as 12 digits.
4. In **D4**, type `=IF(C4="","",LEN(C4))` to count the digits.
5. In **E4**, type the check digit formula from the top of this page. It uses only the first 11 digits.
6. In **F4**, compare: `=IF(--RIGHT(C4,1)=E4,"OK","Wrong check digit")`.
7. Fill **C4:F4** down to the last code.

How the math works: digits 1, 3, 5, 7, 9 and 11 count three times, and digits 2, 4, 6, 8 and 10 count once. The check digit is whatever brings that total up to the next multiple of 10, or 0 if the total already is one.

In the formula, `MID(C4,{1,2,...,11},1)` pulls out each of the first 11 characters, `--` turns them into numbers, and `SUMPRODUCT` multiplies them by the `{3,1,3,...}` weights and adds them up. `MOD(...,10)` keeps the last digit of the total, `10-` gives the distance to the next multiple of 10, and the outer `MOD(...,10)` turns a 10 into 0.

Worked example for 084512300179. The first 11 digits are 0 8 4 5 1 2 3 0 0 1 7. Odd positions: 0+4+1+3+0+7 = 15, times 3 is 45. Even positions: 8+5+2+0+1 = 16. Total 61, and 70 minus 61 is 9. The last digit is 9, so the template shows 9 and OK.

The template wraps these formulas. **E4** only computes when there are 11 or 12 characters and the first 11 are all digits. **F4** shows OK, Wrong check digit, 11 digits, Blank UPC (a product with no code) or Check (wrong length, a letter, or another character). **G4** adds a note for the Wrong check digit and 11 digits rows.

Help pages: Microsoft's [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) and [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028), and Google's [SUMPRODUCT](https://support.google.com/docs/answer/3094294). They describe the functions, not our test.

## Example

Sample data is fictional: 12 made-up products and codes.

| Product | UPC as typed | Length | Check digit | Result | Note |
|---|---|---|---|---|---|
| Lavender soy candle 8 oz | 084512300179 | 12 | 9 | OK | |
| Cedar soy candle 8 oz | 0 84512 30018 6 | 12 | 6 | OK | |
| Oatmeal goat milk soap | 0-84512-30025-4 | 12 | 4 | OK | |
| Charcoal soap bar | 084512300328 | 12 | 2 | Wrong check digit | Last digit is 8, the first 11 give 2 |
| Brass hoop earrings | 84512300407 (stored as a number) | 11 | 4 | 11 digits | 084512300407 passes with a 0 in front; or add 4 at the end |
| Letterpress card set | 08451230050 | 11 | 6 | 11 digits | Add 6 at the end: 084512300506 |
| Ceramic mug 12 oz | 0845123006120 | 13 | | Check | |
| Macrame plant hanger | O84512300759 (letter O) | 12 | | Check | |
| Leather key fob | (blank) | | | Blank UPC | |
| Canvas tote bag | 084512300896 | 12 | 6 | OK | |
| Kraft gift box | 0845123009 | 10 | | Check | |
| Beeswax wrap set | 084512 300933 | 12 | 3 | OK | |

The summary block in **I3:J9** shows 12 UPCs listed, 5 OK, 1 Wrong check digit, 2 with 11 digits, 3 Check, 1 Blank UPC, and a CHECK cell that says OK because the five counts add up to the total.

The two 11 digits rows show why that result needs a person. For the earrings, both fixes pass: a 0 in front, or a 4 at the end. For the card set, only adding 6 at the end passes. The note tells you what fits; only you know which code you were actually given.

## Troubleshooting

### A good code shows 11 digits
The code was stored as a number and lost its leading 0. In the sample, 84512300407 shows 11 digits and the note says 084512300407 passes with a 0 in front. Format column B as text and paste the code again.

### A code that looks right shows Check
Something in it is not a plain digit or the length is off. The Macrame row has a letter O instead of a zero, and a full width digit from some web pages does the same. The mug has 13 digits and the gift box has 10. A 13-digit EAN or GTIN-13 code shows Check because this file covers 12-digit UPC-A only.

### The result says Wrong check digit
The last digit does not match the first 11. For 084512300328, the first 11 give 2 but the code ends in 8. That means at least one digit is wrong somewhere, not always the last one, so compare the whole code with the one you were given.

### A code with a typo still shows OK
The check does not catch every mistake. Swapping two neighboring digits that differ by 5 still passes: 084512300162 and 084512300612 both show OK. Other swaps, such as 8 and 4, were caught. OK means the arithmetic fits, not that the code is the right one.

## Template

[Download the .xlsx](templates/tidy-tabs-upc-check-digit.xlsx)

The file has two tabs. **Barcodes** has products in **A4:A203**, UPCs in **B4:B203** (formatted as text), the formula columns **C:G** (Digits only, Length, Check digit from first 11, Result, Note), and a summary with a CHECK cell in **I3:J9**. **How to use** explains each column and the math.

Yellow cells are inputs and gray cells are formulas. The Result cell is green for OK and red for Check, Wrong check digit, 11 digits and Blank UPC. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
