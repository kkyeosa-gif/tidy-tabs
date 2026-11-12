---
title: How to Write a Check Amount in Words in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 18 workbooks and 14,472 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Spell out a check amount in words, such as One Thousand Two Hundred Fifty and 00/100, with INDEX, INT and MOD and no macro. Free template.
image_prompts: An open business checkbook on a wooden desk with a fountain pen resting across a blank check, a small binder-clipped stack of supplier invoices beside it, all handwriting and print unreadable.
image_alt: Open business checkbook with a fountain pen resting across a blank check
threads: Write check amounts in words with plain formulas, no macro: INT and MOD split the thousands from the rest, INDEX picks each word from a list, and TEXT(F4,"00") keeps the 0 in 07/100.\n$2,019.07 comes out as Two Thousand Nineteen and 07/100.\nStill read the words against the number before you sign.
---
Turn a dollar amount into check wording, such as $1,250.00 into One Thousand Two Hundred Fifty and 00/100, with plain formulas and no macro: split the dollars into thousands `INT(E4/1000)` and the rest `MOD(E4,1000)`, write each part with INDEX from a list of number words, and add the cents as `TEXT(F4,"00")&"/100"`. The download has the full formula in column G and works up to $999,999.99.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (18 workbooks, 14,472 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither was opened; INDEX on another tab and TEXT with "00" are the open questions there.

![Check register with each dollar amount spelled out in words and cents over 100](images/2026-11-13-team-04-write-check-amount-in-words-dollars-and-cents-index-mod-excel-google-sheets/check-amount-in-words-template.png) *Twelve checks with the amount written out; fractions of a cent, a negative and a text amount show Check. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has three tabs: **Checks**, **Words** and **How to use**.

1. On the **Words** tab, list the numbers 0 to 19 in **A4:A23** with their words in **B4:B23** (Zero, One, Two ... Nineteen).
2. In **D4:D13**, list the tens 0 to 90, and in **E4:E13** their words: leave E4 blank, then Ten, Twenty, Thirty ... Ninety.
3. On the **Checks** tab, type the check number in **A4**, the date in **B4**, the payee in **C4** and the amount in **D4**.
4. In **E4**, get the whole dollars from the amount rounded to cents. The core is `=(ROUND(D4*100,0)-MOD(ROUND(D4*100,0),100))/100`.
5. In **F4**, get the cents: `=MOD(ROUND(D4*100,0),100)`.
6. In **G4**, paste the Amount in words formula from the template.
7. Fill **E4:G4** down to the last check.

The template's E4 shows Check instead of a number for a text, negative, too large or fractional-cent amount. **H4** then says Check and G4 stays empty.

### How the words formula works

G4 is long, but it is one piece used twice. For any number from 0 to 999, this piece writes it out (here for a number in M4):

`=TRIM(IF(INT(M4/100)>0,INDEX(Words!$B$4:$B$23,INT(M4/100)+1)&" Hundred","")&" "&IF(MOD(M4,100)=0,"",IF(MOD(M4,100)<20,INDEX(Words!$B$4:$B$23,MOD(M4,100)+1),INDEX(Words!$E$4:$E$13,INT(MOD(M4,100)/10)+1)&IF(MOD(M4,10)>0,"-"&INDEX(Words!$B$4:$B$23,MOD(M4,10)+1),""))))`

- **Hundreds.** `INT(M4/100)` is the hundreds digit. If it is above 0, INDEX picks that word from the list and adds " Hundred". The `+1` is there because Zero is the first word in the list.
- **1 to 19.** `MOD(M4,100)` is the last two digits. Under 20, INDEX picks the word straight from the list.
- **20 to 99.** INDEX picks the tens word from the second list with `INT(MOD(M4,100)/10)+1`, then adds a hyphen and the ones word when `MOD(M4,10)` is not 0: Forty-Five.

In the sample, check 1045 ($312.75) reads Three Hundred Twelve and check 1048 ($700.10) reads Seven Hundred.

G4 runs the piece on `INT(E4/1000)` plus " Thousand", then on `MOD(E4,1000)`, and adds the cents. Amounts under $1.00 get Zero.

Help pages: Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3), and Google's [TEXT](https://support.google.com/docs/answer/3094139). Microsoft's own number-to-words method uses a VBA macro; this one uses formulas only.

## Example

The sample checks and payees are fictional.

| Check no. | Date | Pay to | Amount | Amount in words | Status |
|---|---|---|---|---|---|
| 1041 | 11/02/2026 | Cedar Print Shop | $1,250.00 | One Thousand Two Hundred Fifty and 00/100 | OK |
| 1042 | 11/03/2026 | Harbor Packaging Supply | $86.50 | Eighty-Six and 50/100 | OK |
| 1043 | 11/05/2026 | Lakeside Studio Rentals | $2,019.07 | Two Thousand Nineteen and 07/100 | OK |
| 1044 | 11/06/2026 | Pine Street Market | $45.00 | Forty-Five and 00/100 | OK |
| 1045 | 11/09/2026 | Maria Lopez | $312.75 | Three Hundred Twelve and 75/100 | OK |
| 1046 | 11/10/2026 | Jordan Reed | $0.99 | Zero and 99/100 | OK |
| 1047 | 11/12/2026 | Main St. Gift Shop | $15,600.00 | Fifteen Thousand Six Hundred and 00/100 | OK |
| 1048 | 11/13/2026 | Blue Ridge Wax Supply | $700.10 | Seven Hundred and 10/100 | OK |
| 1049 | 11/16/2026 | Sunrise Sign Co. | $4,080.40 | Four Thousand Eighty and 40/100 | OK |
| 1050 | 11/17/2026 | Harbor Dental | 12.345 (shows $12.35) | | Check |
| 1051 | 11/18/2026 | Lopez family | -$25.00 | | Check |
| 1052 | 11/19/2026 | Lakeside Yoga | "$45.00" (text) | | Check |

The summary in **J3:K6** shows 12 checks listed, 9 written out, 3 rows showing Check, and $24,094.81 as the total of the written-out checks.

The wording follows one style: capitalized words, a hyphen from 21 to 99, "and" only before the cents, cents as 00/100, and no "Dollars" at the end. To add the word, put `&" Dollars"` at the end of the G4 formula; check 1043 then reads Two Thousand Nineteen and 07/100 Dollars. To change the spelling, for example to lower case, edit the words on the **Words** tab.

Always read the words against the number before you sign. The file writes words only. It does not print checks, and your bank decides what it accepts.

## Troubleshooting

### An amount that looks fine shows Check
It has a third decimal or is text. In the sample, 12.345 displays as $12.35 and "$45.00" was pasted as text. Round calculated amounts to cents first.

### A negative or very large amount shows Check
The file covers $0.00 to $999,999.99 only. -$25.00 shows Check, and so does $1,000,000.00 or more.

### The cents read 7/100 instead of 07/100
Joining F4 directly drops the zero. `TEXT(7,"00")` gives 07.

### The cents seem one cent off
0.29 times 100 can land a hair under 29, so the file rounds first: `ROUND(0.29*100,0)` gives 29, and $0.29 ends in 29/100.

### A row with a payee but no amount shows Check
A payee or date without an amount is flagged. Type the amount in D or clear the row.

## Template

[Download the .xlsx](templates/tidy-tabs-check-amount-in-words.xlsx)

**Checks** holds rows 4 to 203: Check no., Date, Pay to and Amount (inputs in **A:D**), then Dollars, Cents, Amount in words and Status (formulas in **E:H**), with the summary in **J3:K6**. It prints Landscape on US Letter. **Words** has the two editable word lists, and **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas; the Status cell turns red for Check. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
