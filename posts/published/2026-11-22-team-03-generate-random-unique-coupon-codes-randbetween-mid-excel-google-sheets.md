---
threads_url: https://www.threads.com/@tin_ylab/post/DgIB_mZG3SS
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-generate-random-coupon-codes-in.html
title: How to Generate Random Coupon Codes in Excel and Google Sheets
labels: excel-formulas, small-business, text-functions
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 33 workbooks and 83,094 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Codes are random, so the test checked format, length, allowed characters and the duplicate counts, not fixed codes. Not tested in Excel or Google Sheets.
search_description: Make readable coupon codes with =MID(chars,RANDBETWEEN(1,LEN(chars)),1) after a prefix, check repeats with COUNTIF, and paste as values before printing.
image_prompts: A short stack of small printed discount cards with blurred unreadable print fanned out next to a kraft paper gift bag and a roll of round sticker labels on a craft fair checkout table, warm indoor light, realistic photo, no screens, no readable text, no logos
image_alt: Fanned stack of small discount cards beside a kraft gift bag and sticker roll
threads: Coupon codes that change after you print them are not great.\n=MID(chars,RANDBETWEEN(1,LEN(chars)),1) picks one character from a set you type, like one with no O, 0, I or 1.\nPaste the list as values before printing, and let COUNTIF catch repeats.
---
Pick one random character from a list you type with `=MID(chars,RANDBETWEEN(1,LEN(chars)),1)`, join six of those after a prefix with `="MKT-"&MID(chars,RANDBETWEEN(1,LEN(chars)),1)&...`, and paste the codes as values before you print or import them, because RANDBETWEEN picks new ones every time the sheet recalculates. With the character set ABCDEFGHJKLMNPQRSTUVWXYZ23456789 (no O, 0, I or 1) and the prefix "MKT-", you get 10-character codes such as MKT-3GBXMZ, and the download counts each code in the list so a repeat shows up before it goes on a flyer.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (33 workbooks, 83,094 cell comparisons, 0 mismatches), plus the edge cases below. The codes are random, so the test checked format, length, allowed characters and duplicate counts, not fixed codes. Excel and Google Sheets were not opened for this post.

![Settings tab with character set, MKT- prefix, six random characters and passing checks](images/2026-11-22-team-03-generate-random-unique-coupon-codes-randbetween-mid-excel-google-sheets/random-coupon-codes-template.png) *Coupon code template with fictional settings (a 32-character set, the prefix MKT-, 6 random characters, 200 codes) and its checks; any codes shown are one recalculation's example and change every time. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Cell addresses match the download.

1. On **Settings**, type the characters to pick from in **B1**. The default ABCDEFGHJKLMNPQRSTUVWXYZ23456789 has 32 characters and leaves out O, 0, I and 1, which customers mix up on a printed flyer.
2. Type a prefix in **B2**, such as MKT- for a market flyer.
3. Type how many random characters you want in **B3** (1 to 8) and how many codes in **B4** (1 to 500).
4. Check **B16**. Codes are made only while it shows OK. The set needs at least 10 characters with no repeats, spaces, `*`, `?` or `~`, and the prefix cannot hold `*`, `?` or `~` or start with `<`, `>` or `=` (COUNTIF reads those as wildcards or comparisons).
5. On **Codes**, **B5** builds one code. The core piece is `MID(Settings!$B$1,RANDBETWEEN(1,LEN(Settings!$B$1)),1)`: RANDBETWEEN picks a position from 1 to the length of the set, and MID takes the character at that position. See [RANDBETWEEN in Excel](https://support.microsoft.com/en-us/office/randbetween-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685), [RANDBETWEEN in Google Sheets](https://support.google.com/docs/answer/3093507) and [MID in Excel](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028).
6. In a sheet of your own, with the set in **B1**, the short form for six characters is `="MKT-"&MID($B$1,RANDBETWEEN(1,32),1)&MID($B$1,RANDBETWEEN(1,32),1)&MID($B$1,RANDBETWEEN(1,32),1)&MID($B$1,RANDBETWEEN(1,32),1)&MID($B$1,RANDBETWEEN(1,32),1)&MID($B$1,RANDBETWEEN(1,32),1)`. The 32 is the length of the set.
7. The template's **B5** joins eight pieces, each wrapped like `IF(Settings!$B$3>=1,MID(...),"")`, so **B3** decides how many are used. An outer `IF(OR(Settings!$B$16<>"OK",ROWS(B$5:B5)>Settings!$B$4),"",...)` stops at the count in **B4** and stays blank while the settings fail.
8. In **C5**, `=IF(B5="","",COUNTIF($B$5:$B$504,B5))` counts how often each code appears. Every row should show 1.
9. Before you print or import, select **B5:B504**, copy, click **D5** and paste as values only (Paste Special).
10. **E5** runs the same COUNTIF on column **D**, **F5** shows OK, Duplicate or Check: length, and **I11** turns OK once the pasted list is clean.

These codes are random, not secret or secure. Do not use them for passwords, gift cards, payments or anything else that needs security. For importing, follow your shop's own help on discount codes.

## Example

The settings below are fictional. The checks on the Settings tab do not change between recalculations:

| Check | Cell | Value |
|---|---|---|
| Characters in the set | B7 | 32 |
| Characters that appear more than once | B8 | 0 |
| Spaces, * ? or ~ in the set | B9 | 0 |
| Code length (prefix + random) | B14 | 10 |
| Possible different codes | B15 | 1,073,741,824 |
| Settings CHECK | B16 | OK |

The codes changed on each of three recalculations in LibreOffice. Each run gave 200 distinct codes, all 10 characters from the 32 allowed, with 1 in every row of column **C**. **I4:I11** read 200, 200, 0, 0, 0, 0, OK and Check; the last stays Check until you paste into **D**.

Other settings from the edge cases:

| Prefix | Random characters | Code length | Possible codes |
|---|---|---|---|
| MKT- | 6 | 10 | 1,073,741,824 |
| SAVE | 4 | 8 | 1,048,576 |
| (blank) | 6 | 6 | 1,073,741,824 |
| MKT- | 8 | 12 | 1,099,511,627,776 |

Possible codes is the set length to the power of the random characters (32^6 for the sample). A repeat is rare but possible, so column **C** is the guard.

## Troubleshooting

### The codes changed after I printed them
Column **B** recalculates, and RANDBETWEEN picks new characters each time. Paste the list into column **D** as values first, and print or export column **D**.

### No codes appear
**Settings B16** shows Check. Causes: 9 or more in **B3** (the formula has eight pieces), a repeated letter or space in the set, fewer than 10 characters, a number typed as text, or more than 500 in **B4**. **B10:B13** shows which.

### A code shows Duplicate
Two codes match. A short length makes this common: 1 random character with the prefix X gave only 32 distinct codes out of 300. Type a new code over one in column **D**. In LibreOffice the check ignores case, so MKT-ABCDEF and mkt-abcdef count as the same code; keep the set in one case.

### A pasted code shows Check: length
The code in column **D** is not the length of the prefix plus the random characters (**Settings B14**). For example, MKT-ABC with a 10-character setting.

### Codes from CHAR(RANDBETWEEN(65,90)) are hard to read
It picks any letter A to Z, including O and I: in LibreOffice, 36 of 500 cells showed O or I. A typed set controls what appears.

## Template

[Download the .xlsx](templates/tidy-tabs-random-coupon-codes.xlsx)

The file has three tabs. **Settings** holds the inputs in **B1:B4** (character set, prefix, random characters, how many codes) and the checks in **A6:B16**, ending with the Settings CHECK.

**Codes** has up to 500 rows (5 to 504) with # in **A**, the live code in **B**, Times in live list in **C**, Pasted codes (paste values here) in **D**, Times in pasted list in **E** and Status in **F**. The summary in **H4:I11** counts live and pasted codes, repeats in each, wrong lengths, and both CHECK cells. **How to use** explains each formula.

Yellow cells are inputs, gray cells are formulas, and every Check and Duplicate turns red. Both tabs print Landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
