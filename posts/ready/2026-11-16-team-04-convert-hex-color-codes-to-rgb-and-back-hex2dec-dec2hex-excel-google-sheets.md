---
title: How to Convert HEX Color Codes to RGB (and Back) in Excel and Google Sheets
labels: excel-formulas, small-business, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 13 workbooks and 5,993 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Convert HEX color codes to RGB with =HEX2DEC(MID(code,1,2)) and back with DEC2HEX(R,2). Handles #fff shorthand and bad codes. Free template.
image_prompts: A fan of blank paint swatch cards in sea green, sandy orange, deep teal and coral spread on a craft table, matching spools of satin ribbon and a few folded fabric squares laid beside them for color matching, bright daylight, realistic photo, no screens, no readable text, no logos
image_alt: Paint swatch cards fanned out beside matching ribbon spools and folded fabric squares
threads: Use DEC2HEX(R,2), not DEC2HEX(R), when you turn RGB back into HEX.\nThe 2 keeps the leading zeros: 5, 10, 15 becomes 050A0F, not 5AF.\nGoing the other way, =HEX2DEC(MID(code,1,2)) pulls the red out of 2E8B57: 46.
---
Split a 6-digit HEX code into red, green and blue with `=HEX2DEC(MID(code,1,2))`, `=HEX2DEC(MID(code,3,2))` and `=HEX2DEC(MID(code,5,2))`, and rebuild it with `="#"&DEC2HEX(R,2)&DEC2HEX(G,2)&DEC2HEX(B,2)`. In the made-up palette, #2E8B57 gives 46, 139, 87 and RGB 5, 10, 15 gives #050A0F because the ,2 keeps the leading zeros, and the download first strips the # and spaces, expands shorthand like #fff to FFFFFF and shows Check for a bad code such as #GGGGGG.

Use it for the brand palette behind your Etsy banners, labels and Canva designs, where one tool may ask for RGB and another for HEX. It converts codes; it does not match screen colors to print.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (13 workbooks, 5,993 cell comparisons, 0 mismatches), plus the edge cases below. HEX2DEC, DEC2HEX and MID also exist in Excel and Google Sheets, but neither app was opened for this post.

![Five palette colors with HEX codes, red green blue values, rebuilt codes and swatches](images/2026-11-16-team-04-convert-hex-color-codes-to-rgb-and-back-hex2dec-dec2hex-excel-google-sheets/hex-to-rgb-color-converter-template.png) *Five fictional palette colors: #2E8B57 splits into 46, 139, 87, shorthand #fff becomes #FFFFFF, and a code pasted with spaces still converts (the helper and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Start in row 5 of the **HEX to RGB** tab.

1. Format column **B** as text before you type, so a code like 000080 is not turned into a number.
2. Type a color name in **A5** and the HEX code in **B5**, with or without the #.
3. In **C5**, type `=IF(TRIM(B5)="","",UPPER(SUBSTITUTE(TRIM(B5),"#","")))`. This removes outer spaces, every # and lower case.
4. In **D5**, type `=IF(LEN(C5)=3,MID(C5,1,1)&MID(C5,1,1)&MID(C5,2,1)&MID(C5,2,1)&MID(C5,3,1)&MID(C5,3,1),C5)`. A 3-digit shorthand such as FFF is doubled digit by digit to FFFFFF; other codes pass through.
5. In **E5**, type `=IF(D5="","",IF(LEN(D5)<>6,"Check",IFERROR(HEX2DEC(MID($D5,1,2)),"Check")))` for R. Copy it to **F5** and **G5** and change the start of MID to 3 for G and 5 for B.
6. In **H5**, type `=IF(COUNT(E5:G5)<3,"","#"&DEC2HEX(E5,2)&DEC2HEX(F5,2)&DEC2HEX(G5,2))` to rebuild the code as a check. The 2 keeps the leading zero: 5 becomes 05, not 5.
7. Fill **C5:H5** down to the last color.

For the other direction, go to the **RGB to HEX** tab, type R, G and B in **B5:D5**, and in **E5** type `="#"&DEC2HEX(B5,2)&DEC2HEX(C5,2)&DEC2HEX(D5,2)`.

A formula cannot color a cell, so the template's Swatch column is a fill set by hand. For a new color, pick a custom fill with your HEX code.

Help pages: Microsoft's [HEX2DEC function](https://support.microsoft.com/en-us/office/hex2dec-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e) and [DEC2HEX function](https://support.microsoft.com/en-us/office/dec2hex-function-6344ee8b-b6b5-4c6a-a672-f64666704619).

## Example

The palette is fictional, typed the way codes often arrive: lower case, short, without the # or with spaces.

| Color name | HEX code typed | 6-digit code | R | G | B | Check back | Status |
|---|---|---|---|---|---|---|---|
| Sea green | #2E8B57 | 2E8B57 | 46 | 139 | 87 | #2E8B57 | OK |
| Sandy orange | #f4a261 | F4A261 | 244 | 162 | 97 | #F4A261 | OK |
| White | #fff | FFFFFF | 255 | 255 | 255 | #FFFFFF | OK |
| Deep teal | 264653 | 264653 | 38 | 70 | 83 | #264653 | OK |
| Coral (pasted with spaces) | " #E76F51 " | E76F51 | 231 | 111 | 81 | #E76F51 | OK |

The summary in **L4:M9** shows 5 codes listed, 5 converted, 0 rows showing Check and CHECK OK.

The **RGB to HEX** tab goes the other way:

| Color name | R | G | B | HEX | Status |
|---|---|---|---|---|---|
| Sea green | 46 | 139 | 87 | #2E8B57 | OK |
| Sandy orange | 244 | 162 | 97 | #F4A261 | OK |
| Dark blue-gray | 5 | 10 | 15 | #050A0F | OK |
| White | 255 | 255 | 255 | #FFFFFF | OK |
| Black | 0 | 0 | 0 | #000000 | OK |

Black at 0, 0, 0 is a real color, so it gets #000000, while an empty row stays blank. CHECK in **I8** shows OK.

## Troubleshooting

### A code shows Check: length
After the # and spaces are removed, the code is not 3 or 6 digits: #12345 has 5, and an 8-digit code with transparency (#RRGGBBAA) has 8, which this sheet does not support. Without the length check, LibreOffice read #12345 as 18, 52, 5 and rebuilt it as #123405, a different color with no warning.

### A code shows Check: letters
HEX codes use only 0 to 9 and A to F, so #GGGGGG and #12Z456 are not colors. In LibreOffice, the bare HEX2DEC formula returned #VALUE! for #GGGGGG; the template wraps it in IFERROR so the row shows Check instead. A doubled # such as ##fff is fine: it reads as FFFFFF.

### A blank row shows 0, 0, 0 and #000000
Your formula has no blank guard. In LibreOffice, HEX2DEC of an empty cell returned 0, so an empty code came out as black. The template's `IF(D5="","",...)` leaves the row blank, and a name with no code shows Check: no code.

### The rebuilt code is too short, like #5AF
DEC2HEX was used without the 2, so 5, 10, 15 became 5, A and F instead of 05, 0A and 0F. Use `DEC2HEX(R,2)` for each part.

### An RGB row shows Check: whole numbers or Check: 0 to 255
RGB parts are whole numbers from 0 to 255. In LibreOffice, `DEC2HEX(127.5,2)` returned 7F with the decimal cut off and `DEC2HEX(256,2)` returned #VALUE!, so the template checks the numbers first and shows Check: whole numbers for 127.5, Check: 0 to 255 for 256 or -1, and Check: blank or text for a part typed as text such as "12".

## Template

[Download the .xlsx](templates/tidy-tabs-hex-rgb-color-converter.xlsx)

The file has three tabs. **HEX to RGB** holds rows 5 to 54: Color name and HEX code (inputs in **A:B**, with B formatted as text), then Clean code, 6-digit code, R, G, B and Check back (formulas in **C:H**), a Swatch column in **I** with fills set by hand for the sample rows only, and Status in **J** (OK, Check: length, Check: letters, Check: no code). The summary in **L4:M9** also counts rebuilt codes that differ and entries typed below row 54, which are not read.

**RGB to HEX** holds rows 5 to 54: Color name, R, G and B (inputs in **A:D**), then HEX and Status (formulas in **E:F**), with the summary in **H4:I8**. **How to use** explains each column. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Status turns red for any Check, and so does the CHECK cell when it is not OK. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
