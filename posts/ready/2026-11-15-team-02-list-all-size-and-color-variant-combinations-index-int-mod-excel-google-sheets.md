---
title: How to List Every Size and Color Combination in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 27 workbooks and 32,616 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: List all combinations of two lists with INDEX, INT and MOD: 3 sizes and 4 colors give 12 variant rows with a price per size. Free Excel template.
image_prompts: Twelve empty glass candle vessels on a wooden shelf in a small candle studio, arranged in three rows by size, small, medium and large, and four columns by glass tint, clear, black, amber and sage green, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Twelve empty glass candle vessels in rows by size and columns by tint
threads: 3 sizes x 4 colors = 12 variant rows, and you type none of them.\nSize: =INDEX(sizes,INT((n-1)/colors)+1)\nColor: =INDEX(colors,MOD(n-1,colors)+1)\nPaste the list as values before you add stock counts. A new color shifts every row after the first size.
---
Number the rows 1, 2, 3 and pull each pairing with `=INDEX(sizes,INT((n-1)/colors)+1)` for the size and `=INDEX(colors,MOD(n-1,colors)+1)` for the color, where n is the row number and colors is how many colors you have. In the made-up candle shop sample, 3 sizes and 4 colors give 12 rows from 8 oz / Natural to 16 oz / Sage, each with its size price of $14.00, $18.00 or $24.00, and the download wraps the formulas so rows past the last combination stay blank.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with a Python model (27 workbooks, 32,616 comparisons, 0 mismatches). Excel and Google Sheets have the same functions, but neither app was opened for this post.

![Twelve candle variants pairing 8, 12 and 16 oz sizes with four colors and prices](images/2026-11-15-team-02-list-all-size-and-color-variant-combinations-index-int-mod-excel-google-sheets/size-color-variant-combinations-template.png) *Three sizes and four colors become twelve variant rows, each with its size's price and an OK status (the On hand column and the summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*

## Steps

This lists all combinations of two lists with plain formulas: no Power Query, no macros. INT steps through the sizes slowly and MOD cycles through the colors quickly, so one row counter gives every pairing exactly once.

### Set up the two lists

1. On a tab named **Options**, type the sizes in **A5:A14** and the price for each size next to it in **B5:B14**.
2. Type the colors in **D5:D24**.
3. Keep each list together, with no blank cell between options.

### Count the options

1. On a second tab, type `=COUNTA(Options!$A$5:$A$14)` in **B1** to count the sizes.
2. Type `=COUNTA(Options!$D$5:$D$24)` in **B2** to count the colors.
3. Type `=B1*B2` in **B3**. That is how many rows the list needs: 3 x 4 = 12.

### List the combinations

1. In **A6**, type `=ROWS($B$6:B6)`. It returns 1 and counts up as you fill down. This is n.
2. In **B6**, type `=INDEX(Options!$A$5:$A$14,INT((ROWS($B$6:B6)-1)/$B$2)+1)` for the size. With 4 colors, rows 1 to 4 get size 1, rows 5 to 8 get size 2, and so on.
3. In **C6**, type `=INDEX(Options!$D$5:$D$24,MOD(ROWS($B$6:B6)-1,$B$2)+1)` for the color. MOD counts 1, 2, 3, 4, 1, 2, 3, 4, so every size gets every color.
4. In **D6**, type `=B6&" / "&C6` for a variant name such as 8 oz / Natural.
5. In **E6**, type `=INDEX(Options!$B$5:$B$14,INT((ROWS($B$6:B6)-1)/$B$2)+1)`. It uses the same position as the size, so each row picks up its own size price.
6. Select **A6:E6** and fill down exactly B3 rows.

If you fill further than B3 rows, the short formulas run past the lists. The template starts each one with `IF(ROWS($B$6:B6)>$B$3,"",...)`, so you can fill 200 rows and the extra ones stay blank. The list then grows by itself when you add a size or color.

Help pages: Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3), and Google's [INDEX](https://support.google.com/docs/answer/3098242) and [MOD](https://support.google.com/docs/answer/3093497). They describe the functions, not our test.

### Before you count stock

Column **F** (On hand) is typed, not a formula, so it does not move when the list changes. Adding a size at the bottom only adds rows at the end. Adding a color, even at the bottom, changes the color count, so every row after the first size moves: with a fifth color, 12 oz / Natural goes from No. 5 to No. 6. Copy the list and paste it as values before you type counts next to it.

## Example

Sample data is fictional: a candle shop with sizes 8 oz ($14.00), 12 oz ($18.00) and 16 oz ($24.00), and colors Natural, Black, Amber and Sage. B1 shows 3, B2 shows 4 and B3 shows 12.

| No. | Size | Color | Variant name | Price | Status |
|---|---|---|---|---|---|
| 1 | 8 oz | Natural | 8 oz / Natural | $14.00 | OK |
| 2 | 8 oz | Black | 8 oz / Black | $14.00 | OK |
| 3 | 8 oz | Amber | 8 oz / Amber | $14.00 | OK |
| 4 | 8 oz | Sage | 8 oz / Sage | $14.00 | OK |
| 5 | 12 oz | Natural | 12 oz / Natural | $18.00 | OK |
| 6 | 12 oz | Black | 12 oz / Black | $18.00 | OK |
| 7 | 12 oz | Amber | 12 oz / Amber | $18.00 | OK |
| 8 | 12 oz | Sage | 12 oz / Sage | $18.00 | OK |
| 9 | 16 oz | Natural | 16 oz / Natural | $24.00 | OK |
| 10 | 16 oz | Black | 16 oz / Black | $24.00 | OK |
| 11 | 16 oz | Amber | 16 oz / Amber | $24.00 | OK |
| 12 | 16 oz | Sage | 16 oz / Sage | $24.00 | OK |

Row 13 onward stays blank. The summary shows 12 combinations, 12 rows listed, 0 rows not OK, 0 options typed past the yellow rows, and CHECK OK.

Numbers work as sizes too. Sizes typed as 8, 12 and 16 give names such as 8 / Black.

## Troubleshooting

These cases were reproduced in LibreOffice 24.2.7.2.

### The last color is missing and one variant has no color
A blank cell sits inside the color list. COUNTA counts one less, so the last color drops off and the blank comes through as an empty color (the short formula showed 0). With Natural, blank, Amber, Sage, the list had 9 rows and no Sage. Delete the blank cell so the list is together; the template marks those rows Blank option.

### #REF! at the bottom of the list
The short formulas were filled past B3 rows. With 3 sizes and a count of 3 colors, rows 10 to 12 returned #REF!. Fill exactly B3 rows, or use the template, which blanks every row past B3.

### Two variants look the same
One color has a trailing space ("Black " vs "Black"), or the same color is typed twice. The template shows Extra space for a space before or after an option, and Duplicate for a variant name that appears twice. In LibreOffice, Black and black also count as a Duplicate, because the comparison ignores case.

### A size shows No price
The size has no price in column **B** of Options, or the price was typed as text such as "$18.00". Type the price as a plain number like 18.

### A size typed below row 14 never shows
It is past the yellow rows, so the formulas do not read it. The summary counts options typed past the yellow rows, and CHECK shows Check. The same goes for a color below row 24.

## Template

[Download the .xlsx](templates/tidy-tabs-size-color-combinations.xlsx)

The file has three tabs. **Options** takes sizes in **A5:A14**, a price for each size in **B5:B14** and colors in **D5:D24**. **Variants** counts them in **B1:B3** and lists No., Size, Color, Variant name and Price in **A:E** on rows 6 to 205, room for 10 sizes x 20 colors = 200 rows. Column **F** (On hand) is a yellow input for a stock count, and **G** shows Status: OK, Blank option, Extra space, Duplicate or No price.

The summary in **I5:J9** shows combinations, rows listed, rows not OK, options typed past the yellow rows, and a CHECK that says OK only when every combination is listed and nothing needs a look. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Any status other than OK is red. Both data tabs print on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
