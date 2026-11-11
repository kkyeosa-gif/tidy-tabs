---
threads_url: https://www.threads.com/@tin_ylab/post/Dfm3cRMGu2c
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-shorten-product-title-to.html
title: How to Shorten a Product Title to a Character Limit Without Cutting a Word in Excel and Google Sheets
labels: excel-formulas, small-business, listings
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 19 workbooks and 6,973 comparisons with 0 mismatches left. Checked: limits 140, 139, 141, 20, 1, 2, 3, 10, 50, 200 and 300, titles just under, at and over each limit, a space at position limit plus 1, doubled spaces, one long word of 3 lengths, leading and trailing spaces, blank rows, a title with a ~, numbers as titles, a limit that is text, blank, 2.5, 0 or -5, and random titles. Excel and Google Sheets were not used. Not run: non-breaking spaces, more rows than the template holds, a non-US language setting.
search_description: Cut a long product title at the last space inside the first 141 characters with LEFT, FIND and SUBSTITUTE, so no word is split. Free Excel template.
image_prompts: A kraft gift box with its lid off beside a small printed card and a yellow tape measure on a craft studio table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Open kraft gift box with a printed card and a tape measure on a table
threads: Listing title 10 characters too long?\nCut it at the last space inside the first 141 characters, so the last word stays whole.\nIn the sample, a 150 character title comes out at 138 and a 145 character one at 140. The limit sits in one cell, B1 (made-up titles).
---
Keep whole words under a limit by cutting the title at the last space inside the first 141 characters, with `=LEFT(TRIM(A4),FIND("~",SUBSTITUTE(LEFT(TRIM(A4),$B$1+1)," ","~",LEN(LEFT(TRIM(A4),$B$1+1))-LEN(SUBSTITUTE(LEFT(TRIM(A4),$B$1+1)," ",""))))-1)`, where **B1** holds the limit. A 150 character sample title with a limit of 140 comes out at 138 characters, ending on a whole word (made-up titles).

The limit is a setting, not a rule from us. Check your marketplace for its real number and type it in B1.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (19 workbooks, 6,973 comparisons). The formulas use only IF, TRIM, LEN, LEFT, FIND, SUBSTITUTE, IFERROR, ISNUMBER and COUNTIF, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Twelve sample product titles with lengths, shortened versions and a status column](images/2026-11-11-team-01-shorten-title-to-character-limit-whole-word-left-find-substitute-excel-google-sheets/shorten-title-whole-word-template.png) *The Titles tab with a limit of 140: titles over the limit are cut at the last space, one long word is flagged. PDF export from LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Titles** tab and a **How to use** tab.

1. In **B1**, type the character limit, for example `140`.
2. Paste your titles into **A4** and down. The template holds rows 4 to 203.
3. In **B4**, type `=LEN(TRIM(A4))`. This is the length after outer spaces are removed and doubled spaces become one.
4. In **C4**, type `=IF(B4<=$B$1,TRIM(A4),` followed by the long formula from the top of this page and a closing `)`. A title that already fits is copied as is.
5. In **D4**, type `=LEN(C4)` to see the length after the cut.
6. Fill **B4:D4** down to the last title.

How it works: `LEFT(TRIM(A4),$B$1+1)` keeps one character more than the limit. `SUBSTITUTE(...," ","~",n)` swaps only the last space in that piece for a `~`, where n is the number of spaces in it. `FIND("~",...)` locates that marker, and the outer `LEFT` keeps everything before it.

The `+1` is the point. If the character after position 140 is a space, the word ending at 140 is already whole and stays. If it is a letter, that word would be split, so it is dropped.

The template wraps this formula so blank rows stay blank, a title with no space in its first 141 characters is flagged, and a bad limit shows Check.

Help pages: Microsoft's [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332), [FIND function](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628) and [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9), and Google's [SUBSTITUTE](https://support.google.com/docs/answer/3094215). They describe the functions, not our test.

## Example

Sample data is fictional: 12 made-up titles with a limit of 140.

| Title (start) | Length | After | Status |
|---|---|---|---|
| Handmade Soy Candle Lavender Vanilla 8 oz Gift Set... | 150 | 138 | Shortened |
| Oatmeal Goat Milk Soap Bar | 26 | 26 | OK |
| Personalized Cutting Board ... Family Name ... Custom Charcuterie | 140 | 140 | OK |
| Personalized Cutting Board ... Family Names ... Custom Charcuterie | 141 | 129 | Shortened |
| Personalized Cutting Board ... Family Name ... Custom Charcuterie Tray | 145 | 140 | Shortened |
| Brass Hoop Earrings Small Gold Huggie ... Set | 158 | 137 | Shortened |
| Hand  Poured  Soy Candle Cedar Pine ... (doubled spaces) | 146 | 138 | Shortened |
| HandmadeSoyCandleLavender... (no spaces) | 149 | 140 | One long word |
| "  Leather Key Fob Keychain Personalized Initial  " | 45 | 45 | OK |

Rows with a title that already fits are not all listed. The summary block shows 12 titles listed, 6 within the limit, 5 shortened at a space, 1 one long word, 0 rows showing Check, a longest length after of 140, and a CHECK cell that says OK.

Compare the 140, 141 and 145 character rows. At 140 nothing changes. At 141 the last word does not fit, so "Charcuterie" is dropped and the title ends at 129. At 145 a space sits at position 141, so the word ending at 140 stays whole and the title ends at 140.

## Troubleshooting

### A word is missing that seems to fit
The word ends after the limit. With a limit of 140, a word that ends at character 141 is dropped even though only one character is over. That is the rule you asked for: whole words only.

### The status says One long word
There is no space in the first limit plus 1 characters, so there is nowhere to cut at a space. The template keeps the first 140 characters and says so. Plain `SUBSTITUTE` with an instance of 0 would return an error here, which is why the template wraps the formula in `IFERROR`. Shorten that title by hand.

### Every row says Check
The limit in B1 is blank, text, 0, negative or has a decimal. Type a whole number of 1 or more. A single row says Check when its title contains a `~`, because the formula uses `~` as a marker. Remove it or change the marker to a character your titles never use.

### The length is off by a few characters
`TRIM` removes outer spaces and squeezes doubled spaces inside the title before anything is counted. The doubled-space sample row is 146 characters raw and 138 after the cut. Non-breaking spaces from a web page copy are not removed by `TRIM` and were not tested.

### The result still breaks your marketplace's rules
This file only counts characters. It knows nothing about banned words, symbols or how the site counts them.

## Template

[Download the .xlsx](templates/tidy-tabs-shorten-title-whole-word.xlsx)

The file has two tabs. **Titles** has the limit in **B1**, titles in **A4:A203**, the formula columns **B:E**, and a summary with a CHECK cell in **G3:H9**. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
