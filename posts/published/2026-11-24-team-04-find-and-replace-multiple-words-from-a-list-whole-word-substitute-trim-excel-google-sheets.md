---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-find-and-replace-multiple-words.html
title: How to Find and Replace Multiple Words at Once From a List in Excel and Google Sheets
labels: excel-formulas, text-functions, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 22 workbooks and 159,060 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Replace a whole list of words in every row with padded SUBSTITUTE helper columns and TRIM, matching whole words so Smoky never becomes Smalloky.
image_prompts: A row of finished soy candles in tins and glass jars on a home workshop bench, each with a blank kraft paper label, next to a roll of blank kraft labels and a pencil, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Row of candle tins and jars with blank kraft labels on a workbench
threads: Replace Sm with Small and Smoky turns into Smalloky.\nPad the text, double every space, and give each find-and-replace pair its own SUBSTITUTE helper column, then TRIM.\nThe doubled spaces matter: with single ones, the second of two neighbors gets missed.
---
Pad the text and double every space with `=SUBSTITUTE(" "&text&" "," ","  ")`, then give each find-and-replace pair its own helper column, `=SUBSTITUTE(prev," "&find&" "," "&replace&" ")` with the pair read from your list by INDEX, and clean up with `=TRIM(last_column)`. With the made-up pairs Lg to Large, Lav to Lavender and w/ to with, "Soy Candle Lg Lav w/ Wood Wick" becomes "Soy Candle Large Lavender with Wood Wick", and "Smoky" stays "Smoky" instead of turning into "Smalloky".

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (22 workbooks, 159,060 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Ten abbreviated candle product titles beside their cleaned versions with Changed? and pair counts](images/2026-11-24-team-04-find-and-replace-multiple-words-from-a-list-whole-word-substitute-trim-excel-google-sheets/replace-words-from-list-template.png) *Text tab with ten fictional product titles: "Soy Candle Lg Lav w/ Wood Wick" becomes "Soy Candle Large Lavender with Wood Wick" with 3 pairs applied, "Smoky" stays Smoky, and 9 of 10 rows change. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets.

1. On the **Replacements** tab, type one pair per row in rows 5 to 24: the word to find in **A** (one word, no spaces) and its replacement in **B**. Leave **B** blank to delete the word.
2. On the **Text** tab, paste your titles into **A5:A304**.
3. In **C5**, pad the text with `=IF(A5="","",SUBSTITUTE(" "&A5&" "," ","  "))`. Now every word, even the first and last, has spaces on both sides.
4. In **D5**, apply pair 1: `=IF(INDEX(Replacements!$A$5:$A$24,COLUMNS($D5:D5))="",C5,SUBSTITUTE(C5," "&INDEX(Replacements!$A$5:$A$24,COLUMNS($D5:D5))&" "," "&INDEX(Replacements!$B$5:$B$24,COLUMNS($D5:D5))&" "))`. COLUMNS counts up as you fill right, so each column reads the next pair. A blank pair row passes the text through.
5. Fill **D5** right to **W5**, one column per pair, 20 in all.
6. In **B5**, get the cleaned text with `=IF(A5="","",TRIM(W5))`. TRIM puts single spaces back.
7. In **X5**, flag changed rows with `=IF(A5="","",IF(EXACT(A5,B5),"","Yes"))`. In **Y5**, count the pairs that changed the row with `=IF(A5="","",SUMPRODUCT(--NOT(EXACT(C5:V5,D5:W5))))`.
8. Fill **B5:Y5** down to row 304. In the template, **C:W** are grouped and hidden.

**Why the spaces are doubled:** with single spaces, the first match uses up the space between two neighbors, so the second one is missed.

More than 20 pairs means adding more helper columns. See the [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) and [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) pages from Microsoft.

## Example

The sample titles and pairs are made up. Pairs on **Replacements**: Lg to Large, Sm to Small, Med to Medium, w/ to with, Lav to Lavender, Vnl to Vanilla, oz. to oz, pk to pack. All eight show OK.

| Original | Cleaned | Changed? | Pairs applied |
|---|---|---|---|
| Soy Candle Lg Lav w/ Wood Wick | Soy Candle Large Lavender with Wood Wick | Yes | 3 |
| Smoky Sm Candle 8 oz. | Smoky Small Candle 8 oz | Yes | 2 |
| Vnl Bean Soy Candle Med | Vanilla Bean Soy Candle Medium | Yes | 2 |
| Beeswax Tea Lights 12 pk | Beeswax Tea Lights 12 pack | Yes | 1 |
| Lav Soap Bar w/ Oatmeal | Lavender Soap Bar with Oatmeal | Yes | 2 |
| Sm Lav Tin Candle | Small Lavender Tin Candle | Yes | 2 |
| Gift Set Lg w/ 3 Sm Candles | Gift Set Large with 3 Small Candles | Yes | 3 |
| Tea Light Lav (Lg) | Tea Light Lavender (Lg) | Yes | 1 |
| Smoky Vnl Wax Melts 3 oz. | Smoky Vanilla Wax Melts 3 oz | Yes | 2 |
| Cedar Wood Wick Candle | Cedar Wood Wick Candle | | 0 |

The **Text** summary shows 10 rows with text, 9 changed and CHECK OK. The **Replacements** summary shows 8 pairs used, 0 Check, 0 cautions and CHECK OK.

The template against a plain nested SUBSTITUTE (pairs Lg, Sm, w/, Lav, oz.) and single-space padding, in LibreOffice Calc:

| Text | Template result | Plain nested SUBSTITUTE | Single-space padding (Lg only) |
|---|---|---|---|
| Smoky Sm Candle 8 oz. | Smoky Small Candle 8 oz | Smalloky Small Candle 8 oz | |
| Lg Lg Gift Set | Large Large Gift Set | Large Large Gift Set | Large Lg Gift Set |
| Lg, Sm and XL | Lg, Small and XL | Large, Small and XL | |
| Tea Light Lav (Lg) | Tea Light Lavender (Lg) | Tea Light Lavender (Large) | |
| sm lav tin | sm lav tin (unchanged) | sm lav tin | |
| "  Lg   Tin  " | Large Tin | | |

A number works as a find: with 8 to eight and oz. to ounce, "8 oz. jar" became "eight ounce jar" and "18 oz. jar" became "18 ounce jar".

## Troubleshooting

### "Smoky" became "Smalloky"
Plain SUBSTITUTE replaces letters inside words, so Sm to Small also hits the "Sm" in "Smoky". Search for " Sm " with spaces around it, as in steps 3 and 4.

### Only the first of two neighbors changed ("Large Lg Gift Set")
The padding uses single spaces, so the first match takes the space the second one needs. Double every space with `SUBSTITUTE(" "&A5&" "," ","  ")` before the pairs run.

### "Lg," or "(Lg)" or "sm" was not replaced
Whole words here are separated by spaces only, so a word next to a comma, parenthesis, hyphen or tab is not matched, and SUBSTITUTE is case-sensitive. Add pairs such as (Lg) to (Large) or sm to Small, or clean punctuation first.

### "Lg" ended up as "XL"
Pairs run in list order, and a later pair changed the replacement. With Lg to Large above Large to XL, "Lg Soy Candle" became "XL Soy Candle"; with the order swapped it became "Large Soy Candle". The Status shows "Caution: a later pair changes this replacement" on the earlier row.

## Template

[Download the .xlsx](templates/tidy-tabs-replace-words-from-list.xlsx)

The file has three tabs. **Replacements** holds 20 pairs with a Status that flags a blank find, a find with a space, the same find twice and the order trap. **Text** holds 300 titles, the hidden helper columns and a summary ending in CHECK. **How to use** explains each formula. Both data tabs print Landscape on US Letter. There is no REDUCE, LAMBDA or macro by design. To turn titles into URL handles instead, see [How to Turn Product Titles Into URL Handles (Slugs)](https://tidytabs.blogspot.com/2026/11/how-to-turn-product-titles-into-url.html).

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
