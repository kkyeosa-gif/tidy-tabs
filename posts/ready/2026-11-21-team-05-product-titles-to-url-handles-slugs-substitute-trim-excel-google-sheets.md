---
title: How to Turn Product Titles Into URL Handles (Slugs) in Excel and Google Sheets
labels: excel-formulas, text-functions, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 20 workbooks and 31,981 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Turn product titles into lowercase, hyphenated URL handles with LOWER, SUBSTITUTE and TRIM, flag leftover symbols and number duplicates as -2, -3.
image_prompts: A small tabletop product photo setup with a white paper backdrop and one amber jar candle, a row of unlabeled jam jars and a hand-painted mug waiting their turn, and a stack of blank kraft hang tags with a pencil at the table edge, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Candle on a white paper backdrop with jam jars, a painted mug and hang tags
threads: Three hyphens in a row is not a URL handle.\n=SUBSTITUTE(LOWER(TRIM(A5))," ","-") turns "Lavender & Honey Soy Candle - 8 oz" into lavender-&-honey-soy-candle---8-oz.\nTurn the symbols into spaces first, let TRIM collapse them, then swap in hyphens.
---
Make a handle by turning every unwanted character into a space, letting TRIM collapse the spaces, then turning the spaces into hyphens: `=SUBSTITUTE(TRIM(cleaned)," ","-")`, where cleaned is `LOWER(title)` wrapped in SUBSTITUTE calls that turn "&" into " and ", drop apostrophes and turn characters like / , . ( ) ! ? into spaces. In the made-up product list, "Lavender & Honey Soy Candle - 8 oz" becomes lavender-and-honey-soy-candle-8-oz, and the download counts any character left that is not a-z, 0-9 or a hyphen and numbers repeat handles as -2 and -3.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (20 workbooks, 31,981 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (LOWER, SUBSTITUTE, TRIM, MID, FIND, ROW, INDIRECT, SUMPRODUCT, COUNTIF) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Twelve product titles beside cleaned lowercase handles, numbered repeats and a quick formula comparison](images/2026-11-21-team-05-product-titles-to-url-handles-slugs-substitute-trim-excel-google-sheets/product-title-to-url-handle-template.png) *Titles tab with 12 fictional product titles, where Lavender & Honey Soy Candle - 8 oz becomes lavender-and-honey-soy-candle-8-oz, repeats get -2 and -3, Crème Brûlée Candle shows Fix: special letters, and the quick formula column keeps the &, colons and extra hyphens. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Titles** tab.

1. Type or paste your product titles in **A5:A204**.
2. In **B5**, type the handle formula below. LOWER makes the title lowercase. The SUBSTITUTE calls turn & into " and ", remove straight and curly apostrophes, and turn each of - / , . : ; ! ? ( ) " # % + _ and the pipe character into a space. TRIM removes spaces at the ends and collapses double spaces, and the outer SUBSTITUTE turns the spaces into hyphens.

   `=IF(TRIM(A5)="","",SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(LOWER(A5),"&"," and "),"'",""),"’",""),"-"," "),"/"," "),","," "),"."," "),":"," "),";"," "),"!"," "),"?"," "),"("," "),")"," "),""""," "),"#"," "),"%"," "),"+"," "),"_"," "),"|"," "))," ","-"))`

3. In **C5**, count what is left over: `=IF(B5="","",SUMPRODUCT(--ISERROR(FIND(MID(B5,ROW(INDIRECT("1:"&LEN(B5))),1),"abcdefghijklmnopqrstuvwxyz0123456789-"))))`. MID pulls out each character, and FIND looks for it in the allowed list. Anything not found (é, an en dash character, $, @) adds 1.
4. In **D5**, type `=IF(B5="","",COUNTIF($B$5:$B$204,B5))` to see how many titles share this handle.
5. In **E5**, type `=IF(B5="","",IF(COUNTIF($B$5:B5,B5)>1,B5&"-"&COUNTIF($B$5:B5,B5),B5))`. The range grows as you fill down, so the first copy keeps the plain handle and later copies get -2, -3. The post on [numbering items within each order](https://tidytabs.blogspot.com/2026/10/how-to-number-items-within-each-order.html) explains this growing range.
6. In **F5**, type `=IF(E5="","",LEN(E5))` for the handle length, and in **H5**, `=IF(E5="","",COUNTIF($E$5:$E$204,E5))` to catch a numbered handle that still repeats.
7. In **G5**, type `=IF(TRIM(A5)="","",IF(B5="","Check: no letters or digits",IF(C5>0,"Fix: special letters",IF(H5>1,"Check: numbered handle repeats",IF(D5>1,"Duplicate (numbered)","OK")))))`.
8. Fill **B5:H5** down to row 204.
9. Copy column E and paste it as values only into your import file. Length limits and allowed characters vary, so check your shop's import help for its own URL handle rules.

Microsoft's [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) and [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) pages and Google's [SUBSTITUTE help page](https://support.google.com/docs/answer/3094215) cover the two main functions.

## Example

The product titles are fictional. Quick formula is the shortcut `=SUBSTITUTE(LOWER(TRIM(A5))," ","-")` from column I, shown for comparison.

| Product title | Unique handle | Leftover | Length | Status | Quick formula |
|---|---|---|---|---|---|
| Lavender & Honey Soy Candle - 8 oz | lavender-and-honey-soy-candle-8-oz | 0 | 34 | Duplicate (numbered) | lavender-&-honey-soy-candle---8-oz |
| Grandma's Apple Butter (16 oz.) | grandmas-apple-butter-16-oz | 0 | 27 | OK | grandma's-apple-butter-(16-oz.) |
| Hand-Painted  Mug, Blue/White (extra spaces) | hand-painted-mug-blue-white | 0 | 27 | OK | hand-painted-mug,-blue/white |
| Holiday Gift Box: 3 Jams! | holiday-gift-box-3-jams | 0 | 23 | OK | holiday-gift-box:-3-jams! |
| 100% Cotton Tote #2 | 100-cotton-tote-2 | 0 | 17 | OK | 100%-cotton-tote-#2 |
| Grandma’s Fig Jam (curly apostrophe) | grandmas-fig-jam | 0 | 16 | OK | grandma’s-fig-jam |
| "Best Seller" Tea Towel? | best-seller-tea-towel | 0 | 21 | OK | "best-seller"-tea-towel? |
| Crème Brûlée Candle | crème-brûlée-candle | 3 | 19 | Fix: special letters | crème-brûlée-candle |
| Lavender and Honey Soy Candle 8 oz | lavender-and-honey-soy-candle-8-oz-2 | 0 | 36 | Duplicate (numbered) | lavender-and-honey-soy-candle-8-oz |
| Lavender and Honey Soy Candle 8 oz | lavender-and-honey-soy-candle-8-oz-3 | 0 | 36 | Duplicate (numbered) | lavender-and-honey-soy-candle-8-oz |

Two more sample rows are not in the table. "Pumpkin_Spice+Clove", a pipe character, then "Wax Melt" becomes pumpkin-spice-clove-wax-melt. The candle title typed with an en dash character instead of a hyphen keeps that character in its handle: 1 leftover, 36 characters, Fix: special letters.

D shows 3 on the three lavender-and-honey rows, because "&" and "and" end up the same. The summary in **K4:L11** shows 12 titles listed, 7 OK, 3 Duplicate (numbered), 2 Fix: special letters, 0 rows showing Check, a longest handle of 36 characters and 0 titles typed below row 204. CHECK shows Check on purpose, because of the two Fix rows.

## Troubleshooting

### Handles come out with & or triple hyphens
That is the quick formula. It only swaps spaces, so " - " turns into three hyphens and & stays in: lavender-&-honey-soy-candle---8-oz. Use the full B5 formula, which turns symbols into spaces before TRIM collapses them.

### A row shows Fix: special letters
The handle still has a character outside a-z, 0-9 and the hyphen: an accented letter like é, an en dash character, $, @ or a non-breaking space pasted from a web page (TRIM does not remove it). The sheet flags these and does not convert them, so retype them in the title, for example Creme Brulee Candle.

### Two products end up with the same handle
Column E numbers the later copies -2 and -3. If a numbered handle matches another title's own handle, such as two "Tote" rows and a "Tote 2" row all giving tote-2, Status shows Check: numbered handle repeats. Change one of the titles.

### The handle is empty
A title made only of symbols, such as !!!, has no letters or digits left, so Status shows Check: no letters or digits. Add a real word to the title.

## Template

[Download the .xlsx](templates/tidy-tabs-url-handle-from-title.xlsx)

The file has two tabs. **Titles** holds the 12 sample titles in rows 5 to 204 with Product title (input in **A**), then Handle, Leftover characters, Times this handle appears, Unique handle (adds -2, -3), Length, Status, Unique handle repeats and the Quick formula (wrong, comparison) in **B:I**. The summary in **K4:L11** counts each status, the longest handle, entries typed below row 204 and the CHECK. **How to use** explains each formula and lists the replaced characters.

Yellow cells are inputs and gray cells are formulas. Duplicate (numbered) turns orange, and every Fix and Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
