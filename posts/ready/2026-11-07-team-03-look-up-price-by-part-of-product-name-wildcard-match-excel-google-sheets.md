---
title: How to Look Up a Price From Part of a Product Name in Excel and Google Sheets
labels: excel-formulas, small-business, lookup
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 16 workbooks and 2,336 comparisons with 0 mismatches. Checked: the sample, no search words, capitals, spaces around and doubled inside the term, a term that matches none, one and several products, the first and last of 100 rows, terms with * ? ~ and 50%, terms of 200, 201, 250 and 300 characters, numeric terms, blank product cells, a blank or text price, duplicate names, 100 products, and six random lists. Excel and Google Sheets were not used. Not run: wildcards in MATCH and COUNTIF in Google Sheets, numeric product names, non-breaking spaces, a non-US language setting.
search_description: Look up a price from part of a product name in Excel and Google Sheets with MATCH, INDEX and a * on each side. Free template flags several matches.
image_prompts: A small kraft price tag on twine, a stack of blank cardboard shipping boxes and a handheld sticker gun on a wooden packing table, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Blank kraft price tag on twine beside stacked cardboard boxes and a sticker gun
threads: Typing a few letters of a product name and getting its price back?\nPut a * on each side of the word inside MATCH, then INDEX returns the price. In the sample, tote finds the $22.00 canvas bag.\nA Several matches flag tells you when to type more.
---
Put a `*` on each side of the word you typed, so `=MATCH("*"&B9&"*",Products!$B$5:$B$104,0)` returns the first product whose name contains it, and `=INDEX(Products!$C$5:$C$104,D9)` returns its price. Typing tote finds "Canvas tote bag, 12 oz canvas" at $22.00 (fictional products).

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model (16 workbooks, 2,336 comparisons). Wildcards in MATCH and COUNTIF are documented by Microsoft for Excel, but neither Excel nor Google Sheets was opened for this post, and wildcards in MATCH and COUNTIF were not tested in Google Sheets.

![Search tab matching typed words to a product list and showing SKU and price](images/2026-11-07-team-03-look-up-price-by-part-of-product-name-wildcard-match-excel-google-sheets/look-up-price-by-part-of-product-name-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The formulas are the same in Excel and Google Sheets. **Products** holds the price list from row 5: SKU in **A**, product name in **B**, price in **C**. **Search** has the words you type in column **A**, from row 5.

1. On **Products**, list every product, one row each. Keep names as text.
2. On **Search**, type part of a name in **A5**, such as `tote`.
3. In **B5**, type `=IF(TRIM(A5)="","",IF(LEN(TRIM(A5))>200,"",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A5),"~","~~"),"*","~*"),"?","~?")))`. This trims the spaces at the ends and turns any `*`, `?` or `~` you typed into plain characters. A term over 200 characters stays blank, because MATCH cannot search a term over 255.
4. In **C5**, type `=IF(B5="","",COUNTIF(Products!$B$5:$B$104,"*"&B5&"*"))`. It counts the products that contain the word.
5. In **D5**, type `=IF(B5="","",IF(C5=0,"",MATCH("*"&B5&"*",Products!$B$5:$B$104,0)))`. It returns the list row of the first match.
6. In **E5**, type `=IF(ISNUMBER(D5),INDEX(Products!$B$5:$B$104,D5),"")`. In **F5**, use the same formula on column **A** and add `&""` after INDEX so a blank SKU shows as empty.
7. In **G5**, use the same formula on column **C**, and show "No price" when the price is not a number.
8. In **H5**, type `=IF(TRIM(A5)="","",IF(B5="","Too long",IF(C5=0,"No match",IF(C5=1,"One match","Several matches"))))`.
9. Fill **B5:H5** down to row 24.

The `*` stands for any run of characters, including none, so `"*"&B5&"*"` means "contains". Matching ignores capitals: SOY CANDLE finds "Soy candle".

MATCH with `0` as the last argument needs an exact result, and the wildcards make "exact" mean "contains". It stops at the first hit. That is why **C** counts the matches: when it shows 2 or more, the price you see is only the first product's price.

Help pages: Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) describe the wildcard characters. They describe Excel, not our test.

## Example

Sample data is fictional: 12 products and 9 search words.

| Word typed | Matches | First match is list row | First matching product | SKU | Price | Status |
|---|---|---|---|---|---|---|
| lavender | 2 | 1 | Soy candle, Lavender, 8 oz | CND-LAV8 | $18.00 | Several matches |
| SOY CANDLE | 3 | 1 | Soy candle, Lavender, 8 oz | CND-LAV8 | $18.00 | Several matches |
| mug (typed with spaces around it) | 2 | 7 | Enamel mug, Blue, 12 oz | MUG-BLU | $16.00 | Several matches |
| 12 oz | 3 | 7 | Enamel mug, Blue, 12 oz | MUG-BLU | $16.00 | Several matches |
| tote | 1 | 9 | Canvas tote bag, 12 oz canvas | TOT-CAN | $22.00 | One match |
| gift |  |  |  |  |  | No match |
| candle | 3 | 1 | Soy candle, Lavender, 8 oz | CND-LAV8 | $18.00 | Several matches |
| zzz |  |  |  |  |  | No match |
| 50% off |  |  |  |  |  | No match |

The summary shows 9 words typed, 1 one match, 5 several matches, 3 no match, 0 too long, and **CHECK** says OK.

Lavender finds two products, Lavender 8 oz and Lavender 4 oz. The sheet shows the first one at $18.00. To get the 4 oz candle, type `lavender, 8` or `lavender, 4`.

## Troubleshooting

### Status says Several matches
More than one product contains your word, and only the first row is returned. Type more of the name, such as `lavender, 4 oz`, until the status says One match.

### Status says No match
The word is not in any product name. Check the spelling, or look at the product list on **Products**. Spaces at the ends of the word are removed, but a double space inside a product name is not collapsed, so a word typed with a single space across that gap will not match.

### Status says Too long
The word is over 200 characters. MATCH cannot search a term longer than 255 characters, so the template stops at 200. Type a shorter piece of the name.

### A search for a symbol finds nothing, or finds too much
The `*`, `?` and `~` you type are treated as plain characters, because the formula puts a `~` in front of them. Searching `50%` works the same way. If a product name has one of these characters, type it the way it appears in the name.

### Numbers in the product name column behave oddly
Keep product names as text. In the LibreOffice test, a name stored as the number 8 was counted for the search 8, but Microsoft documents COUNTIF wildcards for text only. We did not test Excel, so format the name column as text if your names are plain numbers.

### It does not work in Google Sheets
Wildcards in MATCH and COUNTIF were not tested in Google Sheets for this post. Try one search such as `tote` before you rely on the sheet there.

## Template

[Download the .xlsx](templates/tidy-tabs-lookup-by-part-of-name.xlsx)

The file has three tabs. **Search** has 20 yellow input cells for the words you type, with the formulas in gray and a summary in **K4:K9**: words typed, one match, several matches, no match, too long, and a **CHECK** cell. **Products** has 100 yellow rows from row 5 with 12 fictional sample products. **How to use** has notes on each formula.

The template covers 100 products and 20 search words. More rows than that were not tested.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
