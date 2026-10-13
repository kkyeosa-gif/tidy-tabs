---
title: How to Count Characters and Tags in a Product Listing in Excel and Google Sheets
labels: excel-formulas, small-business, text-functions
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and 10 listings were checked against Python len and str.count, plus a title exactly at the limit, an empty tags cell, a trailing comma, a double comma, a trailing space and a cleared sheet. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Help.
search_description: Count characters with =LEN(A2) and comma separated tags with =LEN(B2)-LEN(SUBSTITUTE(B2,",",""))+1. A 46 character title, 4 tags. Free .xlsx.
image_prompts: Three small soy candle jars with blank kraft hang tags tied on with twine, a spool of twine and a wooden clothespin on a worn craft table, soft window light, no screens, no readable text, no logos
image_alt: Candle jars with blank kraft tags tied on twine beside a spool on a craft table
threads: Not sure if that listing title is too long?\nCount it with =LEN(A2). Count comma separated tags with =LEN(B2)-LEN(SUBSTITUTE(B2,",",""))+1.\nAn empty tags cell comes back as 1, not 0, so the template wraps it in an IF.
---
Count characters with `=LEN(A2)` and count comma separated tags in one cell with `=LEN(B2)-LEN(SUBSTITUTE(B2,",",""))+1`. The title "Handmade Lavender Soy Candle 8 oz Gift for Her" is 46 characters long.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Product listings with title character counts, tag counts and OK, Too long or Too many flags next to two limit cells](images/2026-10-14-team-03-count-listing-characters-tags-len-substitute-excel-google-sheets/listing-length-check-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: LEN counts the characters in a cell. To count tags, count the commas and add 1, because 4 tags are separated by 3 commas. SUBSTITUTE deletes the commas so LEN can measure what is left.

1. Type these headers in row 1: Title, Tags, Characters, Tags counted, Title check, Tag check.
2. Put each listing title in column **A** and its tags in column **B**, separated by commas.
3. In **C2**, type `=IF(A2="","",LEN(A2))` and fill it down. Spaces count as characters.
4. In **D2**, type `=IF(AND(A2="",B2=""),"",IF(TRIM(B2)="",0,LEN(B2)-LEN(SUBSTITUTE(B2,",",""))+1-(RIGHT(TRIM(B2),1)=",")))` and fill it down.
5. In **I2**, type your title limit. In **I3**, type your tag limit. Copy both numbers from your marketplace's own help page, because limits change.
6. In **E2**, type `=IF(A2="","",IF(C2>$I$2,"Too long","OK"))` and fill it down.
7. In **F2**, type `=IF(AND(A2="",B2=""),"",IF(D2>$I$3,"Too many","OK"))` and fill it down.

The short tag formula is the one in the first paragraph. Step 4 wraps it in two guards. The IF makes an empty tags cell return 0 instead of 1, and the last term subtracts 1 when the cell ends in a comma.

A title exactly at the limit passes, because E2 only flags counts above the limit.

### In Excel

The formulas work as typed. Microsoft describes the functions on its [LEN function](https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb) page and its [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) page.

### In Google Sheets

The same formulas should work as typed. Google describes the function on its [SUBSTITUTE help page](https://support.google.com/docs/answer/3094215). This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. The limits in I2 and I3 are made-up numbers (60 characters and 8 tags) so that the flags show up. They are not any marketplace's rules.

| Title | Characters | Tags counted | Title check | Tag check |
|---|---|---|---|---|
| Handmade Lavender Soy Candle 8 oz Gift for Her | 46 | 4 | OK | OK |
| Cedar and Sage Soy Candle Small Batch 8 oz Hand Poured Gift | 59 | 8 | OK | OK |
| Oatmeal Goat Milk Soap Bar Handmade | 35 | 9 | OK | Too many |
| Brass Hoop Earrings Minimalist Gold Tone Everyday Jewelry for Women Gift Idea | 77 | 4 | Too long | OK |
| Letterpress Birthday Card Blank Inside | 38 | 3 | OK | OK |
| Charcoal Soap Bar Cold Process Vegan | 36 | 0 | OK | OK |
| Eucalyptus Mint Bath Salts 12 oz Jar | 36 | 3 | OK | OK |
| Beeswax Taper Candles Pair Natural Unscented | 44 | 5 | OK | OK |
| Linen Wick Soy Candle Tin Travel Size 4 oz | 42 | 8 | OK | OK |
| Kraft Gift Box Set Soap and Candle Combo Birthday | 49 | 5 | OK | OK |

The summary block shows 10 listings checked, 1 title too long and 1 listing with too many tags. Cell **I7** shows OK when the character counts add up a second way, with `=IF(SUM(C2:C30)=SUMPRODUCT(LEN(A2:A30)),"OK","CHECK")`.

When the title limit in I2 was changed to 46 in the recalculated file, the 46 character title stayed OK and 3 titles were flagged Too long.

The Charcoal Soap row has no tags typed, so it shows 0. The tags in the Eucalyptus row end with a comma, so it shows 3, not 4.

## Troubleshooting

### An empty tags cell shows 1 tag
The plain formula counts commas and adds 1, so zero commas always gives 1. In the recalculated file, `=LEN(B7)-LEN(SUBSTITUTE(B7,",",""))+1` returned 1 for the empty cell. Wrap it as `=IF(B7="",0,...)` and it returns 0.

### The count is one too high
The cell ends with a comma, such as "bath salts, eucalyptus, spa gift,". The plain formula returned 4 for those 3 tags. The template formula in **D8** subtracts 1 when the last character is a comma and returned 3.

### Two commas in a row count as an extra tag
"a,,b" counts as 3 tags in the recalculated file, because an empty tag between two commas is still a comma. Delete the extra comma. Spaces after commas do not change the count.

### The character count is higher than I expected
LEN counts every space, including one at the start or end of the cell. A test title with a trailing space gave 26, and `=LEN(TRIM(A20))` gave 25. Use TRIM when you paste titles in. If a marketplace's own counter shows a different number, trust its counter and check its help page for how it counts.

### Every row says OK even though a title is long
Check **I2** and **I3**. The flags compare against those two typed numbers, and a limit set too high flags nothing. If I7 shows CHECK, a formula in column C was typed over with a number.

## Template

[Download the .xlsx](templates/tidy-tabs-listing-length-check.xlsx)

The file has two tabs. **Listings** holds 10 sample listings in A2:B11, formulas ready for rows 2 to 30, two yellow limit cells in I2:I3 and the summary in I4:I7. **How to use** has short fill-in notes.

The titles and tags are fictional, and the limits 60 and 8 are placeholders you replace. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
