---
threads_url: https://www.threads.com/@tin_ylab/post/DfX8_AmgW7-
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-count-words-in-cell-in-excel-and.html
title: How to Count Words in a Cell in Excel and Google Sheets
labels: excel-formulas, small-business, text
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 26 workbooks and 1,152 comparisons with 0 mismatches. Checked: the sample, an empty cell, one word, extra spaces, a cell of only spaces, line breaks, a tab, a non-breaking space, numbers, a product without a description, a description without a product, exactly 25 and 24 words with a minimum of 25, minimums of 1, 0, text and blank, a decimal minimum, punctuation tokens, ties for shortest, accents and emoji, and six random sets of 3 to 100 rows. Not run: more than 100 descriptions, a non-US language setting. Non-breaking spaces were tested only in the word-count text. Excel and Google Sheets were not used.
search_description: Count words in a cell in Excel and Google Sheets with LEN, TRIM and SUBSTITUTE. Free template flags short product descriptions against a minimum.
image_prompts: An amber candle jar and a wrapped soap bar on a craft table with a few blank kraft product tags and a pencil, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Amber candle jar and wrapped soap bar beside blank kraft product tags
threads: An empty cell still counts as 1 word. Annoying.\nCount the spaces plus 1 with LEN, TRIM and SUBSTITUTE, and wrap it in an IF so blanks return 0.\nThe free template flags descriptions under a minimum you set, 25 words in the sample.
---
Count the words in B6 with `=LEN(TRIM(B6))-LEN(SUBSTITUTE(TRIM(B6)," ",""))+1`, which is the number of spaces between words plus 1. An empty cell still returns 1, so the template wraps this in an IF that returns 0.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. LEN, TRIM and SUBSTITUTE exist in current Excel and Google Sheets, but neither app was opened for this post.

![Product descriptions with word counts and Short or Missing flags beside a total and average](images/2026-11-05-team-03-count-words-in-cell-product-description-len-substitute-excel-google-sheets/product-description-word-count-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*


## Steps

The steps are the same in Excel and Google Sheets. Put one product description in each cell of column **B**, starting in row 6.

1. In **C6**, type `=LEN(TRIM(B6))-LEN(SUBSTITUTE(TRIM(B6)," ",""))+1`.
2. Select **C6** and fill down.
3. To return 0 for an empty cell, wrap it: `=IF(TRIM(B6)="",0,LEN(TRIM(B6))-LEN(SUBSTITUTE(TRIM(B6)," ",""))+1)`.
4. To compare with a minimum, type the number in **B2** and use `=IF(C6=0,"Missing",IF(C6<$B$2,"Short","OK"))`.

TRIM removes spaces at the start and end and squeezes repeated spaces between words down to one. LEN counts the characters left. SUBSTITUTE then deletes every space and LEN counts again. The difference is the number of spaces, and one more than that is the number of words.

Line breaks inside a cell are not spaces, so "Soft cotton" typed on two lines would count as one word. The template swaps each line break for a space first, using `SUBSTITUTE(B6,CHAR(10)," ")` inside the TRIM.

Help pages: Microsoft's [LEN function](https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb), [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) and [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: ten products from a small craft shop, with a minimum of 25 words.

| Product | Words | Status |
|---|---|---|
| Lavender soy candle | 28 | OK |
| Cedar soy candle | 13 | Short |
| Oatmeal soap bar | 28 | OK |
| Brass hoop earrings | 8 | Short |
| Letterpress birthday card | 33 | OK |
| Enamel camp mug | 9 | Short |
| Canvas tote bag | 31 | OK |
| Dotted notebook | 0 | Missing |
| Sticker pack | 11 | Short |
| Linen napkin set | 34 | OK |

The summary beside the table shows 195 total words, an average of 19.5 words per product, and a shortest description of 0 words (Dotted notebook, found with `=INDEX(A6:A105,MATCH(MIN(C6:C105),C6:C105,0))`). Five products are Short or Missing, and the **CHECK** cell says OK.

A minimum of 25 makes a description of exactly 25 words OK and one of 24 words Short.

## Troubleshooting

### An empty cell shows 1
The basic formula counts spaces plus 1, and an empty cell has no spaces. Use the IF in step 3. In the test, an empty description and a cell of only spaces both gave 0, and the status said Missing.

### Words joined by a tab or a non-breaking space count as one
Only the normal space character splits words. In the test, a tab between two words gave one word, and so did a non-breaking space, which some web pages put into copied text. Replace them with normal spaces. Non-breaking spaces were tested only in the word-count text, not in the minimum or product cells.

### A dash, 1/4 or 4.25 counts as a word
Any run of characters between spaces is a word here. In the test, `1/4` counted as one word and so did a dash standing alone. Type the dash without spaces around it, or accept the small difference.

### The count looks too low on a multi-line cell
Line breaks need the extra SUBSTITUTE with CHAR(10) from above. In the test, line breaks counted as spaces in the template.

### The status says Check
The minimum in **B2** is blank, text, or less than 1. Type a number of 1 or more in the yellow cell.

## Template

[Download the .xlsx](templates/tidy-tabs-word-count-descriptions.xlsx)

The file has two tabs. **Descriptions** has the minimum in **B2**, a settings check in **B3**, and 100 rows from row 6: yellow product names in **A** and descriptions in **B**, word counts in **C** and status in **D**. A summary in **F5:G11** shows products listed, total words, average words, shortest description, the product with the fewest words (the first one if there is a tie), the number of Short or Missing rows and a **CHECK** cell. **How to use** has notes on each formula.

**CHECK** says OK when the minimum is a number of 1 or more and every row has a word count. Not tested: more than 100 descriptions, a non-US language setting.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
