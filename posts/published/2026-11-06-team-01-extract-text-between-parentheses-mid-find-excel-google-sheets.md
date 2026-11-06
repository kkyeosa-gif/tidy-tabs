---
threads_url: https://www.threads.com/@tin_ylab/post/DfZ_cTWkRsb
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-extract-text-between-parentheses.html
title: How to Extract Text Between Parentheses in Excel and Google Sheets
labels: excel-formulas, small-business, text
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 12 workbooks and 6,072 comparisons with 0 mismatches, and no error value appeared in any recalculated workbook. Checked: the sample, no titles, (), ( ), (,), (a,), (,b), several commas, spaces inside and around the parentheses, a parenthesis first in the title, a missing closing ), a ) before the (, a nested pair, three pairs, no parentheses, spaces only, accents and Japanese text, a tab inside, quotes, a number in the cell, a 300 character title, 100 rows, a blank row between titles, and six random sets of up to 100 titles. Excel and Google Sheets were not used. Not run: titles over 300 characters, non-breaking spaces.
search_description: Pull the text inside parentheses from a product title with MID, FIND and TRIM in Excel and Google Sheets. Free template also splits at the first comma.
image_prompts: A row of three small glass jars of scented candles with plain kraft tags, a pair of scissors and a spool of twine on a wooden worktable, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Small candle jars with blank kraft tags beside scissors and twine on a wooden worktable
threads: Need just the part of a title that sits inside the parentheses?\n=TRIM(MID(A5,FIND("(",A5)+1,FIND(")",A5,FIND("(",A5))-FIND("(",A5)-1)) turns Soy candle (Lavender, 8 oz) into Lavender, 8 oz.\nThe free template also splits at the first comma and flags titles with no closing parenthesis.
---
Pull the text inside the first pair of parentheses with `=TRIM(MID(A5,FIND("(",A5)+1,FIND(")",A5,FIND("(",A5))-FIND("(",A5)-1))`. For "Soy candle (Lavender, 8 oz)" it returns Lavender, 8 oz.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every result below was recalculated and compared with a Python model. MID, FIND and TRIM exist in current Excel and Google Sheets, but neither app was opened for this post.

![Listing titles with the part inside parentheses pulled into its own column](images/2026-11-06-team-01-extract-text-between-parentheses-mid-find-excel-google-sheets/text-between-parentheses-listing-titles-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. On **Listings**, the header is row 4 and the titles go in **A5:A104**, one per row. Yellow cells are inputs and gray cells are formulas.

1. Paste your listing titles into **A5** and down. Each title should hold a variation in parentheses, such as `Soy candle (Lavender, 8 oz)`.
2. In **B5**, type the formula below. It returns what sits between the first ( and the first ) after it.
3. In **D5**, type the status formula. It says OK, No (, No closing ) or Empty ().
4. In **C5**, type the title-without-the-variation formula.
5. Select **B5:D5** and fill down to row 104.

**B5**, inside the parentheses:

`=IF(A5="","",IF(ISNUMBER(FIND("(",A5)),IF(ISNUMBER(FIND(")",A5,FIND("(",A5))),TRIM(MID(A5,FIND("(",A5)+1,FIND(")",A5,FIND("(",A5))-FIND("(",A5)-1)),""),""))`

**D5**, status:

`=IF(A5="","",IF(ISNUMBER(FIND("(",A5)),IF(ISNUMBER(FIND(")",A5,FIND("(",A5))),IF(B5="","Empty ()","OK"),"No closing )"),"No ("))`

**C5**, title without them:

`=IF(A5="","",IF(OR(D5="OK",D5="Empty ()"),TRIM(LEFT(A5,FIND("(",A5)-1)&" "&MID(A5,FIND(")",A5,FIND("(",A5))+1,500)),TRIM(A5)))`

How the core works: `FIND("(",A5)` gives the position of the opening parenthesis. `FIND(")",A5,FIND("(",A5))` finds the first closing one at or after it. MID starts one character past the opening and takes the gap between the two positions, minus one. TRIM removes extra spaces at the ends.

The `IF(ISNUMBER(FIND(...)))` wrappers keep a title with no parentheses from showing `#VALUE!`. FIND returns an error when it finds nothing, and ISNUMBER turns that into a plain test.

To split a variation like `Lavender, 8 oz` in two, use **E5** `=IF(B5="","",TRIM(LEFT(B5,FIND(",",B5&",")-1)))` for the part before the first comma and **F5** `=IF(ISNUMBER(FIND(",",B5)),TRIM(MID(B5,FIND(",",B5)+1,500)),"")` for the rest. Adding `&","` inside E5 means a variation with no comma, such as `Blue`, still works.

Help pages: Microsoft's [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028) and the [text functions reference](https://support.microsoft.com/en-us/excel/text-functions-reference) that lists FIND, and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 12 listing titles in rows 5 to 16.

| Title | Inside the parentheses | Title without them | Status | Before the first comma | After the first comma |
|---|---|---|---|---|---|
| Soy candle (Lavender, 8 oz) | Lavender, 8 oz | Soy candle | OK | Lavender | 8 oz |
| Enamel mug (Blue) | Blue | Enamel mug | OK | Blue | blank |
| Linen napkin set (Natural, Set of 4) | Natural, Set of 4 | Linen napkin set | OK | Natural | Set of 4 |
| Letterpress birthday card | blank | Letterpress birthday card | No ( | blank | blank |
| Tote bag (Canvas, 12 oz) Medium | Canvas, 12 oz | Tote bag Medium | OK | Canvas | 12 oz |
| Wool dryer balls (3 pack) (Gray) | 3 pack | Wool dryer balls (Gray) | OK | 3 pack | blank |

The summary in **H4:I9** shows 12 titles listed, 11 with a pair, 1 with no parentheses, 0 missing the closing ), 0 with empty parentheses, and **I9** says OK.

Row 15 shows the "first pair only" rule. The second pair, (Gray), stays in the title and is not read. Row 13 has no parentheses, so the variation is blank and the title comes through unchanged.

## Troubleshooting

### A title has no variation
The title has no ( at all, so **B** stays blank and **D** says No (. In the sample, Letterpress birthday card is the one such row, and **C** returns the title as typed.

### The closing parenthesis is missing
A title like `Mug (Blue` gives a blank **B**, the status No closing ), and **C** shows the trimmed title. Fix the title, then the row turns OK. **I7** counts these rows.

### Empty parentheses
`()` and `( )` both give the status Empty (). **B** is blank and **C** drops the empty pair. **I8** counts them.

### A nested pair is cut short
Only the first ( and the first ) after it are used. In the test, `mug (a (b) c) d` returned `a (b`. Nothing beyond that was checked, so review titles with a pair inside a pair by hand.

### The title has more than one pair
Only the first pair is read. In the test, `Wool dryer balls (3 pack) (Gray)` gave `3 pack`. Move the pair you want to the front, or add a second formula that starts after the first ).

### A number sits in column A
A number such as 123 is treated as text, so the status says No (. TRIM only removes ordinary spaces. A tab inside a variation stays in the result.

### The CHECK cell says Check
**I9** says OK when no title is missing its closing ) and none has empty parentheses. Fix those rows first. A title with no parentheses at all is allowed and does not change the check.

## Template

[Download the .xlsx](templates/tidy-tabs-text-between-parentheses.xlsx)

The file has two tabs. **Listings** has 100 rows from row 5 with 12 fictional sample titles, yellow input cells in column A, formulas in columns B to F, and a summary in **H4:I9**: titles listed, with a pair, with no parentheses, missing the closing ), with empty parentheses, and a **CHECK** cell. **How to use** has notes on each formula.

The template has 100 title rows. A longer list was not tested. The formulas were not opened in Excel or Google Sheets. To open the file in Google Sheets, go to **File > Import > Upload** and select it.
