---
threads_url: https://www.threads.com/@tin_ylab/post/DfC7SUzFDmp
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-closest-value-to-number-in.html
title: How to Find the Closest Value to a Number in a List in Excel and Google Sheets
labels: excel-formulas, small-business, lookup
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with Python. Checked: ties, an unsorted size list, a blank cell in the list, decimals, a length below the smallest and above the largest size, an exact match, one size only, a duplicate size, a length or size typed as text, no sizes at all, and 100 order rows. Excel and Google Sheets were not used.
search_description: Find the closest value to a number in a list in Excel and Google Sheets with ABS, MIN, INDEX and MATCH. Free template with sizes, the next size up and a tie note.
image_prompts: Three unfinished wooden picture frames of different sizes leaning against a workshop wall beside a yellow tape measure and a pencil on a workbench, soft daylight, no screens, no readable text, no logos
image_alt: Three bare wooden frames in different sizes leaning on a wall beside a tape measure
threads: Need the standard size closest to a custom length?\nPut each size's distance from your number in a helper cell: ABS($B9-H$7). Then INDEX/MATCH on the MIN of that row returns the size. No array formula.\nTwo sizes equally close? The one listed first wins, and the template says so in a note column.
---
Put each size's distance from your number in a helper cell with `=ABS($B9-H$7)`, then return the closest size with `=INDEX($H$7:$S$7,MATCH(MIN($H9:$S9),$H9:$S9,0))`. Here H7:S7 is the size list, B9 is the number you are matching, and H9:S9 holds the distances.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Sizes sheet matching fifteen print orders to the nearest standard size with off by amounts, next size that fits and tie notes beside the difference columns](images/2026-10-28-team-03-find-closest-value-in-a-list-abs-min-index-match-excel-google-sheets/closest-match-size-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

Measure how far every size is from the number, then pick the smallest distance. The helper block keeps this a normal formula, with no Ctrl+Shift+Enter.

The layout below matches the template. Sizes go across **H7:S7**, order names go in column A, and each order's length in inches goes in column B, starting in row 9.

### In Excel

1. In **H7:S7**, type the standard sizes, such as 8, 10, 12, 14, 16, 18, 20, 24, 30. Leave unused cells empty.
2. In **A9** and **B9**, type an order name and its length, such as PR-2107 and 21. Add more orders below.
3. In **H9**, type `=ABS($B9-H$7)`. This is the distance between the length and the first size.
4. Select **H9:S9**, drag the fill handle right to column S, then fill the whole block down to your last order.
5. In **C9**, type `=INDEX($H$7:$S$7,MATCH(MIN($H9:$S9),$H9:$S9,0))`. This is the closest size.
6. In **D9**, type `=MIN($H9:$S9)`. This is how far off the closest size is.
7. In **E9**, type `=IF(COUNTIF($H$7:$S$7,">="&B9)=0,"None",SMALL($H$7:$S$7,COUNTIF($H$7:$S$7,"<"&B9)+1))`. This is the smallest size that is at least as long as the length, or None if no size is long enough. It works on an unsorted list.
8. In **F9**, type `=IF(COUNTIF($H9:$S9,D9)>1,"Tie, first listed used","")`.
9. Select **C9:F9** and drag the fill handle down.

Microsoft's pages: [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd). These pages describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 9 above, typed exactly as written.
2. Fill **H9:S9** right and down the same way, then fill **C9:F9** down.

Google's help page: [MATCH function](https://support.google.com/docs/answer/3093378). That page describes the function, not our test. None of these formulas needs AGGREGATE, which Google Sheets does not have.

### The version in the template

The template's C9 adds guards so a blank row stays blank and a length typed as text shows a message:

`=IF($B9="","",IF(ISNUMBER($B9),IF(COUNT($H9:$S9)=0,"No sizes",INDEX($H$7:$S$7,MATCH(MIN($H9:$S9),$H9:$S9,0))),"Check length"))`

Its helper cells are `=IF(AND(ISNUMBER($B9),ISNUMBER(H$7)),ROUND(ABS($B9-H$7),6),"")`. The ROUND to 6 places keeps float noise from hiding a tie.

The summary sits in **D1:D5**. **D5** says OK or CHECK.

## Example

Sample data is fictional: nine standard sizes (8, 10, 12, 14, 16, 18, 20, 24, 30 inches) and 15 made up print orders.

| Order | Length | Closest | Off by | Next size that fits | Note |
|---|---|---|---|---|---|
| PR-2101 | 9 | 8 | 1 | 10 | Tie, first listed used |
| PR-2106 | 18 | 18 | 0 | 18 | |
| PR-2107 | 21 | 20 | 1 | 24 | |
| PR-2110 | 27 | 24 | 3 | 30 | Tie, first listed used |
| PR-2112 | 31.5 | 30 | 1.5 | None | |
| PR-2114 | 35 | 30 | 5 | None | |

The summary for all 15 orders: **D1** is 15 orders, **D2** is 9 where the closest size is smaller than the length, **D3** is 2 longer than every size, **D4** is a largest gap of 5, and **D5** is OK.

A tie takes the size listed first from the left, not the smaller one. Three rows show the rules. Length 9 sits exactly between 8 and 10, so the tie note appears and 8 is used because it is listed first. Length 21 has a closest size of 20, but the next size that fits is 24. Length 35 is longer than every size, so it gets 30 as the closest and None as the next size up.

## Troubleshooting

### The row shows Check length and D5 says CHECK
The length is text, not a number, so it left-aligns. Retype it as a number. A size typed as text is skipped in the helper cells, and D5 also turns to CHECK.

### A tie note shows on a length that is not really tied
The same size is probably listed twice, such as 12 and 12. Two equal distances then count as a tie. Delete the duplicate. In the test, a repeated size gave the note and a CHECK in D5.

### Every row says No sizes
The size cells in H7:S7 are empty or hold text. Type the sizes as numbers. Blank cells in the middle of the list are fine and are skipped.

## Template

[Download the .xlsx](templates/tidy-tabs-closest-match.xlsx)

The file has two tabs. **Sizes** holds up to 12 yellow standard sizes in **H7:S7** and yellow orders in **A9:B108**, with 15 fictional rows filled in. Gray cells are formulas in **C** to **F**, **H** to **S** and **D1:D5**. **How to use** has short notes.

This finds the numerically nearest size above or below. It does not say what to quote or sell. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Sizes tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
