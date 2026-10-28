---
title: How to Pick Random Giveaway Winners From a List in Excel and Google Sheets
labels: excel-formulas, small-business, random-draw
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there 200 times and each draw was checked with Python: draw order 1 to 30 with no repeats, 3 winners, every winner in the list. Also checked: 1 and 10 winners, fewer entries than winners, no entries, blank rows, duplicate names and a bad winners number. The Copy and Paste special freeze step was not run in any app. Excel and Google Sheets were not used.
search_description: Pick random giveaway winners from a list in Excel and Google Sheets with RAND, RANK and INDEX. Free template with 30 sample entries, a winners count you set, and a check for repeats.
image_prompts: A small glass bowl of folded paper slips beside a wooden spoon and a handful of loose wooden raffle tokens on a kitchen table, soft daylight, no screens, no readable text, no logos
image_alt: Glass bowl of folded paper slips beside wooden tokens and a spoon on a table
threads: 30 entries, 3 winners, no repeats.\nGive each row =RAND(), rank them with RANK plus COUNTIF to break ties, then INDEX/MATCH the names for draw order 1, 2, 3.\nRAND redraws on every recalculation, so copy and paste the winners as values once you have picked.
---
Give every entry a random number with `=IF(A5="","",RAND())`, then rank them with `=IF(A5="","",RANK(C5,$C$5:$C$34)+COUNTIF($C$5:C5,C5)-1)` so no two entries share a place. The entries ranked 1, 2 and 3 are your winners.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and checked. The Excel and Google Sheets steps use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Three drawn winners with cities beside a count of 30 entries and an OK check](images/2026-10-29-team-03-pick-random-giveaway-winners-no-repeats-rand-rank-excel-google-sheets/giveaway-random-winners-sheet.png) *caption: One random draw from the template with fictional sample data, rendered from LibreOffice Calc 24.2.7.2*

## Steps

RAND is volatile. It picks new numbers every time the sheet recalculates, which means every time you edit a cell or open the file. The winners change with it. Freeze the draw (step 8) before you share the result.

The layout below matches the template. Names go in **A5:A34**, cities in **B5:B34**, and the number of winners in a separate Winners tab.

### In Excel

1. In **A5:A34**, type or paste one entry per row. Put the city in **B5:B34**. Leave unused rows empty.
2. In **C5**, type `=IF(A5="","",RAND())`. This is the random number for that entry.
3. In **D5**, type `=IF(A5="","",RANK(C5,$C$5:$C$34)+COUNTIF($C$5:C5,C5)-1)`. RANK puts the biggest random number first. The COUNTIF part breaks a tie, so every filled row gets its own order number.
4. Select **C5:D5** and drag the fill handle down to row 34.
5. On a second sheet named Winners, type the number of winners in **B1**, such as 3.
6. In **A5**, type `=IF(ROWS(A$5:A5)<=B1,ROWS(A$5:A5),"")`. This counts 1, 2, 3 up to the number of winners.
7. In **B5**, type `=IF(A5="","",INDEX(Entries!$A$5:$A$34,MATCH(A5,Entries!$D$5:$D$34,0))&"")`. This returns the name whose draw order matches. Copy the same formula to **C5** but point INDEX at `Entries!$B$5:$B$34` to get the city, then fill **A5:C5** down 10 rows.
8. To freeze the draw, select **Entries C5:D34**, copy, then use **Paste special > Values only** on the same cells. The numbers stay put and the winners stop changing.

Microsoft's pages: [RANK function](https://support.microsoft.com/en-us/office/rank-function-6a2fc49d-1831-4a03-9d8c-c279cf99f723). It describes the function, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 7 above, typed exactly as written. Name the second tab Winners and the first Entries.
2. To freeze the draw, select **C5:D34**, copy, then choose **Edit > Paste special > Values only**.

Google's help page: [RAND](https://support.google.com/docs/answer/3093438). It describes the function, not our test. Whether RAND recalculates the same way in Google Sheets was not tested here, so freeze the draw there too.

### What the freeze step did and did not get

Copy and Paste special as values is the standard way to turn formulas into fixed numbers. The team did not run it in any app for this post. The template's How to use sheet describes it, and that is all. The checks below used live RAND formulas that were never frozen.

## Example

Sample data is fictional: 30 first names and cities, and 3 winners. RAND has no fixed answer, so the sample has no fixed winners. In LibreOffice, 200 recalculations gave 196 different sets of 3 names.

| Draw order | Winner | City |
|---|---|---|
| 1 | one of the 30 names | its city |
| 2 | a different name | its city |
| 3 | a third name | its city |

One recalculation of the sample gave Leo (Dayton), Hugo (Provo) and Rosa (Laredo). Treat that as one draw only. The order of the three and the names will differ the next time the file opens.

The Winners tab then shows 30 entries counted, 3 winners to draw, 3 winners listed, 3 rows marked Winner and a check cell that says OK. The check means the draw order is 1 to 30 with no repeats, the number of winners listed is right, and every winner is in the list.

## What was checked

The team recalculated the sample 200 times in LibreOffice Calc 24.2.7.2 and checked each draw with Python. Every time, the draw order was exactly 1 to 30 with no repeats, there were 3 winners on 3 different rows, every winner was in the list, and the check cell said OK. Every one of the 30 names won at least once across those runs. That is a rough look at spread, not a statistical test.

RAND is pseudo-random and leaves no audit trail. This post does not cover giveaway rules or how to run one.

## Troubleshooting

### The winners change every time I edit a cell
RAND recalculates on every edit and when the file opens. Copy **C5:D34** and use Paste special > Values only. Do this before you share or save the result. Keep a copy of the formulas if you want to draw again.

### The Winners tab says "Fewer entries (5) than winners: all of them win"
You asked for more winners than there are entries. The template lists every entry instead of showing an error. Lower the number in **B1** or add entries.

### The check says CHECK
The most common cause is a name typed below row 34, which is not drawn. In the test, a name in A40 and a winners number of 0, 11, 2.5 or the word "three" all gave CHECK and no winners. Fix the list or the number. The check cell is on the Winners tab.

### Two entries have the same name
Each row is a separate entry. In the test, three rows named Sam were treated as three entries, and two different Sam rows could both win. Add the city or an ID if you need to tell them apart.

### The template only draws 30 entries
The template is built for 30 rows and up to 10 winners. Larger lists were not run. Extend the formulas yourself and test them before you use them.

## Template

[Download the .xlsx](templates/tidy-tabs-giveaway-random-winners.xlsx)

The file has three tabs. **Entries** holds 30 yellow fictional names and cities in **A5:B34**, with gray formulas in **C** to **E**. **Winners** has the yellow winners count in **B1**, the winner list in **A5:C14** and a summary with the OK or CHECK cell in **F5**. **How to use** has short notes, including the freeze step.

The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, portrait for Entries and landscape for Winners. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
