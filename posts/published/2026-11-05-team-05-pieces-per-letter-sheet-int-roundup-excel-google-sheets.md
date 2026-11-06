---
threads_url: https://www.threads.com/@tin_ylab/post/DfZrptbm5YL
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-how-many-pieces-fit-on.html
title: How to Calculate How Many Pieces Fit on a US Letter Sheet in Excel and Google Sheets
labels: excel-formulas, small-business, printing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (exact fractions, same 6-decimal rounding) over 33 workbooks and 1,848 comparisons. Checked: the sample, an empty list, exact fits, 0.1 in gaps, a piece 0.001 in too big, margin 0, margins with no room, a negative or text margin, an 8.5 x 14 sheet, text, blank, zero and negative sizes, odd quantities, a tie between typed and turned, 30 jobs, and random lists. Excel and Google Sheets were not used. Not run: more than 30 jobs, printing the layouts, a non-US locale.
search_description: Count how many pieces fit on a US Letter sheet, test the turned layout, and get sheets to print and spare pieces in Excel or Google Sheets.
image_prompts: A stack of printed card sheets with pencil crop marks next to a paper trimmer and a metal ruler on a worktable, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Stack of printed card sheets beside a paper trimmer and a ruler on a table
threads: Turn a 3.5 x 2 in business card sideways and 10 per US Letter sheet becomes 12.\nThat is 21 sheets for 250 cards instead of 25.\nThe free template tests both layouts and has a margin cell for your printer.
---
Multiply pieces across by pieces down with `=INT(ROUND(($B$2-2*$B$4+D8)/(B8+D8),6))*INT(ROUND(($B$3-2*$B$4+D8)/(C8+D8),6))`, then divide the quantity by that with `=ROUNDUP(E8/H8,0)`. Turned sideways, 3.5 x 2 in business cards fit 12 per sheet, so 250 cards need 21 sheets.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. INT, ROUND, ROUNDUP and MAX exist in current Excel and Google Sheets, but neither app was opened for this post. No sheet was printed, so check one test print before a big run.

![Print jobs with pieces per sheet as typed and turned, sheets needed and spare pieces](images/2026-11-05-team-05-pieces-per-letter-sheet-int-roundup-excel-google-sheets/pieces-per-us-letter-sheet-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*


## Steps

The steps are the same in Excel and Google Sheets. Put the paper size and margin in **B2:B4**. Start the job list in row 8: job name in **A**, piece width in **B**, piece height in **C**, gap between pieces in **D** and pieces needed in **E**. All sizes are in inches.

1. In **B2** type `8.5`, in **B3** type `11`, and in **B4** type the blank margin on every edge, for example `0.25`. The usable area is then 8.0 x 10.5 in.
2. In **F8**, type `=INT(ROUND(($B$2-2*$B$4+D8)/(B8+D8),6))*INT(ROUND(($B$3-2*$B$4+D8)/(C8+D8),6))`. This is pieces across times pieces down, with the piece as typed.
3. In **G8**, type the same formula with the width and height swapped: `=INT(ROUND(($B$2-2*$B$4+D8)/(C8+D8),6))*INT(ROUND(($B$3-2*$B$4+D8)/(B8+D8),6))`. This is the piece turned 90 degrees.
4. In **H8**, type `=MAX(F8,G8)` for the better of the two.
5. In **I8**, show which layout won: Turned if **G8** is larger, As typed on a tie, and Does not fit if **H8** is 0.
6. In **J8**, type `=ROUNDUP(E8/H8,0)` for the sheets to print.
7. In **K8**, type `=J8*H8-E8` for the spare pieces left over.
8. Fill **F8:K8** down as far as your jobs go.

The formula adds the gap once to the usable width because the last piece needs no gap after it.

ROUND(...,6) sits inside INT because a size that should fit exactly, such as pieces with 0.1 in gaps, can divide out as 2.9999999999 instead of 3, and INT would drop it to 2.

This counts a plain grid. It ignores bleed, mixed sizes on one sheet and the area your printer cannot print on. Use the margin in **B4** for that last one.

Help pages: Microsoft's [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test, and the pages were not read for this post.

## Example

Sample data is fictional: eight print jobs on US Letter with a 0.25 in margin on every edge.

| Job | Size (in), gap | Per sheet as typed | Per sheet turned | Layout | Qty | Sheets | Spare |
|---|---|---|---|---|---|---|---|
| Business cards | 3.5 x 2, gap 0 | 10 | 12 | Turned | 250 | 21 | 2 |
| Square stickers | 2 x 2, gap 0.125 | 15 | 15 | As typed | 120 | 8 | 0 |
| Thank-you cards | 4.25 x 5.5, gap 0 | 1 | 2 | Turned | 60 | 30 | 0 |
| Product tags | 1.5 x 3, gap 0.125 | 15 | 12 | As typed | 300 | 20 | 0 |
| Shipping labels | 8.5 x 5.5, gap 0 | 0 | 1 | Turned | 40 | 40 | 0 |
| Postcards | 4 x 6, gap 0 | 2 | 2 | As typed | 100 | 50 | 0 |
| Bookmarks | 2 x 7, gap 0.1 | 3 | 5 | Turned | 80 | 16 | 0 |
| Gift tags, small | 1.25 x 2.5, gap 0.1 | 24 | 21 | As typed | 200 | 9 | 16 |

The business cards show the point. As typed, the 8.0 in width holds 2 cards and the 10.5 in height holds 5, so 10 per sheet. Turned, the width holds 4 and the height holds 3, so 12 per sheet. That is 21 sheets instead of 25, with 2 spare cards. The 8.5 in shipping labels are wider than the 8.0 in usable width, so they fit only turned. Across all eight jobs the sample needs 194 sheets, no job fails to fit, and 4 jobs are better turned.

## Troubleshooting

### A job shows Does not fit
The piece is bigger than the usable area both ways. In the test, a piece 0.001 in too big gave 0, and a piece exactly the size of the usable area gave 1. Lower the margin in **B4** or shrink the piece.

### A row shows Check
A size, gap or quantity is blank, typed as text, zero or negative. A text or negative margin, or margins that leave no room, also give Check. Fix the yellow input cells and the row recalculates.

### Gaps of 0.1 in lose a row of pieces
This is the floating-point case above. If you remove the ROUND(...,6) from the formula, a size that fits exactly can drop one piece. Keep it.

### The count is right but the print is cut off
The sheet cannot see your printer. Many printers cannot reach the outer 0.25 in, and some need more. Raise **B4** to match, and print one test sheet first. No printed test was done for this post.

### The quantity does not divide evenly
That is normal. **ROUNDUP** rounds a partial sheet up to a whole one, and the spare column shows the leftover pieces. Gift tags in the sample leave 16.

## Template

[Download the .xlsx](templates/tidy-tabs-pieces-per-letter-sheet.xlsx)

The file has two tabs. **Sheet fit** has the paper size and margin in **B2:B4**, a settings check in **B5**, and 30 job rows from row 8 (yellow inputs in **A:E**, formulas in **F:K**). A summary in **M7:N11** shows jobs listed, jobs that do not fit, jobs where turning helps, total sheets to print and a **CHECK** cell. **How to use** has notes on each formula.

**CHECK** says OK when the settings are valid and no row shows Check or Does not fit.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
