---
title: How to Reverse a List So the Oldest Row Comes First in Excel and Google Sheets
labels: excel-formulas, small-business, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 25 workbooks and 25,275 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Flip a newest-first bank or payout export into oldest-first order with INDEX and a ROWS counter, without sorting, then add a running balance. Free .xlsx template.
image_prompts: A tall stack of plain paper bank statement pages being turned over on a wooden desk, the top half already lying face down in a shallow wooden in-tray, a large black bulldog clip set aside, all print blurred and unreadable, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Stack of statement pages being turned face down into a wooden in-tray
threads: Your bank export lists the newest row first and you want the oldest on top.\n=INDEX(Export!A$5:A$204,$H$1-ROWS($A$5:A5)+1) reads it from the bottom up, no sorting, export untouched.\nCount rows with MATCH(9.99E+307,...), not COUNTA, or one blank row drops your oldest entry.
---
Rewrite a newest-first export in oldest-first order with `=INDEX(Export!A$5:A$204,$H$1-ROWS($A$5:A5)+1)`, where H1 holds the number of rows from `=MATCH(9.99E+307,Export!$A$5:$A$204)`, then copy it across and down. In the made-up sample, the 11/01/2026 Opening deposit of $500.00 moves to the top, the 11/12/2026 Etsy payout of $412.50 moves to the bottom, and the running balance ends at $2,157.66; the download wraps the formula so empty rows stay blank instead of showing 0.

This is how you flip a column upside down without sorting. A row counter counts down from the last filled row, so INDEX reads the export from the bottom up and the original export stays untouched.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with an independent Python model (25 workbooks, 25,275 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets have the same functions but were not opened for this post; the open questions there are MATCH(9.99E+307) on a newest-first (descending) column, INDEX with a ROWS counter and SUMPRODUCT with ISTEXT.

![Bank export rows newest first beside the same rows flipped oldest first with running balance](images/2026-11-15-team-01-reverse-list-order-newest-to-oldest-first-index-rows-excel-google-sheets/reverse-list-newest-to-oldest-template.png) *Left, the export with the newest row first; right, the Flipped tab lists the same eight rows oldest first with a running balance (its summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has an **Export** tab, a **Flipped** tab and a **How to use** tab.

1. On **Export**, paste your bank, Square or Etsy export as it came, newest row first: Date in **A5**, Description in **B5**, Amount in **C5**, and down to row 204.
2. On **Flipped**, in **H1**, type `=MATCH(9.99E+307,Export!$A$5:$A$204)` to get the number of rows. 9.99E+307 is bigger than any date, so MATCH lands on the last cell in column A that holds a number.
3. In **A5**, type `=INDEX(Export!A$5:A$204,$H$1-ROWS($A$5:A5)+1)`. `ROWS($A$5:A5)` is 1 on the first row and 2 on the next, so with 8 rows the formula reads the 8th export row, then the 7th, then the 6th.
4. Copy **A5** right to **C5**, then select **A5:C5** and drag the fill handle down as far as your export goes.
5. Format column **A** as a date (MM/DD/YYYY) and column **C** as currency.
6. In **D5**, type `=SUM($C$5:C5)` for the running balance and drag it down.

The template wraps these core formulas. **H1** is `=IFERROR(MATCH(9.99E+307,Export!$A$5:$A$204),0)`, so an empty export gives 0 instead of an error. **A5:C5** use `=IF(ROWS($A$5:A5)>$H$1,"",IF(INDEX(...)="","",INDEX(...)))`, so rows past the list stay blank and a blank export cell does not come through as 0. **E** gives each row a Status: OK, Blank row (orange) or Check (red, when the date or the amount is text).

**H2** shows `=COUNTA(Export!$A$5:$A$204)` next to H1 for comparison only. COUNTA counts filled cells, not the last row, which matters as soon as a blank row sits in the middle (see Troubleshooting).

Why not **Data > Sort**? Sorting by date changes the export itself, and rows that share a date can end up in any order. The formulas leave the export as it is.

Help pages: Microsoft's [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [MATCH](https://support.google.com/docs/answer/3093378). Both MATCH pages describe approximate match on a list sorted in ascending order. A newest-first export is in descending order, and the 9.99E+307 lookup on it was checked only in LibreOffice.

## Example

Sample data is fictional: 8 made-up bank rows from 11/01/2026 to 11/12/2026, exported newest first. This is the **Flipped** tab after recalculation.

| Date | Description | Amount | Running balance | Status |
|---|---|---|---|---|
| 11/01/2026 | Opening deposit | $500.00 | $500.00 | OK |
| 11/02/2026 | Etsy payout | $199.99 | $699.99 | OK |
| 11/03/2026 | Booth fee | -$75.00 | $624.99 | OK |
| 11/05/2026 | Stripe payout | $1,150.00 | $1,774.99 | OK |
| 11/06/2026 | Shipping labels | -$42.18 | $1,732.81 | OK |
| 11/10/2026 | Square payout | $98.75 | $1,831.56 | OK |
| 11/10/2026 | Supplies | -$86.40 | $1,745.16 | OK |
| 11/12/2026 | Etsy payout | $412.50 | $2,157.66 | OK |

Rows found (H1) and Filled Date cells (H2) both show 8. The summary in **G4:H12** shows 8 rows in Export, 8 rows flipped, 0 blank rows inside, 0 text dates, 0 rows showing Check, 0 rows below row 204, $2,157.66 in both totals and CHECK OK.

Look at the two 11/10/2026 rows. In the export, Supplies sits above Square payout. Flipped, Square payout comes first. Rows that share a date come out in exactly the reverse of the export order, which is what you want from a newest-first export.

With one blank row inserted in the middle of the export, H1 shows 9 and H2 shows 8. Every row is still flipped, the 11/12/2026 Etsy payout is still last, the blank shows Blank row and CHECK stays OK.

## Troubleshooting

### The oldest row is missing
The row count came from COUNTA and the export has a blank row inside. In LibreOffice, with one blank row, COUNTA gave 8 and MATCH gave 9, and the COUNTA version left out the 11/01/2026 Opening deposit of $500.00. Use the MATCH row count in H1.

### A 0 or 12/30/1899 row appears
A blank export row was read by the core formula without the blank guard, and INDEX returned 0, which a date format shows as 12/30/1899. The template's full formula shows a blank cell and marks the row Blank row.

### The bottom row of the export is missing
Its date is text, such as "11/01/2026" typed or pasted as text. MATCH(9.99E+307) skips text, so in LibreOffice it stopped one row early and that row was not flipped. The template shows Text dates in Export 1 and CHECK Check; retype the date as a real date.

### The totals agree but CHECK says Check
An amount came in as text, such as "$412.50". SUM skips it on both tabs, so both totals show $1,745.16 and match. The row shows Check, and CHECK counts Check rows, so it catches it. Retype the amount as a number.

## Template

[Download the .xlsx](templates/tidy-tabs-reverse-list-order.xlsx)

**Export** is the yellow paste area: Date, Description and Amount in **A5:C204**, with the 8 fictional sample rows filled in. **Flipped** is all formulas: Date, Description, Amount, Running balance and Status in **A5:E204**, Rows found in **H1**, the COUNTA comparison in **H2**, and the summary in **G4:H12** (rows in Export, rows flipped, blank rows inside, text dates, rows showing Check, rows below row 204, both totals and CHECK). **How to use** explains each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Status is orange for Blank row and red for Check. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
