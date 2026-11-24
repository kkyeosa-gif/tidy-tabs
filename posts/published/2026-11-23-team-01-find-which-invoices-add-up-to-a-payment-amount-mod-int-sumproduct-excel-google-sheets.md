---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-find-which-invoices-add-up-to.html
title: How to Find Which Invoices Add Up to a Payment in Excel and Google Sheets
labels: excel-formulas, invoicing, freelancers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 25 workbooks and 1,640,175 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Number every combination of up to 12 invoices, mark each one in or out with MOD and INT, total it with SUMPRODUCT and flag the sets that equal the payment.
image_prompts: A bank deposit slip lying on a small office desk next to a fanned stack of paper invoices with three pulled out and set apart, and a desktop adding machine with a curled paper tape, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Stack of paper invoices with three pulled aside next to an adding machine
threads: Which of your 12 open invoices add up to that $416.24 payment?\nNumber every combination from 1 to 4,095 and mark each invoice in or out with =MOD(INT(n/2^k),2).\nSUMPRODUCT totals each set. In the sample, 2 different sets hit $416.24, so you know to ask the client.
---
Number every combination of your open invoices from 1 to 4,095 (all the ways to pick from 12), use `=MOD(INT(n/2^k),2)` to mark whether invoice k+1 is in combination n, total each combination with `=ROUND(SUMPRODUCT(amounts,in_out),2)`, and flag the totals that equal the payment. In the made-up list of 12 invoices, a $416.24 payment matches 2 different sets, INV-1042 + INV-1043 + INV-1049 and INV-1043 + INV-1048 + INV-1052, so the download lists both and you know to ask the client which one they paid.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (25 workbooks, 1,640,175 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Twelve sample invoice amounts with Yes marks for two sets totaling $416.24](images/2026-11-23-team-01-find-which-invoices-add-up-to-a-payment-amount-mod-int-sumproduct-excel-google-sheets/invoices-that-add-up-to-payment-template.png) *Invoices tab with 12 fictional invoices and a $416.24 payment: Match 1 (INV-1042, INV-1043, INV-1049) and Match 2 (INV-1043, INV-1048, INV-1052) each total $416.24, found among 4,095 combinations checked. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

No Solver or macro is needed. The sheet tries every set, so it takes up to 12 invoices (4,095 rows); 13 would need 8,191 and 20 over a million. Narrow the list by client and dates first.

The steps are the same in Excel and Google Sheets.

1. On the **Invoices** tab, type the payment in **B1** and the allowed difference in **B2**, such as $0.00 for an exact match or $3.00 to allow a small fee.
2. Type one open invoice per row from row 5 to row 16: Invoice # in **A** and Amount in **B**. Leave no blank rows between them.
3. On the **Combos** tab, **B3:M3** pull the amounts across with `=INDEX(Invoices!$B$5:$B$16,1)` (then 2, 3 up to 12).
4. In **A5**, number the combinations with `=IF(ROWS(A$5:A5)>2^COUNT(Invoices!$B$5:$B$16)-1,"",ROWS(A$5:A5))`. With 12 amounts, rows 5 to 4099 hold combinations 1 to 4,095.
5. In **B5**, mark the first invoice in or out with `=IF($A5="","",MOD(INT($A5/2^0),2))`. **C5** uses `2^1`, and so on up to **M5** with `2^11`. Each row's 1s and 0s are its number in binary, so every set shows up once. See the [MOD help from Microsoft](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3).
6. In **N5**, total the set with `=IF($A5="","",ROUND(SUMPRODUCT($B$3:$M$3,B5:M5),2))`. SUMPRODUCT multiplies each amount by its 1 or 0 and adds the results; see the SUMPRODUCT help from [Microsoft](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) and [Google](https://support.google.com/docs/answer/3094294). ROUND is a guard that keeps the total in whole cents.
7. In **O5**, flag a match with `=IF(OR(N5="",Invoices!$I$11<>"OK"),"",IF(ABS(N5-ROUND(Invoices!$B$1,2))<=Invoices!$B$2,1,0))`. The flag waits until **I11** is OK: the payment a number above 0 and the difference a number of 0 or more.
8. In **P5**, number the matches with `=IF(O5=1,COUNTIF($O$5:O5,1),"")`, then fill **A5:P5** down to row 4099.
9. Back on **Invoices**, **D5** shows "Yes" for each invoice in the first match with `=IFERROR(IF(INDEX(Combos!$B$5:$M$4099,MATCH(1,Combos!$P$5:$P$4099,0),ROWS(D$5:D5))=1,"Yes",""),"")`. **E** and **F** use `MATCH(2,...)` and `MATCH(3,...)` for the second and third matches.
10. Check **I5**, `=COUNTA(B5:B16)-COUNT(B5:B16)`, the amounts typed as text. It must be 0.

The sheet finds sets that add up; it does not know which invoices the client meant. Confirm with the client before you mark anything paid. Booking short payments or fees is not covered.

## Example

The invoices are made up. Payment $416.24, allowed difference $0.00.

| Invoice # | Amount | Match 1 | Match 2 |
|---|---|---|---|
| INV-1041 | $412.50 | | |
| INV-1042 | $189.99 | Yes | |
| INV-1043 | $76.25 | Yes | Yes |
| INV-1044 | $1,050.00 | | |
| INV-1045 | $230.10 | | |
| INV-1046 | $95.00 | | |
| INV-1047 | $318.40 | | |
| INV-1048 | $64.99 | | Yes |
| INV-1049 | $150.00 | Yes | |
| INV-1050 | $520.35 | | |
| INV-1051 | $88.80 | | |
| INV-1052 | $275.00 | | Yes |
| Total of this set | | $416.24 | $416.24 |
| Combination # | | 262 | 2180 |

Combination 262 is 256 + 4 + 2, so invoices 2, 3 and 9 are in: $189.99 + $76.25 + $150.00. Combination 2180 is 2048 + 128 + 4: invoices 3, 8 and 12, or $76.25 + $64.99 + $275.00.

The summary in **I4:I12** shows 12 invoices, 0 text amounts, 4,095 combinations checked, 2 matches, 0 rows showing Check and CHECK OK.

A $413.24 payment with $3.00 allowed gives "5 matches (first 3 listed)": INV-1041 alone at $412.50, INV-1046 + INV-1047 at $413.40, and the first $416.24 set. The $3.00 edge counts as within.

## Troubleshooting

### No match on a payment that should match
A text amount or a blank row cuts the combinations. With INV-1045 typed as text "150.00", 11 amounts counted, 2,047 combinations ran, INV-1052 was never tried and the INV-1043 + INV-1048 + INV-1052 set was lost. With row 13 empty, INV-1052 was never tried: No match. **I5** or Status shows which.

### A match lists an invoice whose amount is text
The text amount counts as $0.00. In the "150.00" case a false Match 2 appeared, INV-1042 + INV-1043 + INV-1045 + INV-1049, and row 17 still read $416.24 because SUMIF skips text. Retype the amount as a number.

### More than one match
Two sets add up to the same amount, as in the sample. The sheet lists up to 3 and the count in **I8**. Ask the client.

### Nothing listed and Result says "Type the payment in B1"
The payment is blank or $0.00, or the allowed difference is text, such as "3". **I11** shows Check and nothing is flagged. Type both as numbers.

For one-to-one matching, see [how to match payouts to bank deposits](https://tidytabs.blogspot.com/2026/11/how-to-match-payouts-to-bank-deposits.html). The same INT and MOD numbering lists [every size and color combination](https://tidytabs.blogspot.com/2026/11/how-to-list-every-size-and-color.html).

## Template

[Download the .xlsx](templates/tidy-tabs-invoices-that-add-up-to-payment.xlsx)

The file has three tabs. **Invoices** holds the payment in **B1**, the allowed difference in **B2** and up to 12 invoices in rows 5 to 16, with Status in **C**, the first three matching sets in **D:F**, their totals and combination numbers in rows 17 and 18, and a summary in **H4:I12**.

Status flags no amount, a text amount, no invoice #, an invoice # twice, more than 2 decimals and a blank row above a filled row. A negative credit memo is still used in the combinations but shows Check. An invoice below row 16 is not used and turns CHECK to Check.

**Combos** is visible, with one row per combination in rows 5 to 4099; it prints only the first 32. **How to use** explains each formula. Both data tabs print Landscape on US Letter.

Yellow cells are inputs, gray cells are formulas, and every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
