---
title: How to Find Transposed Digits When Totals Don't Match in Excel and Google Sheets
labels: excel-formulas, small-business, bookkeeping
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model on the sample log, six sample deposits, refund, size, text and blank edge cases, and 48 random logs of 20 to 499 entries with one planted error each (the planted entry was listed in all 48). Not tested in Excel or Google Sheets.
search_description: Get the gap in cents, test it with MOD(gap,9), and list every entry whose swapped digits, slid decimal, double or wrong sign explains the gap.
image_prompts: A zippered bank deposit bag and a rubber-banded stack of bills on a folding farmers market table, beside an open handwritten ledger notebook with a pencil resting on a column of figures and a wooden produce crate behind, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Bank deposit bag and banded cash beside a ledger notebook on a market table
threads: The log says $1,214.04 and the deposit says $1,223.04.\nA swap of two neighboring digits always moves a total by a multiple of 9, so test the gap in cents with =MOD(ABS(gap),9)=0.\nHere the 900-cent gap flags four entries to recheck, like $78.50 that may be $87.50.
---
Get the gap in cents with `=ROUND((deposit-log_total)*100,0)` and check `=MOD(ABS(gap),9)=0`. If it divides by 9, each entry's possible swap change for digit pair k is `=(MID(digits,k+1,1)-MID(digits,k,1))*9*10^(5-k)`, where digits is `=TEXT(cents,"000000")`, and any entry whose change equals the gap could be the typo. In the sample farmers market log, the entries add to $1,214.04 but the deposit is $1,223.04, and the $9.00 gap points to four entries to recheck: $78.50 (maybe $87.50), $45.60 (maybe $54.60), $112.80 (maybe $121.80) and $301.25 (maybe $310.25).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model, including 48 random logs with one planted error each (the planted entry was listed in all 48). The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Market sales log with a $9.00 gap and four Maybe swap amounts flagged](images/2026-11-24-team-05-find-transposed-digits-when-totals-dont-match-divisible-by-9-mid-mod-excel-google-sheets/find-transposed-digits-template.png) *Entries tab of a fictional farmers market log: 14 entries total $1,214.04 against a $1,223.04 deposit, the 900-cent gap divides by 9, and four amounts show a possible swap, such as $78.50 (maybe $87.50). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The sheet checks four mistakes: two neighboring digits swapped, a decimal slid one place, an extra or doubled entry, and a refund typed as a sale. It points at entries to recheck; it is not bookkeeping, audit or tax advice.

The steps are the same in Excel and Google Sheets.

1. On the **Entries** tab, type the amount the log should total, such as the bank deposit, in **B1**.
2. Type one entry per row from row 6: date in **A**, amount in **B**, refunds as negatives.
3. **B2** adds the log with `=SUM(B6:B505)`, and **B3** gets the gap in cents with `=IF(ISNUMBER(B1),ROUND((B1-B2)*100,0),"")`.
4. **B4** tests the gap with `=IF(B3="","",IF(B3=0,"Totals match",IF(MOD(ABS(B3),9)=0,"Yes","No")))`. A neighbor swap always changes a total by a multiple of 9. See the [MOD help from Microsoft](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3).
5. In each row, **C** gets cents with `=ROUND(ABS(B6)*100,0)`, **D** pads them with `=TEXT(C6,"000000")` and **E** counts real digits with `=LEN(C6)`. $78.50 becomes 007850, length 4.
6. **F:J** hold the swap change for each digit pair, 1-2 through 5-6, using the core formula from the top with [MID](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028) and `*SIGN($B6)` so refunds flip. A length guard blanks pairs that would swap a padding zero or put a 0 first. For $78.50, the hidden **H**, **I** and **J** show 900, -270 and -45.
7. **K** checks the swaps with `COUNTIF(F6:J6,$B$3)>0` and shows the corrected amount, such as Maybe $87.50.
8. **L** checks a slide: a gap of 9 times the cents means the entry should be 10 times larger; -0.9 times means a tenth.
9. **M** shows Extra entry? when the gap is minus the entry, or Typed twice? when that amount is in the log more than once. **N** shows the flipped amount when the gap is -2 times the entry.
10. **O** says OK or names the problem: no amount, text, more than 2 decimals, or $10,000.00 or more.
11. Read the summary in **Q5:R14**: **R11** gives the result and **R14** the CHECK.

The sheet does not find swaps of digits that are not neighbors, a missing entry, or two mistakes at once. Divisible by 9 is a clue, not proof.

## Example

The sample data is made up: 14 farmers market entries, 7 on 11/14/2026 and 7 on 11/15/2026. The log adds to $1,214.04, the deposit in **B1** is $1,223.04, the gap is 900 cents and Divisible by 9? says Yes.

| Amount | Digits | Possible swap |
|---|---|---|
| $42.00 | 004200 | |
| $18.50 | 001850 | |
| $78.50 | 007850 | Maybe $87.50 |
| $126.00 | 012600 | |
| $33.75 | 003375 | |
| $9.00 | 000900 | |
| $64.25 | 006425 | |
| $250.00 | 025000 | |
| $15.00 | 001500 | |
| $87.40 | 008740 | |
| $45.60 | 004560 | Maybe $54.60 |
| $112.80 | 011280 | Maybe $121.80 |
| $29.99 | 002999 | |
| $301.25 | 030125 | Maybe $310.25 |

The summary shows 14 entries, 4 possible swaps, no other matches and CHECK OK. Four entries fit the same $9.00 gap, so the sheet lists them all instead of picking one. Check each against its receipt.

Type a different deposit in **B1** to see the other checks with the same 14 entries:

| Deposit | Gap (cents) | Divisible by 9? | Listed | Result |
|---|---|---|---|---|
| $1,223.04 | 900 | Yes | 4 swaps (above) | 4 possible fixes to recheck |
| $989.04 | -22,500 | Yes | $250.00 slide: Maybe $25.00 | 1 possible fixes to recheck |
| $1,146.54 | -6,750 | Yes | $33.75 wrong sign: Maybe -$33.75 | 1 possible fixes to recheck |
| $1,172.04 | -4,200 | No | $42.00 Extra entry? | 1 possible fixes to recheck |
| $1,178.04 | -3,600 | Yes | none | No single entry explains the gap |
| $1,214.04 | 0 | Totals match | none | Totals match |

The Result keeps the plural even for one match. With $42.00 typed twice (log $1,256.04, deposit $1,214.04), the gap is -4,200 and both $42.00 rows say Typed twice?.

## Troubleshooting

### A $42.00 entry is flagged as maybe $402.00
TEXT pads with leading zeros, and a swap formula without the length guard swaps one of them. In LibreOffice Calc, that version suggested $402.00 for $42.00 and $405.60 for $45.60. Keep the guard if you copy the formula.

### Divisible by 9 says Yes but nothing is listed
No single mistake the sheet checks explains it. With a $1,178.04 deposit, the -$36.00 gap matched nothing. Look for two mistakes, a non-neighbor swap or a missing entry.

### The gap does not divide by 9
It is not a swap or a slide. Check the Extra entry column; with a $1,172.04 deposit, the $42.00 entry is listed as Extra entry?.

### A refund was typed as a sale
The gap is minus 2 times the amount. With a $1,146.54 deposit, the Wrong sign column shows Maybe -$33.75 for the $33.75 entry.

### An amount is not in the total
Status says Check: amount is text (not in the total), and **R6** counts it. Retype it as a number.

## Template

[Download the .xlsx](templates/tidy-tabs-find-transposed-digits.xlsx)

The file has two tabs. **Entries** holds 500 entries in rows 6 to 505: you type **B1** and **A:B**, and **C:O** hold the formulas above. The summary in **Q5:R14** counts each kind of match, Check rows and anything typed below row 505. **How to use** explains each formula. Entries prints Landscape on US Letter.

Related: [How to Match Payouts to Bank Deposits](https://tidytabs.blogspot.com/2026/11/how-to-match-payouts-to-bank-deposits.html) and [How to Count a Cash Drawer and Find Over or Short](https://tidytabs.blogspot.com/2026/10/how-to-count-cash-drawer-and-find-over.html).

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
