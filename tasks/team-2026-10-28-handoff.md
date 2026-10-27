# Handoff 2026-10-28: five templates built and checked (LibreOffice Calc 24.2.7.2 only)

Status: topics (tasks/team-topics-2026-10-28.md) done; the posts are the next step. The templates were built by five new functions in `scripts/build-templates.py` (`loan_payment_schedule`, `recurring_billing_dates`, `closest_match`, `inches_to_fractions`, `batches_and_rotation`, added after `next_anniversary` and registered at the end of the main block). Only these five were run (a small runner that imports the script and calls just them, with `PYTHONDONTWRITEBYTECODE=1`); older .xlsx were not rewritten. Nothing was committed. `git status` shows the five new .xlsx, the edited script, tasks/briefs.md, this file, and files that were there before this task (tasks/topics-seed.md, tasks/team-topics-2026-10-28.md, tasks/hooks-2026-10-28.md).

Exact version string used for `tested_in`: **LibreOffice 24.2.7.2 420(Build:2)** (Calc, Linux, headless). Write "LibreOffice Calc 24.2.7.2". Nothing was opened in Excel or Google Sheets and nothing may be claimed about them. Method: `soffice --headless --convert-to xlsx` into a scratch dir, read back with openpyxl `data_only=True`, compared with independent Python models that read only the typed inputs (Fraction cents with half-up rounding for the loan, `calendar` and `datetime` for dates, Decimal `min(key=abs)` for closest, Fraction rounding for inches, `divmod` for batches). 126 workbooks (the 5 committed samples plus 121 edited copies: 25 loan, 34 dates, 21 closest, 22 inches, 24 batches, counted with their sample files), 44,992 cell comparisons, 0 mismatches in the final run, no error value (#NAME?, #VALUE!, #REF!, #N/A, #DIV/0!, #NUM!, Err:) in any workbook. Formulas use only: PMT, IPMT, EDATE, EOMONTH, DATE, YEAR, MONTH, DAY, WEEKDAY, TEXT, MROUND, GCD, INT, MOD, ROUND, ROUNDUP, ABS, MIN, MAX, SMALL, INDEX, MATCH, COUNT, COUNTA, COUNTIF, SUMPRODUCT, SUM, ROWS, ROW, TRIM, IF, AND, OR, ISNUMBER (no AGGREGATE, TODAY, NOW, FILTER, LET, LAMBDA, XLOOKUP, no array entry, no `_xlfn`). Sample data fictional, US dates (MM/DD/YYYY) and dollars. Yellow (FFF2CC) = inputs, gray (EDEDED) = formulas. Each file has a "How to use" sheet and one CHECK cell. PDF export: every page US Letter (landscape main tabs, portrait How to use).

Titles and slugs are in tasks/team-topics-2026-10-28.md. Posts must say: LibreOffice Calc 24.2.7.2 only, sample data fictional, no advice (borrowing, interest, tax, what to quote, when a client must pay).

---

## 1. Loan payment schedule (`templates/tidy-tabs-loan-payment-schedule.xlsx`, function `loan_payment_schedule`)
- Sheets: Schedule (landscape, print area A1:I43 = sample rows, header row 7 repeats), How to use.
- Inputs (yellow): C1 loan amount $12,000.00, C2 yearly rate 7.00% (typed as 0.07), C3 months 36 (1 to 60), C4 first payment date 11/28/2026. Labels in A1:A4. C5 = input check (OK or a message; the table stays empty unless it is OK).
- Table: header row 7 (Payment no., Date, Payment, Interest, Principal, Balance after), rows 8 to 67 (60 payments, rows past the term are blank).
- Summary (labels H, values I): I1 monthly payment, I2 last payment, I3 total interest, I4 total paid, I5 last payment date, I6 CHECK.
- Formulas to quote (row 8; row 9 and below use the balance above, F8, instead of $C$1):
  - C5 input check: `=IF(AND(ISNUMBER(C1),ISNUMBER(C2),ISNUMBER(C3),ISNUMBER(C4)),IF(C1<=0,"Amount must be above 0",IF(OR(C2<0,C2>1),"Rate must be 0% to 100%",IF(OR(C3<1,C3>60,C3<>INT(C3)),"Months must be a whole number from 1 to 60","OK"))),"Type a number in each yellow cell")`
  - I1 monthly payment: `=IF(C5="OK",IF(C2=0,ROUND(C1/C3,2),ROUND(-PMT(C2/12,C3,C1),2)),"")`
  - A8 payment no.: `=IF($C$5="OK",IF(ROWS(A$8:A8)<=$C$3,ROWS(A$8:A8),""),"")`
  - B8 date: `=IF(A8="","",EDATE($C$4,A8-1))`
  - D8 interest: `=IF(A8="","",IF($C$2=0,0,ROUND(-IPMT($C$2/12,A8,$C$3,$C$1),2)))`
  - C8 payment: `=IF(A8="","",IF(A8=$C$3,$C$1+D8,$I$1))` (row 9: `...IF(A9=$C$3,F8+D9,$I$1)`; the last payment is the balance still owed plus that month's interest)
  - E8 principal: `=IF(A8="","",ROUND(C8-D8,2))`
  - F8 balance: `=IF(A8="","",ROUND($C$1-E8,2))` (row 9: `ROUND(F8-E9,2)`)
  - I6 CHECK: `=IF(C5<>"OK","CHECK",IF(AND(ROUND(SUM(E8:E67)-C1,2)=0,INDEX(F8:F67,C3)=0,ROUND(I4-I3-C1,2)=0,COUNT(A8:A67)=C3),"OK","CHECK"))`
- Example table values (recalculated, equal to the Python model): I1 $370.53, I2 $370.38, I3 $1,338.93, I4 $13,338.93, I5 10/28/2029, I6 OK.
  | No. | Date | Payment | Interest | Principal | Balance |
  | 1 | 11/28/2026 | $370.53 | $70.00 | $300.53 | $11,699.47 |
  | 2 | 12/28/2026 | $370.53 | $68.25 | $302.28 | $11,397.19 |
  | 3 | 01/28/2027 | $370.53 | $66.48 | $304.05 | $11,093.14 |
  | 35 | 09/28/2029 | $370.53 | $4.29 | $366.24 | $368.23 |
  | 36 | 10/28/2029 | $370.38 | $2.15 | $368.23 | $0.00 |
- Rounding decision (open case from the topics file): every payment but the last is the rounded PMT; the last payment is balance plus interest, so the final balance is exactly $0.00 and principal adds up to the loan. The last payment differs from the others by a few cents (sample 15 cents less). Interest is IPMT on the exact (unrounded-payment) schedule, rounded to cents each month; a real lender may compute from the actual balance and differ by cents.
- Edge findings (each compared with the model, all matched; CHECK OK):
  - 0% rate: PMT/IPMT are avoided; payment = amount / months rounded ($12,000 / 36 = $333.33, last payment $333.45, interest $0.00). $500.00 over 7 months at 0%: $71.43 x 6 then $71.42.
  - 1 month: payment $12,070.00 (interest $70.00).
  - 60 months at 7%: $237.61, last $237.85, total interest $2,256.84.
  - $500.00 / 24 months / 7%: $22.39, last $22.30, interest $37.27. $250,000.00 / 60 / 7%: $4,950.30, last $4,950.27, interest $47,017.97. 4.25% / 36: $355.62, last $355.75, interest $802.45. 100% / 6 months: $2,622.09, last $2,622.10.
  - Tiny loan: $1.00 / 60 months / 12%: payment $0.02, last payment $0.17 (rounding piles up in the last payment; say this is why a last payment can look odd).
  - First payment 01/31/2027: dates 01/31, 02/28, 03/31, 04/30 (EDATE counts from the first date each time). First payment 02/29/2028: 02/29/2028, 03/29/2028 ... 02/28/2029 (13th payment). Start 12/31/2026 works.
  - Bad input (message in C5, empty table, I6 CHECK): rate typed as 7 (not 7%) -> "Rate must be 0% to 100%"; months 61, 0 or 12.5 -> months message; amount 0 -> "Amount must be above 0"; negative rate; amount, rate or date typed as text ("12k", "7%", "11/28/2026") -> "Type a number in each yellow cell". Troubleshooting use: "table is empty" = read C5; "rate looks 100 times too big" = typed 7 instead of 7%.
- Limits: no extra payments, no payment typed over the PMT value, no daily interest, no fees. Not run: rate above 100%, more than 60 months.

---

## 2. Recurring billing dates (`templates/tidy-tabs-recurring-billing-dates.xlsx`, function `recurring_billing_dates`)
- Sheets: Dates (landscape, print area A1:G32, header row 6 repeats), How to use.
- Inputs (yellow): C1 start date 11/15/2026, C2 dropdown "Twice a month" or "Every 14 days" (sample: Twice a month), C3 dates to list 24 (1 to 26). C4 = input check (OK or a message).
- Table: header row 6 (No., Date, Weekday, Weekend?), rows 7 to 32 (26 slots). Weekend rows turn red.
- Summary (labels F, values G): G7 first date, G8 last date, G9 dates listed, G10 dates on a Saturday or Sunday, G11 CHECK.
- Formulas to quote:
  - B7 first date: `=IF(A7="","",IF($C$2="Every 14 days",$C$1,IF(DAY($C$1)<=15,DATE(YEAR($C$1),MONTH($C$1),15),EOMONTH($C$1,0))))`
  - B8 next date (fill down): `=IF(A8="","",IF($C$2="Every 14 days",B7+14,IF(DAY(B7)=15,EOMONTH(B7,0),DATE(YEAR(B7),MONTH(B7)+1,15))))`
  - Core rule without the guards: `=IF(DAY(B7)=15,EOMONTH(B7,0),DATE(YEAR(B7),MONTH(B7)+1,15))` and `=B7+14`.
  - C7 weekday `=IF(B7="","",TEXT(B7,"ddd"))`; D7 `=IF(B7="","",IF(WEEKDAY(B7,2)>5,"Weekend",""))`; A7 `=IF($C$4="OK",IF(ROWS(A$7:A7)<=$C$3,ROWS(A$7:A7),""),"")`
  - G11 CHECK: `=IF(C4<>"OK","CHECK",IF(AND(G9=C3,SUMPRODUCT(--(A8:A32<>""),--(B8:B32>B7:B31))=G9-1),"OK","CHECK"))`
- Example table values, sample (Twice a month, start 11/15/2026, 24 dates): G7 11/15/2026, G8 10/31/2027, G9 24, G10 7, G11 OK.
  | No. | Date | Weekday | Weekend? |
  | 1 | 11/15/2026 | Sun | Weekend |
  | 2 | 11/30/2026 | Mon | |
  | 3 | 12/15/2026 | Tue | |
  | 4 | 12/31/2026 | Thu | |
  | 5 | 01/15/2027 | Fri | |
  | 6 | 01/31/2027 | Sun | Weekend |
  | 7 | 02/15/2027 | Mon | |
  | 8 | 02/28/2027 | Sun | Weekend |
  Every 14 days from 11/15/2026: 11/15/2026, 11/29/2026, 12/13/2026, 12/27/2026, 01/10/2027, 01/24/2027 (all Sundays, so with 26 dates all 26 are flagged Weekend).
- Open case resolved: the list starts on the first billing date on or after the start date (a start of 11/10 gives 11/15; 11/16 gives 11/30; a start that already is the 15th or the last day is kept). The weekend flag only flags; dates are never moved.
- Edge findings (13 start dates x both modes, 24 or 26 dates, every date, weekday and flag matched `calendar`/`datetime`; no repeats, always increasing):
  - Start 01/31/2026 twice a month: 01/31, 02/15, 02/28, 03/15, 03/31. Start 02/15/2026: 02/15, 02/28, 03/15, 03/31, 04/15. Start 02/28/2026: 02/28, 03/15, 03/31, 04/15, 04/30. Start 12/31/2026: 12/31, 01/15/2027, 01/31, 02/15, 02/28. Start 02/29/2028: 02/29, 03/15/2028, 03/31, 04/15, 04/30. Start 11/01/2026: first date 11/15/2026. Start 11/30/2026: kept.
  - Every 14 days from 01/31/2026: 01/31, 02/14, 02/28, 03/14, 03/28 (no month-end logic, as intended). From 02/29/2028: 02/29, 03/14, 03/28, 04/11, 04/25 (24 dates, none on a weekend).
  - 1 date and 26 dates both work.
  - Bad input (message in C4, empty list, G11 CHECK): start date typed as text -> "Type a date in C1"; "Weekly" in C2 -> "Pick Twice a month or Every 14 days in C2"; count 27, 0 or 12.5 -> "Dates to list must be a whole number from 1 to 26"; count typed as text -> "Type a number in C3".
- Limits: no business-day shifting, no holidays (WORKDAY not used), no TODAY. Not run: more than 26 dates, dates before 1900 or after 9999. Excel/Sheets TEXT "ddd" output not tested.

---

## 3. Closest match (`templates/tidy-tabs-closest-match.xlsx`, function `closest_match`)
- Sheets: Sizes (landscape, print area A1:S30, header row 8 repeats), How to use.
- Inputs (yellow): standard sizes in inches across H7:S7 (12 slots, 9 filled: 8, 10, 12, 14, 16, 18, 20, 24, 30); orders A9:B108 (A order name, B item length in inches; 15 fictional print orders PR-2101 to PR-2115).
- Columns: C closest size, D off by, E next size that fits, F note, G spacer, H:S visible difference helper columns (header H8 `="Off from "&H7`).
- Summary (labels A1:A5, values D1:D5): D1 orders with a length, D2 closest size is smaller than the item, D3 longer than every size, D4 largest gap, D5 CHECK.
- Formulas to quote (row 9):
  - H9 (fill right to S9 and down): `=IF(AND(ISNUMBER($B9),ISNUMBER(H$7)),ROUND(ABS($B9-H$7),6),"")`
  - C9 closest: `=IF($B9="","",IF(ISNUMBER($B9),IF(COUNT($H9:$S9)=0,"No sizes",INDEX($H$7:$S$7,MATCH(MIN($H9:$S9),$H9:$S9,0))),"Check length"))`
  - D9 `=IF(ISNUMBER(C9),MIN($H9:$S9),"")`
  - E9 next size that fits: `=IF(ISNUMBER(C9),IF(COUNTIF($H$7:$S$7,">="&B9)=0,"None",SMALL($H$7:$S$7,COUNTIF($H$7:$S$7,"<"&B9)+1)),"")`
  - F9 tie note: `=IF(ISNUMBER(C9),IF(COUNTIF($H9:$S9,D9)>1,"Tie, first listed used",""),"")`
  - D2 `=SUMPRODUCT(--ISNUMBER(C9:C108),--(C9:C108<B9:B108))`; D3 `=COUNTIF(E9:E108,"None")`; D4 `=IF(COUNT(D9:D108)=0,"",MAX(D9:D108))`
  - D5 CHECK: `=IF(AND(COUNT(H7:S7)>0,COUNT(H7:S7)=COUNTA(H7:S7),COUNT(B9:B108)=COUNTA(B9:B108),COUNT(D9:D108)=COUNT(B9:B108),SUMPRODUCT(--(COUNTIF(H7:S7,H7:S7)>1),--ISNUMBER(H7:S7))=0,COUNTA(A109:B500)=0),"OK","CHECK")`
- Example table values, sample (sizes 8 10 12 14 16 18 20 24 30). Summary: D1 15, D2 9, D3 2, D4 5, D5 OK.
  | Order | Length | Closest | Off by | Next size that fits | Note |
  | PR-2101 | 9 | 8 | 1 | 10 | Tie, first listed used |
  | PR-2102 | 11.5 | 12 | 0.5 | 12 | |
  | PR-2103 | 13 | 12 | 1 | 14 | Tie, first listed used |
  | PR-2104 | 15.25 | 16 | 0.75 | 16 | |
  | PR-2105 | 17 | 16 | 1 | 18 | Tie, first listed used |
  | PR-2106 | 18 | 18 | 0 | 18 | |
  | PR-2107 | 21 | 20 | 1 | 24 | |
  | PR-2108 | 22 | 20 | 2 | 24 | Tie, first listed used |
  | PR-2109 | 26 | 24 | 2 | 30 | |
  | PR-2110 | 27 | 24 | 3 | 30 | Tie, first listed used |
  | PR-2111 | 29 | 30 | 1 | 30 | |
  | PR-2112 | 31.5 | 30 | 1.5 | None | |
  | PR-2113 | 5 | 8 | 3 | 8 | |
  | PR-2114 | 35 | 30 | 5 | None | |
  | PR-2115 | 24 | 24 | 0 | 24 | |
  Note the topics file's 21 example (sizes 18, 24, 30) is a tie only with that short list: with the sample list 21 is closest to 20 (1 inch), no tie. The tie for 21 shows in the edited copy with sizes 18, 24, 30: closest 18, "Tie, first listed used", next size that fits 24.
- Open cases resolved: tie -> the size listed first (leftmost in H7:S7, not the smaller one; with sizes in rising order that is the smaller) and a visible note; blank cell in the size list -> skipped; text in the list or length -> not a number, the length row shows "Check length" and D5 shows CHECK; unsorted list works (E uses SMALL and COUNTIF, not an approximate MATCH); no array formula, no AGGREGATE.
- Edge findings (all matched `min(sizes, key=abs diff)`, ties first listed):
  - Sizes 18, 24, 30: 21 -> 18 (tie), 27 -> 24 (tie), 1 -> 18 (off by 17), 100 -> 30, next size up for 100 is None.
  - Unsorted list 30, 10, 24, 8, 18, 12: length 9 -> 10 (tie between 10 and 8; 10 is listed first), 11 -> 10, 21 -> 24 (tie 18/24; 24 listed first), 27 -> 30 (tie 24/30; 30 listed first), 28 -> 30, 3 -> 8, 31 -> 30, 11 -> 10 (tie with 12; 10 listed first).
  - Blank cell in the middle of the size list (8, blank, 12, blank, 16): skipped, D5 OK.
  - Decimals: sizes 10.1, 10.3, 11.5, 12 with length 10.2 -> 10.1 tie (the ROUND to 6 places makes 0.1 and 0.1 equal; without it float noise could hide the tie), 10.9 -> 10.3 (tie with 11.5), 11.75 -> 11.5 (tie with 12). Sizes 0.1 to 0.4 step 0.1 with 0.15, 0.25, 0.35 all tie (first listed). Sizes 0, 5, 10: length 0 -> 0, 2.5 -> 0 (tie), 7.5 -> 5 (tie); a negative length -3 -> 0, off by 3.
  - Below the smallest (length 1, sizes 12, 16, 24): closest 12, next size that fits 12. Above the largest (99): closest 24, None. Exact match (12, 24): closest equal, off by 0, next size equal to it. Length 24.01: closest 24, next None.
  - One size only (20): everything maps to 20. 12 sizes in all slots, 100 lengths, and 5 random lists all matched.
  - Duplicate size in the list (12, 12, 16): the tie note shows for 13, 14 (two equal gaps) and D5 shows CHECK. Troubleshooting use: "Tie note on a size that is not really tied" = the same size listed twice.
  - Length typed as text "12": row shows "Check length", D5 CHECK. Size typed as text "10": ignored in the helper columns, D5 CHECK. No sizes at all: "No sizes", CHECK. Empty order list: summary zeros/blank, D5 OK.
- Limits: this finds the numerically nearest size (above or below); it does not say what to quote or sell. Not run: more than 12 sizes, negative sizes.

---

## 4. Decimal inches to fractions (`templates/tidy-tabs-inches-to-fractions.xlsx`, function `inches_to_fractions`)
- Sheets: Inches (landscape, print area A1:K30, header row 4 repeats), How to use.
- Inputs (yellow): C1 precision dropdown 8, 16 or 32 (sample 16); decimal inches in A5:A104 (sample 20 values). C2 = precision check (OK or "Pick 8, 16 or 32 in C1").
- Columns: B rounded number, C fraction text from TEXT (fixed denominator), D whole inches, E parts of an inch (steps of 1/N), F reduced fraction (ruler style, built from D and E with GCD), G feet and inches, H rounded minus typed.
- Summary (labels J, values K): K5 numbers rounded, K6 rounded up, K7 rounded down, K8 largest difference, K9 CHECK.
- Formulas to quote (row 5; N is C1):
  - B5: `=IF(A5="","",IF($C$2<>"OK","Check precision",IF(ISNUMBER(A5),IF(A5<0,"Check number",MROUND(A5,1/$C$1)),"Check number")))`; core `=MROUND(A5,1/$C$1)`
  - C5: `=IF(ISNUMBER(B5),IF(B5=0,"0",TRIM(TEXT(B5,"# ?/"&$C$1))),"")`; core `=TRIM(TEXT(B5,"# ?/16"))`
  - D5 `=IF(ISNUMBER(B5),INT(B5),"")`; E5 `=IF(ISNUMBER(B5),ROUND((B5-D5)*$C$1,0),"")`
  - F5: `=IF(ISNUMBER(B5),IF(E5=0,D5&"",IF(D5=0,"",D5&" ")&E5/GCD(E5,$C$1)&"/"&$C$1/GCD(E5,$C$1)),"")`
  - G5 feet and inches (long; the post can quote the idea: feet from `INT(B5/12)`, inches left over `B5-12*INT(B5/12)`, same GCD reduction; under 12 inches the feet are left out); H5 `=IF(ISNUMBER(B5),ROUND(B5-A5,8),"")`
  - K9 CHECK: `=IF(C2<>"OK","CHECK",IF(AND(COUNT(A5:A104)=COUNTA(A5:A104),COUNT(A5:A104)=COUNT(B5:B104),SUMPRODUCT(--(A5:A104<0))=0,OR(K5=0,K8<=0.5/C1+0.000000001),COUNTA(A105:A500)=0),"OK","CHECK"))`
- Example table values, sample at 1/16 (K5 20, K6 4, K7 5, K8 0.03125, K9 OK). Columns: typed, rounded, TEXT fraction (C), reduced (F), feet and inches (G), difference.
  | Typed | Rounded | Fraction (TEXT) | Reduced | Feet and inches | Diff |
  | 0.3125 | 0.3125 | 5/16 | 5/16 | 5/16 in | 0 |
  | 2.7 | 2.6875 | 2 11/16 | 2 11/16 | 2 11/16 in | -0.0125 |
  | 5 | 5 | 5 | 5 | 5 in | 0 |
  | 11.875 | 11.875 | 11 14/16 | 11 7/8 | 11 7/8 in | 0 |
  | 0.03 | 0 | 0 | 0 | 0 in | -0.03 |
  | 12.999 | 13 | 13 | 13 | 1 ft 1 in | 0.001 |
  | 0.5 | 0.5 | 8/16 | 1/2 | 1/2 in | 0 |
  | 7.0625 | 7.0625 | 7 1/16 | 7 1/16 | 7 1/16 in | 0 |
  | 36.4 | 36.375 | 36 6/16 | 36 3/8 | 3 ft 3/8 in | -0.025 |
  | 0.09375 | 0.125 | 2/16 | 1/8 | 1/8 in | 0.03125 |
  | 24.25 | 24.25 | 24 4/16 | 24 1/4 | 2 ft 1/4 in | 0 |
  | 48.5 | 48.5 | 48 8/16 | 48 1/2 | 4 ft 1/2 in | 0 |
  | 1.03125 | 1.0625 | 1 1/16 | 1 1/16 | 1 1/16 in | 0.03125 |
  | 0 | 0 | 0 | 0 | 0 in | 0 |
  | 3.14159 | 3.125 | 3 2/16 | 3 1/8 | 3 1/8 in | -0.01659 |
  | 15.9 | 15.875 | 15 14/16 | 15 7/8 | 1 ft 3 7/8 in | -0.025 |
  | 72 | 72 | 72 | 72 | 6 ft 0 in | 0 |
  | 100.4375 | 100.4375 | 100 7/16 | 100 7/16 | 8 ft 4 7/16 in | 0 |
  | 8.5 | 8.5 | 8 8/16 | 8 1/2 | 8 1/2 in | 0 |
  | 0.96875 | 1 | 1 | 1 | 1 in | 0.03125 |
  The topics file's spot checks hold: `=MROUND(2.7,1/16)` = 2.6875; `=TRIM(TEXT(MROUND(0.3125,1/16),"# ?/16"))` = "5/16"; 5 gives "5" after TRIM.
- Open cases resolved: a text format built with `&` (`"# ?/"&$C$1`) works in LibreOffice, so no CHOOSE was needed. TEXT keeps the fixed denominator (8/16, not 1/2), so the template adds the reduced column from GCD. TEXT of 0 gives only spaces, so the IF writes "0" (0.03 at 1/16 rounds to 0). Negative numbers and text show "Check number" (MROUND errors on mixed signs, so negatives are refused up front). Exactly halfway values round up (MROUND rounds half away from zero).
- Edge findings (all rows of 14 edited copies (plus the sample) matched the Fraction model: rounded value, TEXT text, whole, parts, reduced text, feet and inches text, difference, summary):
  - Halfway at 1/16: 0.03125 -> 0.0625 (1/16), 0.09375 -> 0.125, 0.15625 -> 0.1875, 1.03125 -> 1.0625, 0.96875 -> 1, 11.96875 -> 12 ("1 ft 0 in"), 12.96875 -> 13. Halfway at 1/8: 0.0625 -> 0.125, 0.1875 -> 0.25. Halfway at 1/32: 0.015625 -> 0.03125, 0.046875 -> 0.0625, 0.984375 -> 1. At 1/16 the same 0.015625 rounds down to 0.
  - Rounds up to the next whole inch: 11.99 -> 12 ("1 ft 0 in"), 23.999 -> 24, 35.97 -> 36 ("3 ft 0 in"), 143.99 -> 144. 12.0001 -> 12.
  - Precision 8: 11.875 -> "11 7/8" in both columns (TEXT "11 7/8", reduced "11 7/8"); 11.8749 -> same. Precision 32: 0.015625 -> 0.03125 ("1/32").
  - Tiny values at 1/16: 0.0001, 0.03, 0.031 and 0.0312 -> 0 ("0"); 0.032 and 0.0313 (just above the halfway point 0.03125) -> 0.0625 ("1/16"); 0.06 -> 1/16. At precision 8: 0.0624 -> 0, 0.0625 -> 0.125, 100.07 -> 100 1/8 ("8 ft 4 1/8 in"), 17.99 -> 18 ("1 ft 6 in"), 12.5 -> "12 1/2".
  - Blank rows in the middle stay blank. 100 random values with 1 to 4 decimals at 16 and at 32 matched.
  - Negative value (-1): "Check number", K9 CHECK. A number typed as text ("2.5") or a word: "Check number", K9 CHECK. Precision 10, 0, text "16": every row "Check precision", C2 shows the message, K9 CHECK. Precision cell cleared: C2 shows the message too (a blank is not 8, 16 or 32).
  - Largest difference never exceeded half a step (1/32 at 16).
- Limits: Excel and Google Sheets TEXT fraction output was not tested; the post says so and links the Microsoft and Google TEXT help pages. Rounds for display and cut lists only; says nothing about tolerances or fit. Not run: values above 1,000,000, a precision of 64.

---

## 5. Batches and rotation (`templates/tidy-tabs-batches-and-rotation.xlsx`, function `batches_and_rotation`)
- Sheets: Orders (portrait, print area A1:E50), Helpers, Summary (landscape), How to use.
- Inputs (yellow): Orders B1 batch size 10; Orders A5:A204 order numbers (sample 45, numbers 4101 to 4145); Helpers A4:A8 first names (sample Ana, Ben, Cora; 1 to 5, no empty row between). Orders B2 = batch size check (OK or a message).
- Orders columns: B item no., C batch, D position in batch, E helper (header row 4, data rows 5 to 204). The first row of each batch is green.
- Summary: A1:B6 (labels A, values B): B1 orders counted, B2 batch size, B3 number of batches, B4 orders in the last batch, B5 helpers listed, B6 CHECK. Helper table D1:E8: D2:D6 names, E2:E6 orders per helper, E7 total, E8 most minus fewest.
- Formulas to quote (row 5):
  - B5 item no.: `=IF(A5="","",COUNTA($A$5:A5))`
  - C5 batch: `=IF(B5="","",IF($B$2="OK",INT((B5-1)/$B$1)+1,"Check batch size"))`; core `=INT((B5-1)/$B$1)+1`
  - D5 position: `=IF(B5="","",IF($B$2="OK",MOD(B5-1,$B$1)+1,"Check batch size"))`; core `=MOD(B5-1,$B$1)+1`
  - E5 helper: `=IF(B5="","",IF(COUNTA(Helpers!$A$4:$A$8)=0,"No helper",INDEX(Helpers!$A$4:$A$8,MOD(B5-1,COUNTA(Helpers!$A$4:$A$8))+1)&""))`; core `=INDEX(Helpers!$A$4:$A$8,MOD(B5-1,COUNTA(Helpers!$A$4:$A$8))+1)`
  - B2: `=IF(ISNUMBER(B1),IF(AND(B1>=1,B1=INT(B1)),"OK","Batch size must be a whole number, 1 or more"),"Batch size must be a number")`
  - Summary B3 `=IF(B2="","",ROUNDUP(B1/B2,0))`; B4 `=IF(OR(B2="",B1=0),"",B1-(B3-1)*B2)`; E2 `=IF(D2="","",COUNTIF(Orders!$E$5:$E$204,D2))`
  - B6 CHECK: `=IF(AND(Orders!B2="OK",B5>=1,SUMPRODUCT((Helpers!A4:A8<>"")*(ROW(Helpers!A4:A8)-ROW(Helpers!A4)+1<=B5))=B5,SUMPRODUCT(--(Helpers!A4:A8<>""),--(COUNTIF(Helpers!A4:A8,Helpers!A4:A8)>1))=0,E7=B1,E8<=1,COUNTA(Orders!A205:A1000)=0,MAX(Orders!C5:C204)=B3),"OK","CHECK")`
- Example table values, sample (45 orders, batch size 10, 3 helpers): B1 45, B2 10, B3 5, B4 5, B5 3, B6 OK; Ana 15, Ben 15, Cora 15, total 45, most minus fewest 0.
  | Order | Item no. | Batch | Position | Helper |
  | 4101 | 1 | 1 | 1 | Ana |
  | 4102 | 2 | 1 | 2 | Ben |
  | 4103 | 3 | 1 | 3 | Cora |
  | 4104 | 4 | 1 | 4 | Ana |
  | 4110 | 10 | 1 | 10 | Ana |
  | 4111 | 11 | 2 | 1 | Ben |
  | 4112 | 12 | 2 | 2 | Cora |
  | 4113 | 13 | 2 | 3 | Ana |
  | 4141 | 41 | 5 | 1 | Ben |
  | 4145 | 45 | 5 | 5 | Cora |
  The helper rotation does not restart at a new batch (item 11 goes to Ben because item 10 went to Ana).
- Open cases resolved: gaps -> numbering uses COUNTA, not ROW, so an empty row in the middle is skipped and numbering stays 1, 2, 3 (a cell holding only a space counts as an order); last batch smaller -> shown in B4; zero orders -> everything blank or 0, CHECK OK; deleted helper -> clear the cell and move the names up (a gap is caught by CHECK).
- Edge findings (every cell in rows 5 to 204 of each case matched `divmod`; counts per helper matched):
  - 45 orders, batch 7, 4 helpers: 7 batches, last batch 3, Ana 12, Ben 11, Cora 11, Dev 11 (most minus fewest 1, CHECK OK). 199 orders, batch 7, 3 helpers: 29 batches, last 3. 0 orders: 0 batches, last batch blank, helper counts 0, CHECK OK. 1 order, 9, 10, 11, 20, 100, 199 and 200 orders with batch sizes 1, 7, 10 and 1 to 5 helpers all matched (12 size/helper combinations). 200 orders fill rows 5 to 204.
  - Blank rows in the middle (3 blanks in 25 orders; 5 blanks in 12 orders): skipped, numbering continues, CHECK OK. Text, decimal and number order numbers mixed (1001, "A-12", 1003.5, "x y", 0, "007"): all counted.
  - Bad input (CHECK in Summary!B6): batch size 0, -3, 2.5 or "ten" -> B2 message, Batch and Position show "Check batch size" (helper still shows); a gap in the helper list (Ana, empty, Cora) -> some orders show blank helper and the counts are unfair (CHECK); a duplicate helper name (Ana, Ben, Ana) -> counts merge (CHECK); no helpers -> "No helper" in every row (CHECK); an order typed in A205 (below the table) -> not numbered, CHECK. Troubleshooting use: "a helper's count looks too high" = name listed twice; "blank helper" = empty row between names.
- Limits: splits and deals a list, it does not weigh how long an order takes. Not run: helper names with * or ? (COUNTIF wildcards), more than 200 orders.
