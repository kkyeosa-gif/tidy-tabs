# Handoff 2026-11-14: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-14.md (researcher) and tasks/hooks-2026-11-14.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`_digits_only`, `best_month_per_product`, `julian_date_lot_codes`, `match_payouts_to_deposits`, `split_dates_across_months`, `payment_plan_interest_rate`, registered after `variation_values_by_label()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-14 build and test"). Heroes are not made yet (designer, see the hero section below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 73 edited workbooks recalculated and compared (132,590 cell comparisons, 0 mismatches against independent Python models), plus 1 probe workbook of the short (core) formulas and failure cases read by hand; no error value in any recalculated workbook. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text. Bugs found and fixed are in tasks/briefs.md.

Repo note: only the five new builder functions were run against templates/ (imported and run with /usr/bin/python3 -B), so no older template was rewritten. The whole script was also run once from a scratch directory (exit 0, all files written there, including the five new ones), so a full re-run still works. `git status` shows scripts/build-templates.py, the five new .xlsx files, this handoff and tasks/briefs.md.

Post file names: `posts/ready/2026-11-14-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks Check and the other flagged results named below, green marks Matched / No interest, orange marks ties for best, Check date and days outside the grid. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, MINIFS/MAXIFS, dynamic arrays (none used; EOMONTH, EDATE and RATE are plain, no _xlfn in any file). Help links were not checked by the builder; the script-writer checks every link as usual.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-14-team-01-find-best-and-worst-month-per-product-column-header-of-max-index-match-excel-google-sheets | tidy-tabs-best-month-per-product.xlsx | 13 workbooks, 5,746 comparisons |
| 2 | 2026-11-14-team-02-convert-date-to-julian-date-code-yyddd-batch-lot-number-excel-google-sheets | tidy-tabs-julian-date-lot-codes.xlsx | 16 workbooks, 28,896 comparisons |
| 3 | 2026-11-14-team-03-match-payouts-to-bank-deposits-repeat-amounts-countif-match-excel-google-sheets | tidy-tabs-match-payouts-to-deposits.xlsx | 19 workbooks, 41,933 comparisons |
| 4 | 2026-11-14-team-04-split-booking-days-and-amount-across-months-eomonth-min-max-excel-google-sheets | tidy-tabs-split-dates-across-months.xlsx | 15 workbooks, 52,965 comparisons |
| 5 | 2026-11-14-team-05-find-interest-rate-of-payment-plan-offer-rate-function-excel-google-sheets | tidy-tabs-payment-plan-interest-rate.xlsx | 10 workbooks, 3,050 comparisons |

tested_in text pattern (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets."

Hero files: for the designer (not made by the builder). Key columns for a hero (real LibreOffice PDF render of the template with its sample data, 1600x900):
1. Best month: Products tab A (Product), B:K (Jan to Oct), O Best month, P Best sales, Q Months tied for best, R Worst month, S Worst sales, U Status; row 56 Top product is optional. L:M (Nov, Dec, blank), N, T and the summary W:X can be hidden in the hero.
2. Julian: Make codes tab A Made on, B Product prefix, E Julian code, F Lot code, G Status; a second crop of Read codes A Code as printed, E Date, F Status. D (Day of year) is worth keeping on Make codes.
3. Payouts: Payouts tab A Payout date, B Source, C Amount, E Copy no., H Bank date, I Days to bank, J Status; summary L4:M10. D, F, G (helpers) can be hidden.
4. Split: Amounts tab A Job, D Amount, G:K (Oct 2026 to Feb 2027) and S In the grid; or Days tab A, B, C, E, G:K, S. Month columns after Mar 2027 can be hidden for width.
5. RATE: Offers tab A Offer, B Cash price, E Monthly payment, F Number of payments, G First payment at signing, I Extra vs cash, K Rate x 12, L Effective yearly rate, M Status.

Every workbook has a "How to use" tab and a US Letter page setup with the header row repeated (data tabs landscape, fit to one page wide). Deviations from tasks/team-topics-2026-11-14.md are listed per section.

---
## 1. find-best-and-worst-month-per-product-column-header-of-max-index-match-excel-google-sheets
Title (from topics): How to Find Each Product's Best and Worst Month in Excel and Google Sheets
Template: templates/tidy-tabs-best-month-per-product.xlsx (tabs: Products, How to use)

Layout
- Products: A1:A2 notes. Header row 4 (month names typed as text), data rows 5:54 (8 sample products, 2026 sales Jan to Oct, Nov and Dec blank). A Product, B:M Jan to Dec (inputs, $); N Total, O Best month, P Best sales, Q Months tied for best, R Worst month, S Worst sales, T Months filled in, U Status (formulas). Row 56 Top product per month, row 57 its sales, row 58 Month total. Summary W4:X9: Products listed (X4), Best month, all products (X5), Products with a tie for best (X6), Rows showing Check (X7), Grand total (X8), CHECK (X9). U red for Check, Q orange when above 1. Landscape.
- Deviations: T is "Months filled in" (COUNT, a 0 month counts) instead of "Months with sales", because a $0.00 month is not a month with sales; row 58 Month total added (X5 uses it); summary has Grand total as its own row.

Key formulas (row 5)
- O5 `=IF(AND(TRIM(A5)="",COUNTA(B5:M5)=0),"",IF(COUNTA(B5:M5)<>COUNT(B5:M5),"Check",IF(COUNT(B5:M5)=0,"No sales",INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0)))))` (core `=INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0))`)
- R5 same as O5 with MIN (core `=INDEX($B$4:$M$4,MATCH(MIN(B5:M5),B5:M5,0))`)
- N5 `=IF(AND(TRIM(A5)="",COUNTA(B5:M5)=0),"",IF(COUNTA(B5:M5)<>COUNT(B5:M5),"Check",SUM(B5:M5)))`
- U5 `=IF(AND(TRIM(A5)="",COUNTA(B5:M5)=0),"",IF(OR(TRIM(A5)="",COUNTA(B5:M5)<>COUNT(B5:M5)),"Check",IF(COUNT(B5:M5)=0,"No sales","OK")))`
- P5 `=IF(U5="OK",MAX(B5:M5),"")`; S5 `=IF(U5="OK",MIN(B5:M5),"")`; Q5 `=IF(U5="OK",COUNTIF(B5:M5,MAX(B5:M5)),"")` (core `=COUNTIF(B5:M5,MAX(B5:M5))`); T5 `=IF(AND(TRIM(A5)="",COUNTA(B5:M5)=0),"",COUNT(B5:M5))`
- B56 `=IF(COUNTA(B5:B54)<>COUNT(B5:B54),"Check",IF(COUNT(B5:B54)=0,"No sales",INDEX($A$5:$A$54,MATCH(MAX(B5:B54),B5:B54,0))))` (core `=INDEX($A$5:$A$54,MATCH(MAX(B5:B54),B5:B54,0))`); B57 `=IF(ISNUMBER(B58),MAX(B5:B54),"")`; B58 `=IF(COUNTA(B5:B54)<>COUNT(B5:B54),"Check",IF(COUNT(B5:B54)=0,"",SUM(B5:B54)))`
- X4 `=SUMPRODUCT(--(U5:U54<>""))`; X5 `=IF(COUNT(B58:M58)=0,"No sales",INDEX(B4:M4,MATCH(MAX(B58:M58),B58:M58,0)))`; X6 `=COUNTIF(Q5:Q54,">1")`; X7 `=COUNTIF(U5:U54,"Check")`; X8 `=SUM(B5:M54)`; X9 `=IF(AND(X7=0,ROUND(SUM(N5:N54)-X8,2)=0),"OK","Check")`
- O and R are formatted mmm, so if a reader types real dates in row 4 the result still shows Jan, Feb... (tested).

Sample values (fictional; recalculated). Sales Jan to Oct:
| Product | Jan to Oct | Total | Best | Best sales | Tied | Worst | Worst sales | Months | Status |
|---|---|---|---|---|---|---|---|---|---|
| Soy candle 8 oz | 420, 380, 455, 510, 540, 495, 470, 525, 610, 890 | $5,295.00 | Oct | $890.00 | 1 | Feb | $380.00 | 10 | OK |
| Wax melts | 180, 165, 210, 240, 275, 310, 310, 260, 225, 290 | $2,465.00 | Jun | $310.00 | 2 | Feb | $165.00 | 10 | OK |
| Holiday ornament | 0 x 8, 215, 640 | $855.00 | Oct | $640.00 | 1 | Jan | $0.00 | 10 | OK |
| Gift box | blank Jan to Jun, 340, 280, 365, 720 | $1,705.00 | Oct | $720.00 | 1 | Aug | $280.00 | 4 | OK |
| Market tote | 95, -25, 60, 120, 145, 160, 150, 130, 110, 175 | $1,120.00 | Oct | $175.00 | 1 | Feb | -$25.00 | 10 | OK |
| Lip balm | 140, 150, 160, 175, 190, 205, 235, 200, 185, 180 | $1,820.00 | Jul | $235.00 | 1 | Jan | $140.00 | 10 | OK |
| Linen tea towel | 260, 240, 480, 530, 565, 340, 320, 345, 430, 410 | $3,920.00 | May | $565.00 | 1 | Feb | $240.00 | 10 | OK |
| Discontinued soap | all blank | $0.00 | No sales | | | No sales | | 0 | No sales |
Row 56 Top product: Jan Soy candle 8 oz ($420.00), Feb Soy candle 8 oz ($380.00), Mar, Apr, May Linen tea towel ($480.00, $530.00, $565.00), Jun to Oct Soy candle 8 oz ($495.00, $470.00, $525.00, $610.00, $890.00), Nov and Dec No sales. Row 58 Month total Jan to Oct: $1,095.00, $910.00, $1,365.00, $1,575.00, $1,715.00, $1,510.00, $1,825.00, $1,740.00, $2,140.00, $3,305.00.
Summary X4:X9: 8, Oct, 1, 0, $17,180.00, OK. Quotable: Wax melts ties Jun and Jul at $310.00 (Q = 2, the first from the left wins); Holiday ornament's worst month is Jan at $0.00 because a 0 counts (eight months are tied at $0.00; the file has no tie count for worst); Gift box's blank Jan to Jun are skipped, so its worst month is Aug ($280.00).

Edge cases tested (all matched the model)
- Ties for best (three at 9: first wins, Q = 3) and for worst; 0 vs blank (0 counts, blank skipped); all negative months; one month only (best and worst the same month); only December filled; all twelve months 0 (Jan, tie 12); decimals 12.35 vs 12.345; 1,000,000.01 vs 999,999.99.
- "$120" typed as text: N, O, R show Check, U Check (the core formula skips it silently: probe below). A cell holding one space: Check.
- A product name missing but sales typed: O and R still show months, U shows Check. A product name that is a number (12345): works. A name of spaces only with no sales: blank row.
- =0.1+0.2 vs 0.3 in the same row: LibreOffice treats them as equal (Q = 2, and `=B4=C4` gave TRUE in the probe).
- Row 4 as real dates (01/01/2026 ...): best and worst still show the month. All 50 rows blank: every formula blank, X5 No sales, CHECK OK. 50 products x 12 months. Product on the last row (54). 8 random grids of 50 rows (blanks, 0, negatives, text in half the sets).

Pitfalls the post must state
- MAX and MIN skip text without a warning: a sale typed as "$120" (text) is ignored by the core formula. The file shows Check; the post should say to check that sales are numbers.
- A 0 is a month with $0.00 sales and can be the worst month; leave a month blank if the product was not for sale.
- With a tie, MATCH returns the first month from the left; Q shows how many months tie.
- An all-blank row gives #N/A with the core formula (probe), No sales in the file.
- Not tested: Excel, Google Sheets (COUNTIF with a decimal MAX criterion, the mmm format on text, and COUNTA on cells with formulas returning "" are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook)
- #N/A for a product with no sales: the core INDEX/MATCH on an all-blank row returned #N/A.
- Best month ignores a big month: the core formula on 100, "$120" (text), 90 returned Jan; the $120 was skipped.
- Two months tie but only one shows: MATCH returns the first (Wax melts Jun, not Jul).
- Worst month is a month with no sales: a 0 was typed (Holiday ornament Jan).

---
## 2. convert-date-to-julian-date-code-yyddd-batch-lot-number-excel-google-sheets
Title (from topics): How to Convert a Date to a Julian Date Code for Batch Lot Numbers in Excel and Google Sheets
Template: templates/tidy-tabs-julian-date-lot-codes.xlsx (tabs: Make codes, Read codes, How to use)

Layout
- Make codes: A1 note. Header row 3, rows 4:203 (10 sample batches). A Made on, B Product prefix, C Batch of the day (inputs); D Day of year, E Julian code YYDDD, F Lot code, G Status. Summary I3:J5: Rows listed (J3), Codes made (OK) (J4), Rows showing Check (J5).
- Read codes: A1 note. Header row 3, rows 4:203 (9 sample codes). A Code as printed (input); B Cleaned, C Year, D Day, E Date, F Status. Summary H3:I5: Codes listed (I3), Codes read (OK) (I4), Rows showing Check (I5). G and F red for Check. Landscape.
- Deviations: none in columns; Make codes has a 10th row with a text date to show Check; a blank Batch of the day gives a lot code without the suffix (CND-26318).

Key formulas (row 4)
- Make D4 `=IF(ISNUMBER(A4),IF(A4>=1,INT(A4)-DATE(YEAR(A4),1,0),""),"")` (core `=A4-DATE(YEAR(A4),1,0)`; INT so a date with a time does not round up the day)
- Make E4 `=IF(A4="","",IF(ISNUMBER(D4),IF(AND(YEAR(A4)>=2000,YEAR(A4)<=2099),RIGHT(YEAR(A4),2)&TEXT(D4,"000"),"Check"),"Check"))` (core `=RIGHT(YEAR(A4),2)&TEXT(A4-DATE(YEAR(A4),1,0),"000")`)
- Make F4 `=IF(OR(E4="",E4="Check"),"",IF(TRIM(B4)="","Check",IF(C4="",TRIM(B4)&"-"&E4,IF(NOT(ISNUMBER(C4)),"Check",IF(OR(C4<1,C4<>INT(C4)),"Check",TRIM(B4)&"-"&E4&"-"&C4)))))` (core `=B4&"-"&E4&"-"&C4`)
- Make G4 `=IF(AND(A4="",TRIM(B4)="",C4=""),"",IF(OR(E4="",E4="Check",F4="Check"),"Check","OK"))`
- Read B4 `=IF(ISNUMBER(A4),IF(AND(A4>=0,A4=INT(A4),A4<10000000),TEXT(A4,IF(A4<100000,"00000","0000000")),"Check"),TRIM(A4))`
- Read C4 `=IF(B4="","",IF(AND(OR(LEN(B4)=5,LEN(B4)=7),SUMPRODUCT(LEN(B4)-LEN(SUBSTITUTE(B4,{"0","1","2","3","4","5","6","7","8","9"},"")))=LEN(B4)),IF(LEN(B4)=5,2000+LEFT(B4,2),--LEFT(B4,4)),"Check"))`
- Read D4 `=IF(ISNUMBER(C4),--RIGHT(B4,3),"")`
- Read E4 `=IF(ISNUMBER(D4),IF(C4<1900,"Check",IF(AND(D4>=1,D4<=DATE(C4,12,31)-DATE(C4,1,0)),DATE(C4,1,D4),"Check")),IF(B4="","","Check"))` (core `=DATE(2000+LEFT(A4,2),1,RIGHT(A4,3))`)
- Read F4 `=IF(B4="","",IF(ISNUMBER(E4),"OK","Check"))`
- Summaries: `=SUMPRODUCT(--(G4:G203<>""))`, `=COUNTIF(G4:G203,"OK")`, `=COUNTIF(G4:G203,"Check")` (F on Read codes).

Sample values (fictional; recalculated)
| Made on | Prefix | Batch | Day | Code | Lot code | Status |
|---|---|---|---|---|---|---|
| 11/14/2026 | CND | 1 | 318 | 26318 | CND-26318-1 | OK |
| 11/14/2026 | CND | 2 | 318 | 26318 | CND-26318-2 | OK |
| 01/01/2026 | SOP | 1 | 1 | 26001 | SOP-26001-1 | OK |
| 12/31/2026 | JAM | 1 | 365 | 26365 | JAM-26365-1 | OK |
| 02/29/2028 | CND | 1 | 60 | 28060 | CND-28060-1 | OK |
| 03/01/2027 | SOP | 1 | 60 | 27060 | SOP-27060-1 | OK |
| 12/31/2028 | JAM | 1 | 366 | 28366 | JAM-28366-1 | OK |
| 02/01/2005 | SOP | 1 | 32 | 05032 | SOP-05032-1 | OK |
| 11/13/2026 | JAM | 2 | 317 | 26317 | JAM-26317-2 | OK |
| "11/15/2026" (text) | CND | 1 | | Check | | Check |
Make summary J3:J5: 10, 9, 1.
| Code as printed | Cleaned | Year | Day | Date | Status |
|---|---|---|---|---|---|
| "26318" | 26318 | 2026 | 318 | 11/14/2026 | OK |
| " 26318 " | 26318 | 2026 | 318 | 11/14/2026 | OK |
| "2026032" | 2026032 | 2026 | 32 | 02/01/2026 | OK |
| 28366 (number) | 28366 | 2028 | 366 | 12/31/2028 | OK |
| 5032 (number, zero lost) | 05032 | 2005 | 32 | 02/01/2005 | OK |
| "05032" | 05032 | 2005 | 32 | 02/01/2005 | OK |
| "26366" | 26366 | 2026 | 366 | Check | Check |
| "26000" | 26000 | 2026 | 0 | Check | Check |
| "26A18" | 26A18 | Check | | Check | Check |
Read summary I3:I5: 9, 6, 3.

Edge cases tested (all matched the model: Python date.timetuple().tm_yday and calendar.isleap)
- Every day of 2026, 2027 and 2028 made into codes, and the same codes read back as 5-digit text, 5-digit numbers and 7-digit text (round trip, 1,096 days).
- Make: 01/01/2000 (00001), 12/31/2009 (09365), 12/31/2099 (99365), 03/01/2100 and 12/31/1999 (Check, outside 2000 to 2099), dates with 9:00 AM and 11:59:59 PM (same day number, not rounded up), text date, no date with a prefix, no prefix, prefix with spaces (trimmed), no batch (CND-26318), batch 0, 1.5, "1" as text, -1 (Check), spaces only in A, serials 0, -5, 0.5 (Check), 12/31/2024 (24366), 02/29/2028 (28060) vs 03/01/2028 (28061) vs 03/01/2027 (27060), batch 10.
- Read: "26 318", "26A18", "ZZZZZ", "abc", "26.318", "-26318", -26318, 26318.5, "26318.0", full-width digits, 4, 6 and 8 digit text, 263180 and 26318000 as numbers, 10000000 (Check); 0 and "00000" (day 000, Check); "00001", 1 and 1001 as numbers (01/01/2000, 01/01/2001); "2100060" (03/01/2100) and "2100366" (Check, 2100 is not a leap year); "2028366" OK, "2027366" Check; "1999365" OK, "1899365" and "0001001" Check; "99365", "00366" (2000 is a leap year), "04366"; "28366 " with a trailing space; one space (blank row); "5032" as 4-digit text (Check: the file only pads numbers).
- 8 random sets of 200 make rows and 200 codes (2000 to 2099, times, text dates, bad prefixes and batches, codes as text, numbers, 7 digits, spaces, day 000 to 368, junk).

Pitfalls the post must state
- The code needs the day as three digits: TEXT(...,"000"). Without it 02/01/2005 gives 0532 (probe: `=RIGHT(YEAR(A18),2)&(A18-DATE(YEAR(A18),1,0))` returned "0532").
- 2000 to 2009 codes lose their leading zero when typed or imported as numbers (05032 becomes 5032). The file puts it back for numbers; a 4-digit text code shows Check.
- Leap years shift every day after February 28 by one (03/01 is day 60 in 2027 and day 61 in 2028).
- DATE alone does not complain about impossible days: 26366 becomes 01/01/2027 and 26000 becomes 12/31/2025 (probe); the file shows Check.
- YYDDD cannot tell 1926 from 2026: the file makes codes for 2000 to 2099 only. 7-digit codes (YYYYDDD) are read for 1900 on.
- TEXT(A4,"yy") also gave 26 in the LibreOffice probe, but the file uses RIGHT(YEAR()) because format codes can change with the language setting (not tested in other languages).
- This is not the astronomical Julian Day number. The file does not cover any lot code rule (FDA or other); the post says that and links nothing about rules unless checked.
- Not tested: Excel, Google Sheets (TEXT "000", SUMPRODUCT with an array constant inside SUBSTITUTE, -- on text, DATE with day 0 are the open questions), a non-US language setting.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook, core read formula `=DATE(2000+LEFT(A,2),1,RIGHT(A,3))`)
- 7-digit code read wrong: "2026032" gave 02/01/2020.
- Code typed as a number lost its zero: 5032 gave 02/01/2050.
- Code with spaces: " 26318 " gave 01/18/2002.
- Impossible day: "26366" gave 01/01/2027, "26000" gave 12/31/2025 (no error).

---
## 3. match-payouts-to-bank-deposits-repeat-amounts-countif-match-excel-google-sheets
Title (from topics): How to Match Payouts to Bank Deposits When the Same Amount Repeats in Excel and Google Sheets
Template: templates/tidy-tabs-match-payouts-to-deposits.xlsx (tabs: Payouts, Bank, How to use)

Layout
- Payouts: A1 "Days allowed", B1 3 (input), C1 note. Header row 4, rows 5:204 (10 sample payouts). A Payout date, B Source, C Amount (inputs); D Cents, E Copy no., F Match key, G Bank row, H Bank date, I Days to bank, J Status (Matched green, Check date orange, Not in bank and Check red). Summary L4:M10: Payouts matched (M4), Payouts to check (date) (M5), Payouts not in bank (M6), Found in the bank ($) (M7, matched + check date), Not in bank ($) (M8), Deposits that are not payouts ($) (M9), CHECK (M10). Landscape.
- Bank: A1 note. Header row 4, rows 5:204 (10 sample deposits). A Date, B Description, C Amount (inputs); D Cents, E Copy no., F Match key, G Status (Matched green, Not a payout and Check red).
- Deviations: H shows Check when the matched deposit has no real date (then J says Check date); D shows Check for fractions of a cent.

Key formulas (row 5)
- D5 `=IF(AND(A5="",TRIM(B5)="",C5=""),"",IF(ISNUMBER(C5),IF(ABS(C5*100-ROUND(C5*100,0))>0.0001,"Check",ROUND(C5*100,0)),"Check"))` (core `=ROUND(C5*100,0)`; same on Bank)
- E5 `=IF(ISNUMBER(D5),COUNTIF($D$5:D5,D5),"")` (core `=COUNTIF($D$5:D5,D5)`; same on Bank)
- F5 `=IF(ISNUMBER(D5),D5&"-"&E5,"")` (core `=D5&"-"&E5`; same on Bank)
- G5 `=IF(F5="","",IFERROR(MATCH(F5,Bank!$F$5:$F$204,0),""))`
- H5 `=IF(ISNUMBER(G5),IF(ISNUMBER(INDEX(Bank!$A$5:$A$204,G5)),INDEX(Bank!$A$5:$A$204,G5),"Check"),"")`
- I5 `=IF(AND(ISNUMBER(H5),ISNUMBER(A5)),INT(H5)-INT(A5),"")`
- J5 `=IF(D5="","",IF(OR(D5="Check",NOT(ISNUMBER(A5)),NOT(ISNUMBER($B$1))),"Check",IF($B$1<0,"Check",IF(G5="","Not in bank",IF(NOT(ISNUMBER(I5)),"Check date",IF(OR(I5<0,I5>$B$1),"Check date","Matched"))))))`
- Bank G5 `=IF(D5="","",IF(OR(D5="Check",NOT(ISNUMBER(A5))),"Check",IF(ISNA(MATCH(F5,Payouts!$F$5:$F$204,0)),"Not a payout","Matched")))`
- One-cell key the post can teach: `=ROUND(C5*100,0)&"-"&COUNTIF($C$5:C5,C5)`, then `=MATCH(key, other list's keys, 0)`.
- M4 `=COUNTIF(J5:J204,"Matched")`; M5 `=COUNTIF(J5:J204,"Check date")`; M6 `=COUNTIF(J5:J204,"Not in bank")`; M7 `=SUMIFS(C5:C204,J5:J204,"Matched")+SUMIFS(C5:C204,J5:J204,"Check date")`; M8 `=SUMIFS(C5:C204,J5:J204,"Not in bank")`; M9 `=SUMIFS(Bank!C5:C204,Bank!G5:G204,"Not a payout")`; M10 `=IF(COUNTIF(J5:J204,"Check")+COUNTIF(Bank!G5:G204,"Check")=0,"OK","Check")`

Sample values (fictional; B1 = 3; recalculated)
| Payout date | Source | Amount | Copy | Key | Bank row | Bank date | Days | Status |
|---|---|---|---|---|---|---|---|---|
| 11/02/2026 | Square | $412.50 | 1 | 41250-1 | 1 | 11/03/2026 | 1 | Matched |
| 11/02/2026 | Etsy | $286.40 | 1 | 28640-1 | 2 | 11/04/2026 | 2 | Matched |
| 11/03/2026 | Stripe | $318.60 | 1 | 31860-1 | 4 | 11/05/2026 | 2 | Matched |
| 11/04/2026 | Square | $412.50 | 2 | 41250-2 | 3 | 11/05/2026 | 1 | Matched |
| 11/05/2026 | Stripe | $1,150.00 | 1 | 115000-1 | 6 | 11/09/2026 | 4 | Check date |
| 11/06/2026 | Square | $98.75 | 1 | 9875-1 | 7 | 11/09/2026 | 3 | Matched |
| 11/09/2026 | Etsy | $199.99 | 1 | 19999-1 | | | | Not in bank |
| 11/10/2026 | Stripe | $75.00 | 1 | 7500-1 | | | | Not in bank |
| 11/10/2026 | Square | $498.00 | 1 | 49800-1 | 9 | 11/11/2026 | 1 | Matched |
| 11/11/2026 | Square | $412.50 | 3 | 41250-3 | 10 | 11/12/2026 | 1 | Matched |
Bank (date, description, amount, copy, status): 11/03 Square transfer $412.50 1 Matched; 11/04 Etsy deposit $286.40 1 Matched; 11/05 Square transfer $412.50 2 Matched; 11/05 Stripe transfer $318.60 1 Matched; 11/06 Cash deposit, market $640.00 1 Not a payout; 11/09 Stripe transfer $1,150.00 1 Matched; 11/09 Square transfer $98.75 1 Matched; 11/10 Etsy deposit $199.90 1 Not a payout; 11/11 Square transfer $498.00 1 Matched; 11/12 Square transfer $412.50 3 Matched.
Summary M4:M10: 7, 1, 2, $3,589.25, $274.99, $839.90, OK. Quotable: the three $412.50 payouts (11/02, 11/04, 11/11) pair with the three $412.50 deposits in order, 1 day each; the Friday 11/06 $98.75 lands Monday 11/09 (3 days, Matched); Stripe $1,150.00 takes 4 days (Check date); the $199.99 Etsy payout and the $199.90 Etsy deposit do not pair (9 cents apart).
Days allowed changes (same sample): 0 gives 0 matched, 8 check date; 1 gives 4 and 4; 2.5 gives 6 and 2; 4 gives 8 and 0; blank, "3" as text or -1 make every payout Check and CHECK Check.

Edge cases tested (all matched a Python model that pairs the k-th copy of each cent amount)
- The same amount twice on one day (pairs 1-1, 2-2); a deposit dated before its payout (-1 day, Check date); both lists swapped the same way (still Matched); only the payout list out of date order: 412.50 on 11/09 then 11/02 against deposits 11/03 and 11/10 gives -6 and 8 days, both Check date.
- 0.1+0.2 and 0.30 (both 30 cents, pair); 12.345 (Check, fractions of a cent) vs a 12.35 deposit (Not a payout); "$412.50" as text (Check) vs a real 412.50 deposit (Not a payout); "99.99" text deposit (Check); -$25.00 refund on both lists (Matched); $0.00 (Matched); 0.0000001 (counts as 0 cents); 1,234,567.89.
- Blank rows inside both lists, a text payout date and a missing payout date (Check, the deposits they use are still taken), a missing source (fine), date with a time (11:30 PM to 8:00 AM next day = 1 day), a deposit with a text date or no date (payout Check date, Bank Check), a payout with no amount (Check).
- 200 payouts and 200 deposits all $25.00 (copies 1 to 200 pair in order). 8 random sets of up to 160 payouts and 200 deposits (repeat amounts, lags -1 to 5 days, missing deposits, 1-cent differences, extra deposits, sorted and partly shuffled bank lists, B1 2, 3 or 5).

Pitfalls the post must state
- Copies pair in the order listed. Keep both lists in date order (oldest first); a list out of order shows Check date rather than a silent wrong pair.
- Amounts must match to the cent; a payout that lands a few cents short (a fee taken) shows Not in bank and the deposit Not a payout. The file does not guess.
- Days allowed is calendar days (a Friday payout landing Monday is 3).
- The core MATCH on the amount alone pairs every repeat with the first deposit (probe below).
- No accounting judgment: the file only says which rows to look at; the post does not say how to book anything.
- Not tested: Excel, Google Sheets (COUNTIF with a growing range, MATCH on text keys, ROUND of C*100 are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Every repeat matches the same deposit: `=MATCH(A20,$B$20:$B$22,0)` on three $412.50 rows returned 1 for all three (probe); the copy number key gives 1, 2, 3.
- A payout shows Check date though it arrived: it took longer than Days allowed (Stripe $1,150.00, 4 days), or the lists are not in the same order (pay_swapped test).
- Not in bank and Not a payout for amounts a few cents apart ($199.99 vs $199.90).
- A row shows Check: the amount is text ("$412.50") or has fractions of a cent (12.345), or a date is text or blank, or B1 is not a number of 0 or more.

---
## 4. split-booking-days-and-amount-across-months-eomonth-min-max-excel-google-sheets
Title (from topics): How to Split a Booking or Contract Across Months by Days in Excel and Google Sheets
Template: templates/tidy-tabs-split-dates-across-months.xlsx (tabs: Days, Amounts, How to use)

Layout
- Days: A1 "First month of the grid", B1 10/01/2026 (input), C1 note. Header row 4: A Job, B Start, C End, D Amount (inputs), E Days, F Days outside the grid, G4:R4 the 12 months (formulas, mmm yyyy, Oct 2026 to Sep 2027), S Status (OK, "N days outside" orange, Check red). Rows 5:104 (6 sample jobs). Row 105 Total (F and G:R). Summary U4:V9: Jobs listed (V4), Rows showing Check (V5), Jobs with days outside the grid (V6), Dollars in the grid (V7), Dollars outside the grid (V8), CHECK (V9). Landscape.
- Amounts: A1 note. Same rows, all formulas: A:D copied from Days, E Days, F Outside the grid ($), G:R dollars per month, S In the grid ($), T Adds up (OK or Check). Row 105 totals.
- Deviations: G4 is `DATE(YEAR(B1),MONTH(B1),1)` (the topic's alternative), so any date in the month works; the summary sits on the Days tab; Amounts has the extra columns E, F, S, T; amounts with fractions of a cent are Check; the running-share formula has one more outer ROUND.

Key formulas (row 5)
- Days G4 `=IF(ISNUMBER($B$1),DATE(YEAR($B$1),MONTH($B$1),1),"Check")`; H4 `=IF(ISNUMBER(G4),EDATE(G4,1),"")` (core `=EDATE(G4,1)`)
- E5 `=IF(AND(TRIM(A5)="",B5="",C5="",D5=""),"",IF(AND(ISNUMBER(B5),ISNUMBER(C5)),INT(C5)-INT(B5)+1,"Check"))` (core `=C5-B5+1`)
- G5 `=IF(AND(ISNUMBER($E5),ISNUMBER(G$4)),IF($E5>0,MAX(0,MIN(INT($C5),EOMONTH(G$4,0))-MAX(INT($B5),G$4)+1),""),"")` (core `=MAX(0,MIN($C5,EOMONTH(G$4,0))-MAX($B5,G$4)+1)`)
- F5 `=IF(ISNUMBER(E5),IF(E5>0,E5-SUM(G5:R5),""),"")`
- S5 `=IF(E5="","",IF(OR(E5="Check",NOT(ISNUMBER($G$4))),"Check",IF(E5<1,"Check",IF(NOT(ISNUMBER(D5)),"Check",IF(OR(D5<0,ABS(D5*100-ROUND(D5*100,0))>0.0001),"Check",IF(F5>0,F5&" days outside","OK"))))))`
- Amounts G5 `=IF(AND(ISNUMBER(Days!G5),OR(Days!$S5="OK",RIGHT(Days!$S5,7)="outside")),ROUND($D5*Days!G5/$E5,2),"")` (core `=ROUND($D5*Days!G5/$E5,2)`)
- Amounts H5 `=IF(AND(ISNUMBER(Days!H5),OR(Days!$S5="OK",RIGHT(Days!$S5,7)="outside")),ROUND(ROUND($D5*SUM(Days!$G5:H5)/$E5,2)-SUM($G5:G5),2),"")` (core `=ROUND($D5*SUM(Days!$G5:H5)/$E5,2)-SUM($G5:G5)`)
- Amounts S5 `=IF(COUNT(G5:R5)=0,"",SUM(G5:R5))`; F5 `=IF(ISNUMBER(S5),ROUND(D5-S5,2),"")`; T5 `=IF(Days!S5="OK",IF(ROUND(S5-D5,2)=0,"OK","Check"),IF(Days!S5="Check","Check",""))`; A5 `=IF(Days!A5="","",Days!A5)` (B:D the same); E5 `=IF(ISNUMBER(Days!E5),Days!E5,"")`
- V4 `=SUMPRODUCT(--(S5:S104<>""))`; V5 `=COUNTIF(S5:S104,"Check")`; V6 `=COUNTIF(S5:S104,"*outside")`; V7 `=SUM(Amounts!S5:S104)`; V8 `=SUM(Amounts!F5:F104)`; V9 `=IF(AND(V5=0,COUNTIF(Amounts!T5:T104,"Check")=0),"OK","Check")`

Sample values (fictional; B1 = 10/01/2026; recalculated)
| Job | Start | End | Amount | Days | Days per month (Oct 26 to Sep 27, non-zero) | Dollars per month | Status |
|---|---|---|---|---|---|---|---|
| Booth season | 10/15/2026 | 01/14/2027 | $1,200.00 | 92 | Oct 17, Nov 30, Dec 31, Jan 14 | $221.74, $391.30, $404.35, $182.61 (= $1,200.00) | OK |
| Studio sublet | 12/20/2026 | 02/19/2027 | $750.00 | 62 | Dec 12, Jan 31, Feb 19 | $145.16, $375.00, $229.84 (= $750.00) | OK |
| Tent rental | 11/14/2026 | 11/14/2026 | $85.00 | 1 | Nov 1 | $85.00 | OK |
| Site care retainer | 01/01/2027 | 12/31/2027 | $1,800.00 | 365 | Jan 31, Feb 28, Mar 31, Apr 30, May 31, Jun 30, Jul 31, Aug 31, Sep 30; 92 outside | $152.88, $138.08, $152.88, $147.94, $152.88, $147.94, $152.88, $152.88, $147.94 (= $1,346.30); $453.70 outside | 92 days outside |
| Window display | 12/05/2026 | 12/01/2026 | $300.00 | -3 | | | Check |
| Pop-up shop | 02/01/2027 | 03/31/2027 | $590.00 | 59 | Feb 28, Mar 31 | $280.00, $310.00 | OK |
Summary V4:V9: 6, 1, 1, $3,971.30, $453.70, Check. CHECK shows Check on purpose because of the Window display row (end before start); fixing or deleting it turns CHECK to OK. Amounts row 105 (Oct 2026 to Sep 2027): $221.74, $476.30, $549.51, $710.49, $647.92, $462.88, $147.94, $152.88, $147.94, $152.88, $152.88, $147.94; total in the grid $3,971.30. Days row 105: 17, 31, 43, 76, 75, 62, 30, 31, 30, 31, 31, 30.

Edge cases tested (all matched a Python model with exact fractions and the same running-share rounding)
- Same month (11/03 to 11/20, 18 days); first to last day of a month; 10/31 to 12/01 (1 + 30 + 1 days, $3.13 / $93.75 / $3.12 = $100.00); 01/30/2027 to 03/02/2027 across February 2027 (28 days) and 02/01/2028 to 03/01/2028 (29 + 1 days, leap year, outside this grid); starting before the grid (47 days outside), ending after it (10 outside), both sides (61 outside), entirely outside (60 outside, $0.00 in the grid).
- Dates with times (10/30 9:00 AM to 11/02 5:00 PM = 4 days), text start date, blank end, end before start (-3, Check), amount 0 (all $0.00, OK), negative, text "$100.00" and blank amount (Check), 100.005 (Check, fractions of a cent), $100.00 over 92 days ($33.70 / $32.60 / $33.70, adds to $100.00), $1,000.01 over 363 days, $0.01 over a year (the cent lands in one month), $123,456.78, a job with no name (works), a name with no dates (Check).
- B1 on 10/17/2026 (grid still starts Oct 2026), B1 with a time on 09/30/2026 (Sep 2026 grid), B1 as text or blank (every row Check), grid starting Jun 2027.
- 100 jobs filled; 8 random sets of 100 jobs (lengths 0 to 403 days, times, end before start, amounts 0 to $5,000.00, random grid start).

Pitfalls the post must state
- Both end dates count (+1): 10/15 to 01/14 is 92 days, not 91.
- Rounding each month on its own can miss by a cent: the probe `=ROUND(100*31/92,2)` etc. gave $33.70 + $32.61 + $33.70 = $100.01. The running-share formula gives $33.70 + $32.60 + $33.70 = $100.00.
- Days outside the grid are counted and flagged, never dropped silently.
- B1 can be any day of the first month; the grid starts on the 1st.
- It is a planning view; the post must not call it revenue recognition or say how to book income or expenses (gate 3).
- Not tested: Excel, Google Sheets (EOMONTH and EDATE on formula dates, COUNTIF with "*outside", INT on dates with times are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Months add up to $100.01: rounding each month separately (probe above).
- A month is off by one day: leaving out the +1 (92 vs 91 days).
- A negative day count: end date before start (Window display, -3; the file shows Check).
- Months total less than the amount: the job runs past the grid (Site care retainer, 92 days and $453.70 outside).
- Every row shows Check: B1 is text or blank.

---
## 5. find-interest-rate-of-payment-plan-offer-rate-function-excel-google-sheets
Title (from topics): How to Find the Interest Rate of a Monthly Payment Plan With RATE in Excel and Google Sheets
Template: templates/tidy-tabs-payment-plan-interest-rate.xlsx (tabs: Offers, How to use)

Layout
- Offers: A1:A2 notes (A2: the rate is what the payments imply, not the lender's disclosed APR, no advice on whether to finance). Header row 4, rows 5:54 (8 sample offers). A Offer, B Cash price, C Down payment, D Upfront fee, E Monthly payment, F Number of payments, G First payment at signing (Yes/No) (inputs); H Total paid, I Extra vs cash, J Monthly rate (0.0000%), K Rate x 12, L Effective yearly rate (0.00%), M Status (OK, No interest green, Costs less than cash, Check red). Summary O4:P8: Offers listed (P4), Lowest yearly rate with interest (P5, offer name), Its effective yearly rate (P6), No-interest offers (P7), Rows showing Check (P8). Landscape.
- Deviations: J returns 0 directly when the payments add up to the cash price (RATE itself returned 1.7E-10 for 12 x $100.00 on $1,200.00 in the probe); a single payment due at signing is Check (nothing is financed; RATE returned -0.999999999999998 there in the probe instead of an error).

Key formulas (row 5)
- H5 `=IF(AND(TRIM(A5)="",B5="",C5="",D5="",E5="",F5="",TRIM(G5)=""),"",IF(AND(ISNUMBER(B5),ISNUMBER(E5),ISNUMBER(F5),OR(C5="",ISNUMBER(C5)),OR(D5="",ISNUMBER(D5)),OR(TRIM(G5)="",TRIM(G5)="Yes",TRIM(G5)="No")),IF(AND(B5>0,E5>0,F5>=1,F5=INT(F5),C5>=0,D5>=0,B5-C5-D5>0),C5+D5+E5*F5,"Check"),"Check"))` (core `=C5+D5+E5*F5`)
- I5 `=IF(ISNUMBER(H5),H5-B5,"")`
- J5 `=IF(ISNUMBER(H5),IF(ROUND(H5-B5,2)=0,0,IF(AND(F5=1,TRIM(G5)="Yes"),"Check",IFERROR(RATE(F5,-E5,B5-C5-D5,0,IF(TRIM(G5)="Yes",1,0)),"Check"))),IF(H5="","","Check"))` (cores `=RATE(F5,-E5,B5-C5)` and with fee and timing `=RATE(F5,-E5,B5-C5-D5,0,IF(G5="Yes",1,0))`)
- K5 `=IF(ISNUMBER(J5),J5*12,"")`; L5 `=IF(ISNUMBER(J5),(1+J5)^12-1,"")`; M5 `=IF(H5="","",IF(NOT(ISNUMBER(J5)),"Check",IF(J5=0,"No interest",IF(J5<0,"Costs less than cash","OK"))))`
- P4 `=SUMPRODUCT(--(M5:M54<>""))`; P5 `=IF(COUNTIF(M5:M54,"OK")=0,"",INDEX(A5:A54,MATCH(P6,L5:L54,0)))`; P6 `=IF(COUNTIF(M5:M54,"OK")=0,"",SMALL(L5:L54,COUNTIF(L5:L54,"<=0")+1))`; P7 `=COUNTIF(M5:M54,"No interest")`; P8 `=COUNTIF(M5:M54,"Check")`

Sample values (fictional; recalculated)
| Offer | Price | Down | Fee | Payment | N | At signing | Total paid | Extra | Monthly | x 12 | Effective | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Kiln, 24 months | $2,400.00 | | | $109.00 | 24 | No | $2,616.00 | $216.00 | 0.7012% | 8.41% | 8.75% | OK |
| Kiln, 24 months, 1st at signing | $2,400.00 | | | $109.00 | 24 | Yes | $2,616.00 | $216.00 | 0.7644% | 9.17% | 9.57% | OK |
| Booth tent | $899.00 | $99.00 | | $72.00 | 12 | No | $963.00 | $64.00 | 1.2043% | 14.45% | 15.45% | OK |
| Laptop | $1,500.00 | $300.00 | | $75.00 | 18 | No | $1,650.00 | $150.00 | 1.2704% | 15.24% | 16.36% | OK |
| Display case | $1,800.00 | | $49.00 | $158.00 | 12 | No | $1,945.00 | $145.00 | 1.2457% | 14.95% | 16.02% | OK |
| Sewing machine, 0% | $1,200.00 | | | $100.00 | 12 | No | $1,200.00 | $0.00 | 0.0000% | 0.00% | 0.00% | No interest |
| Card reader, pay in 4 | $299.00 | | | $74.75 | 4 | Yes | $299.00 | $0.00 | 0.0000% | 0.00% | 0.00% | No interest |
| Vinyl cutter | $600.00 | | | $95.00 | 6 | No | $570.00 | -$30.00 | -1.4461% | -17.35% | -16.04% | Costs less than cash |
Summary P4:P8: 8, Kiln, 24 months, 8.75%, 2, 0. Quotable: the same kiln costs 8.41% x 12 with payments at the end of each month and 9.17% with the first payment at signing; the Display case's $49.00 fee lifts it from 9.70% (RATE ignoring the fee, probe) to 14.95% x 12.

Edge cases tested (all matched a Python bisection solver to 1E-7)
- Payment 0, blank, "$90" text, negative; 11.5 payments; 0 payments; down payment equal to and above the price; down + fee equal to the price; price "$1000" text and 0; down "100" text and -100; fee -10; G "Y" or 1 (Check); G "yes" and "  YES " (read as Yes); G blank (read as No); a name only; G only.
- 1 payment of $1,050.00 on $1,000.00 (5.00% a month); the same at signing (Check); 360 payments of $1,000.00 on $160,000.00 financed (0.5327%, 6.39% x 12); fee with no down (2.0305%) and fee with signing (2.4228%); 12 x $100.01 on $1,200.00 (tiny interest, becomes the lowest-rate offer); 12 x $99.999 (Costs less than cash); 4 x $74.00 at signing on $300.00 (-0.8876%).
- Very high rates: 12 x $60.00 on $100.00 (59.78% a month) and 24 x $100.00 on $100.00 (about 100% a month) were found; 36 x $1,000.00 on $100.00 (1,000% a month) shows Check because RATE did not converge (probe: `=RATE(36,-1000,100)` and with a guess of 5 both returned #N/A).
- 8 random sets of 50 offers (prices $100.00 to $6,000.00, optional down and fee, 3 to 60 payments, rates from -1% to 4% a month and 0%, Yes/No/blank).

Pitfalls the post must state
- The payment is negative in RATE (money going out); `=RATE(24,109,2400)` with both positive returned #N/A in the probe.
- Use the type argument 1 when the first payment is due at signing; it raises the rate (8.41% to 9.17% for the kiln).
- Subtract the down payment and any upfront fee from the price before RATE; leaving the fee out understates the cost (9.70% vs 14.95% for the Display case).
- Rate x 12 is the simple yearly figure; the effective yearly rate counts interest on interest and is higher. Neither is the APR a lender must disclose, and the post gives no advice on whether to finance (gate 3; link the CFPB APR page only if checked).
- 0% plans: RATE can return a tiny number instead of 0 (1.7E-10 in the probe); the file returns 0 when the payments add up to the price.
- Not tested: Excel, Google Sheets (RATE convergence and its default guess, the type argument, SMALL with COUNTIF are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- #N/A from RATE: the payment was entered as a positive number (`=RATE(24,109,2400)`), or the rate is extreme (`=RATE(36,-1000,100)`, also with guess 5).
- The rate looks too low: the upfront fee was left out (Display case 0.8086% vs 1.2457% a month), or the first payment is at signing but type was left at 0 (kiln 0.7012% vs 0.7644%).
- A 0% plan shows 0.0000% but is not exactly 0 (1.7E-10); harmless, the file shows No interest.
- A negative rate: the payments add up to less than the cash price (Vinyl cutter, 6 x $95.00 = $570.00 on $600.00).
- A one-payment plan due at signing returned -99.99999% from RATE (probe); the file shows Check.

---
What was NOT run (all five): Microsoft Excel, Google Sheets, any non-US language setting, more rows than the slots (54, 203, 204, 104), a printer (PDF export only). Nothing is claimed about Excel or Sheets.

---
Images 2026-11-14 (designer)

Heroes: `heroes_2026_11_14` in scripts/make-heroes.py, run with `python3 -B scripts/make-heroes.py --only-2026-11-14` (no earlier entry run or changed). Each hero line sits directly after the `> Works in` callout in the draft posts (blank line before and after).

Hero files (1600x900 PNG, checked with PIL, real LibreOffice PDF render of the template with its sample data, no error values or ### in any image). Hero alt texts below are 8 to 15 words and differ from the titles:
1. images/2026-11-14-team-01-find-best-and-worst-month-per-product-column-header-of-max-index-match-excel-google-sheets/best-and-worst-month-per-product-template.png
   - caption: "Eight products over ten months; Wax melts ties for best in Jun and Jul, and Discontinued soap shows No sales. PDF export from LibreOffice Calc 24.2."
   - alt: "Monthly product sales grid with best month, worst month and tie count for each row"
2. images/2026-11-14-team-02-convert-date-to-julian-date-code-yyddd-batch-lot-number-excel-google-sheets/julian-date-lot-codes-template.png
   - caption: "Left, ten batches get a YYDDD code and lot code (a date typed as text shows Check); right, nine printed codes read back into dates, three of them Check. PDF export from LibreOffice Calc 24.2."
   - alt: "Batch dates turned into Julian lot codes, and printed codes read back into dates"
3. images/2026-11-14-team-03-match-payouts-to-bank-deposits-repeat-amounts-countif-match-excel-google-sheets/match-payouts-to-bank-deposits-template.png
   - caption: "Ten payouts with 3 days allowed: seven Matched, one Check date at 4 days and two Not in bank, with the summary at right. PDF export from LibreOffice Calc 24.2."
   - alt: "Three $412.50 payouts numbered 1 to 3 and each paired with its own deposit"
4. images/2026-11-14-team-04-split-booking-days-and-amount-across-months-eomonth-min-max-excel-google-sheets/split-amount-across-months-template.png
   - caption: "Six jobs split by days into month columns, shown through Mar 2027 (later months are hidden here, not in the file); an end date before the start shows Check. PDF export from LibreOffice Calc 24.2."
   - alt: "Booking amounts spread over October to March columns by days, with an Adds up check"
5. images/2026-11-14-team-05-find-interest-rate-of-payment-plan-offer-rate-function-excel-google-sheets/payment-plan-interest-rate-template.png
   - caption: "Eight payment plan offers side by side (down payment and fee columns are hidden here); two 0% plans show No interest and the vinyl cutter costs less than cash. PDF export from LibreOffice Calc 24.2."
   - alt: "Eight financing offers with cash price, monthly payment, rate times 12 and effective yearly rate"
Shown in each hero: 1 Products rows 4-12, Nov, Dec, Total, Months filled in and the summary (L, M, N, T, W:X) hidden; 2 Make codes without Batch of the day (C) and the summary, beside Read codes without the Cleaned helper (B) and the summary; 3 Payouts without the helpers D, F, G, summary L4:M10 kept; 4 Amounts with Days (E) and the months after Mar 2027 (M:R) hidden, so the retainer's visible months do not add up to In the grid (the caption says later months are hidden); 5 Offers without Down payment, Upfront fee, Total paid, Monthly rate (C, D, H, J) and the summary, columns A, F, G, M widened in the hero copy only. All hiding is in the hero copy, not in the template files.

AI photo (`image_prompts`, `image_alt`), one scene each, 2 or 3 specific objects, no screens, no readable text, no logos, no laptop and coffee:
1. Best month: image_prompts "A small gift shop's seasonal shelf with glass jars of wax melts, a wicker basket of glass holiday ornaments and a stack of folded linen tea towels, price tags blank and unreadable, soft window light, realistic photo, no screens, no readable text, no logos" image_alt "Wax melts, holiday ornaments and stacked linen tea towels on a shop shelf"
2. Julian: image_prompts "Freshly filled jam jars cooling upside down on a wire rack in a home kitchen, a small handheld roller coder for marking lot codes resting beside them and an open carton of spare jar lids, any marks too small to read, soft window light, realistic photo, no screens, no readable text, no logos" image_alt "Filled jam jars cooling upside down on a wire rack beside a handheld coder"
3. Payouts: image_prompts "Three matching rubber-banded bundles of paper receipts lined up beside an open bank deposit slip book and a pen on a small shop counter, all print blurred and unreadable, soft daylight, realistic photo, no screens, no readable text, no logos" image_alt "Three matching receipt bundles beside an open deposit slip book on a counter"
4. Split: image_prompts "An empty rented stall at an indoor makers market between seasons, bare pegboard walls, a hand truck loaded with folded display shelves and a padlock with key hanging on the stall's wire gate, soft overhead light, realistic photo, no screens, no readable text, no logos" image_alt "Empty indoor market stall with bare pegboard, a hand truck and a padlocked gate"
5. RATE: image_prompts "A new glass display case still on its shipping pallet in a small gift shop, bubble wrap pulled back from one corner, a clipboard with a blank payment plan form and a pen resting on top, soft daylight, realistic photo, no screens, no readable text, no logos" image_alt "New glass display case on a shipping pallet with a clipboard and pen on top"

All ten alt texts (5 hero, 5 photo) were checked against every `![alt]` and `image_alt` in posts/ (ready and published) and the other new drafts: no duplicates. Photo scenes avoid objects already used in earlier prompts (date stamp and ink pad, canopy tent bag, kiln, zippered deposit bag).
