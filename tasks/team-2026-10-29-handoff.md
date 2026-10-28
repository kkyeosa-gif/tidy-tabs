# Handoff 2026-10-29: five templates built and checked (LibreOffice Calc 24.2.7.2 only)

Status: topics (tasks/team-topics-2026-10-29.md) done; the posts are the next step. The templates were built by five new functions in `scripts/build-templates.py` (`hide_div_zero_errors`, `totals_off_by_a_cent`, `giveaway_random_winners`, `average_ignore_zeros`, `customer_tiers`, added after `batches_and_rotation` and registered at the end of the main block). Only these five were run (a small runner that imports the script and calls just them, with `PYTHONDONTWRITEBYTECODE=1`, using `/usr/bin/python3` because the default `python3` on this machine has no openpyxl); older .xlsx were not rewritten. Nothing was committed. New in `git status`: the five .xlsx, the edited script, tasks/briefs.md, this file (tasks/topics-seed.md, tasks/team-topics-2026-10-29.md and tasks/hooks-2026-10-29.md were there before this task).

Exact version string used for `tested_in`: **LibreOffice 24.2.7.2 420(Build:2)** (Calc, Linux, headless). Write "LibreOffice Calc 24.2.7.2". Nothing was opened in Excel or Google Sheets and nothing may be claimed about them (including whether IFS, TEXT, RANK or error names behave the same there). Method: `soffice --headless --convert-to xlsx` into a scratch dir, read back with openpyxl `data_only=True`, compared with independent Python models that read only the typed inputs (plain division and error rules; `decimal.Decimal` with ROUND_HALF_UP; property checks for the random draw; `statistics.mean`; a threshold function). 277 workbooks in all (5 committed samples + 272 edited copies, of which 200 are recalculations of the giveaway sample), 2,764 cell comparisons in the four deterministic templates plus the property checks on 215 giveaway workbooks. Mismatches: 0 except 4 cells, all explained below (text "0" in Visits; an error value inside the average range, where the model did not expect an error). No error value (#NAME?, #VALUE!, #REF!, #N/A, #DIV/0!, #NUM!, Err:) in any workbook except the cells listed under "Errors shown on purpose" per template. Formulas use only: IF, IFERROR, IFS (written as `_xlfn.IFS`), AND, OR, NOT, ISNUMBER, ISERROR, SUMPRODUCT, SUM, SUMIF, COUNT, COUNTA, COUNTIF, AVERAGE, AVERAGEIF, ROUND, ABS, MIN, MAX, RAND, RANK, INDEX, MATCH, ROWS, LEFT, NA (no AGGREGATE, TODAY, NOW, FILTER, LET, LAMBDA, XLOOKUP, SORT, RANDARRAY, no array entry). Sample data fictional, dollars US. Yellow (FFF2CC) = inputs, gray (EDEDED) = formulas. Each file has a "How to use" sheet and one CHECK cell. PDF export: every page US Letter (landscape main tabs, portrait How to use / giveaway Entries).

Titles and slugs are in tasks/team-topics-2026-10-29.md. Posts must say: LibreOffice Calc 24.2.7.2 only, sample data fictional, no advice (no contest or sweepstakes rules, no tax or invoice rounding requirement, no discount or customer-treatment advice).

IFS result (topic 5 gate): typed as `=_xlfn.IFS(...)` in the .xlsx XML, it calculates in LibreOffice (1200 -> Gold, 500 -> Silver, 499.99 -> Bronze). Written as plain `=IFS(...)` in the file XML it gives #NAME? in LibreOffice (the prefix is required in the file; a user typing IFS( in the app is a different path and was not run). Without the TRUE catch-all pair, a value below Silver gives #N/A (confirmed). Text such as "abc" tests as `>=1000` TRUE and gets Gold, so the template guards with ISNUMBER first. The nested IF alternative is built into the template (column F) and matched IFS in all cases. The post may quote either; say which was tested.

---

## 1. Hide #DIV/0! (`templates/tidy-tabs-hide-div-zero-errors.xlsx`, function `hide_div_zero_errors`)
- Sheets: Examples (landscape, print area A1:J16, header row 4 repeats on print), How to use.
- Inputs (yellow): B5:B16 Orders, C5:C16 Visits (12 fictional weeks, A = "Week 1" ... "Week 12", not an input). No dropdowns. Rows 5 to 16 only (insert rows inside the table to add more).
- Columns: D raw division, E IFERROR, F IF visits is 0, G IF both numbers and visits not 0. Headers show the formula text. Errors in D:F show red text.
- Summary (labels I, values J): J5 errors in raw column, J6 rates shown in G, J7 average of G, J8 average of raw D (error on purpose), J9 CHECK. Typo demo I11:J14.
- Formulas to quote (row 5):
  - D5 raw: `=B5/C5`
  - E5 IFERROR: `=IFERROR(B5/C5,"")`
  - F5 targeted IF: `=IF(C5=0,"",B5/C5)`
  - G5 number and not zero: `=IF(AND(ISNUMBER(B5),ISNUMBER(C5)),IF(C5<>0,B5/C5,""),"")` (nested IF; the topics file's `=IF(AND(ISNUMBER(C5),C5<>0),B5/C5,"")` was not run as written, G adds ISNUMBER(B5) so a text or blank Orders cell gives a blank, not a 0 or #VALUE!)
  - J5 `=SUMPRODUCT(--ISERROR(D5:D16))`; J7 `=IF(COUNT(G5:G16)=0,"No rates",AVERAGE(G5:G16))`; J8 `=AVERAGE(D5:D16)`
  - Typo demo: J12 `=B5/Visits` (name does not exist), J13 `=IFERROR(B5/Visits,"")`, J14 `=IF(C5=0,"",B5/Visits)`
  - J9 CHECK: `=IF(AND(SUMPRODUCT(--ISERROR(E5:E16))=0,SUMPRODUCT(--ISERROR(G5:G16))=0,COUNT(G5:G16)=SUMPRODUCT(--ISNUMBER(B5:B16),--ISNUMBER(C5:C16),--(C5:C16<>0))),"OK","CHECK")`
- Errors shown on purpose (sample): D6, D8, D9, D15 `#DIV/0!` (Visits 0, 0/0, blank, 0), D10 and F10 `#VALUE!` (Visits is the text "n/a"), J8 `#DIV/0!`, J12 and J14 `#NAME?`. J13 is blank (IFERROR hid the typo).
- Example table values (recalculated, equal to the model):
  | Week | Orders | Visits | D raw | E IFERROR | F IF zero | G IF number |
  | 1 | 14 | 410 | 3.4% | 3.4% | 3.4% | 3.4% |
  | 2 | 9 | 0 | #DIV/0! | blank | blank | blank |
  | 3 | 22 | 655 | 3.4% | 3.4% | 3.4% | 3.4% |
  | 4 | 0 | 0 | #DIV/0! | blank | blank | blank |
  | 5 | 18 | (blank) | #DIV/0! | blank | blank | blank |
  | 6 | 11 | n/a | #VALUE! | blank | #VALUE! | blank |
  | 7 | 30 | 812 | 3.7% | 3.7% | 3.7% | 3.7% |
  | 9 | 0 | 380 | 0.0% | 0.0% | 0.0% | 0.0% |
  | 11 | 25 | 0 | #DIV/0! | blank | blank | blank |
  Summary: J5 5, J6 7, J7 2.8% (0.028454), J8 #DIV/0!, J9 OK.
- Edge findings (14 workbooks, 728 cell comparisons, 2 explained mismatches):
  - All Visits 0 or all blank: J5 12, J6 0, J7 "No rates", J9 OK. All Visits text "n/a": raw column is #VALUE!, J8 is #VALUE! (AVERAGE returns the first error it meets, not always #DIV/0!). 0/0 in all rows: same as zero.
  - Negative Visits: rate is negative (-5 / -10 and so on), shown and averaged, no error. Decimal and huge numbers (1e9 / 3e12) fine.
  - Orders blank with Visits 50: raw D, E and F give 0.0% (a blank counts as 0 in the division); G stays blank (ISNUMBER(B5) is false). Orders text "x": D #VALUE!, E blank, F #VALUE!, G blank.
  - Visits typed as the text "0" (quotes, stored as text): LibreOffice treats the text as 0 in the division, so D5 and F5 give #DIV/0! (not #VALUE!); ISNUMBER is false so G is blank and E is blank. This is the 2 explained mismatches. Say numbers stored as text can behave unexpectedly and G is the safe column.
  - The "" results are text: `=G9*2` on a "" cell gives #VALUE!, `=G9&"x"` gives "x", SUM(G)+SUM(E) works (text skipped). IFERROR never changed a good result.
  - Typo: the plain formula gives #NAME?, IFERROR returns "", the IF version still gives #NAME?. This is the central example for the post: a blank can mean a hidden mistake.
- Limits: LibreOffice names and wording of the error values only. Excel and Sheets error names/wording are not tested (the topics file asked for a link to the Microsoft IFERROR help page). Not run: #REF! from a deleted column (the typo demo uses #NAME?), percent format on the result in another app.

---

## 2. Totals off by a cent (`templates/tidy-tabs-totals-off-by-a-cent.xlsx`, function `totals_off_by_a_cent`)
- Sheets: Invoice (landscape, print area A1:J24, header row 4 repeats), How to use.
- Inputs (yellow): A5:A24 item, B5:B24 Qty, C5:C24 Unit price (formatted to 4 decimals). 20 slots, 8 filled with fictional supplies.
- Columns: D line unrounded (formatted to 2 decimals), E line rounded, F rounded minus unrounded (4 decimals), G line in whole cents.
- Summary (labels I, values J): J5 total of unrounded lines, J6 total of rounded lines, J7 difference of the two totals shown to cents, J8 lines counted, J9 CHECK.
- Formulas to quote (row 5):
  - D5 `=IF(AND(ISNUMBER(B5),ISNUMBER(C5)),B5*C5,"")`; core `=B5*C5`
  - E5 `=IF(D5="","",ROUND(D5,2))`; core `=ROUND(B5*C5,2)`
  - F5 `=IF(D5="","",ROUND(E5-D5,4))`; G5 `=IF(E5="","",ROUND(E5*100,0))`
  - J5 `=SUM(D5:D24)`; J6 `=SUM(E5:E24)`; J7 `=ROUND(ROUND(J5,2)-J6,2)` (the topics file had `ROUND(total_unrounded,2)-total_rounded`; J7 wraps the outer ROUND so a 2-cent float tail cannot appear); J8 `=COUNT(D5:D24)`
  - J9 CHECK: `=IF(AND(COUNT(B5:B24)=COUNTA(B5:B24),COUNT(C5:C24)=COUNTA(C5:C24),COUNT(B5:B24)=COUNT(C5:C24),ROUND(SUM(G5:G24)-J6*100,0)=0,ABS(J7)<=0.005*J8+0.005),"OK","CHECK")`
- Example table values (sample; Decimal half-up model agrees on every cell):
  | Item | Qty | Unit | Unrounded | Rounded |
  | Soy wax flakes (oz) | 3 | $0.3350 | 1.005 | $1.01 |
  | Cotton wick | 7 | $1.1450 | 8.015 | $8.02 |
  | Paper label | 12 | $0.0425 | 0.51 | $0.51 |
  | Tin lid | 5 | $2.3490 | 11.745 | $11.75 |
  | Fragrance oil (ml) | 9 | $0.8750 | 7.875 | $7.88 |
  | Gift tag | 24 | $0.1675 | 4.02 | $4.02 |
  | Kraft sleeve | 6 | $3.4575 | 20.745 | $20.75 |
  | Twine (ft) | 15 | $0.0850 | 1.275 | $1.28 |
  Summary: J5 $55.19, J6 $55.22, J7 -$0.03 (the displayed lines add to 55.22 but a plain SUM of the unrounded column shows 55.19), J8 8, J9 OK. The sample's half-cent lines (1.005, 8.015, 11.745, 7.875, 20.745, 1.275) all rounded up in LibreOffice and match Decimal ROUND_HALF_UP.
- Edge findings (17 workbooks, 848 cell comparisons, 0 mismatches; the model uses Decimal(Qty) times Decimal(repr(price)) with ROUND_HALF_UP):
  - All whole cents (3 x 1.25, 10 x 0.99, 7 x 4.5, 2 x 12.34): J7 $0.00, totals both $69.83.
  - Half-cent prices with odd quantities (8 lines): J5 $3.42, J6 $3.46, J7 -$0.04. 20 lines of 1 x $0.125: J5 $2.50, J6 $2.60, J7 -$0.10. 20 lines of 1 x $0.005: J5 $0.10, J6 $0.20 (every line rounds up), J7 -$0.10.
  - Large quantity (1,000,000 x 0.335; 250,000 x 1.145; 999,999 x 0.0425): J5 $663,749.9575, J6 $663,749.96, J7 $0.00 (J7 is shown to cents, so a 0.0025 gap vanishes).
  - Negative quantities (returns) round half away from zero: -3 x 0.335 = -1.005 -> -$1.01. Zero quantity line: $0.00 and difference $0.0000. One line (3 x 0.335): J5 $1.005, J6 $1.01, J7 $0.00 (1.005 shown to cents is $1.01 too, so both totals look equal).
  - Random 3 and 4 decimal prices, 20 lines each (5 workbooks): J7 between -$0.03 and +$0.02, always within the CHECK bound.
  - Blank row in the middle: skipped. Price typed as text ("1.145") or Qty without a price: that row has no line and J9 shows CHECK.
- Limits: says nothing about which rounding an invoice, customer or authority requires. Excel's "Set precision as displayed" is mentioned in the How to use sheet as a risky option (not tested here). Excel and Sheets half-cent rounding not tested. Not run: more than 20 lines, prices above $1,000,000.

---

## 3. Giveaway random winners (`templates/tidy-tabs-giveaway-random-winners.xlsx`, function `giveaway_random_winners`)
- Sheets: Entries (portrait, print area A1:E34), Winners (landscape, print area A1:F14), How to use.
- Inputs (yellow): Entries A5:B34 name and city (30 fictional entries, one per row, blank rows allowed); Winners B1 number of winners (3; whole number 1 to 10, with data validation).
- Entries columns: C random number, D draw order, E "Winner" when D is within the number of winners (green row). Headers in row 4.
- Winners: B2 input check (OK or a message); table header row 4, rows 5 to 14 (No., Winner, City); summary E1:F5.
- Formulas to quote:
  - Entries C5 `=IF(A5="","",RAND())`
  - Entries D5 `=IF(A5="","",RANK(C5,$C$5:$C$34)+COUNTIF($C$5:C5,C5)-1)`
  - Entries E5 `=IF(D5="","",IF(D5<=Winners!$F$2,"Winner",""))`
  - Winners F1 entries counted `=COUNTA(Entries!$A$5:$A$34)`; F2 winners to draw `=IF(ISNUMBER(B1),IF(AND(B1>=1,B1<=10,B1=INT(B1)),MIN(B1,F1),0),0)`
  - Winners A5 `=IF(ROWS(A$5:A5)<=$F$2,ROWS(A$5:A5),"")`
  - Winners B5 `=IF(A5="","",INDEX(Entries!$A$5:$A$34,MATCH(A5,Entries!$D$5:$D$34,0))&"")` (the topics file's version `=IF(ROWS($A$2:A2)>$B$1,"",INDEX(...,MATCH(ROWS($A$2:A2),Entries!$D$5:$D$34,0)))` is the same idea; it would give #N/A when winners > entries, so the template draws MIN(B1, entries))
  - Winners C5 same with Entries!$B$5:$B$34
  - Winners B2 `=IF(ISNUMBER(B1),IF(AND(B1>=1,B1<=10,B1=INT(B1)),IF(F1=0,"Add names on the Entries tab",IF(B1>F1,"Fewer entries ("&F1&") than winners: all of them win","OK")),"Winners must be a whole number from 1 to 10"),"Winners must be a number")`
  - F5 CHECK: `=IF(AND(F1>=1,F2>=1,F3=F2,F4=F2,MAX(Entries!$D$5:$D$34)=F1,COUNT(Entries!$D$5:$D$34)=F1,SUMPRODUCT((Entries!$D$5:$D$34<>"")*(COUNTIF(Entries!$D$5:$D$34,Entries!$D$5:$D$34)>1))=0,SUMPRODUCT(--(B5:B14<>""),--(COUNTIF(Entries!$A$5:$A$34,B5:B14)>0))=F3,COUNTA(Entries!A35:A500)=0),"OK","CHECK")` (F3 winners listed `=SUMPRODUCT(--(B5:B14<>""))`, F4 rows marked Winner `=COUNTIF(Entries!$E$5:$E$34,"Winner")`)
- Stable values: RAND is volatile, so the sample has NO fixed winners. The committed .xlsx stores no results; LibreOffice draws a new set every time the file opens or any cell changes (200 opens gave 196 different 3-name sets). The post should not show a fixed answer as "the" answer; for a screenshot or example use one recalculated copy and say the names are one random draw, or freeze it as the How to use sheet says (copy Entries C5:D34, Paste special, Values only). Freezing was not run through an app (the checks froze nothing); the claim in the file is the standard copy / paste-values step and the post should say it was not tested in the template.
- Example values: structure only. Each recalculation gives Entries D = a permutation of 1..30, three rows marked Winner (green), Winners B5:C7 = those three names and cities in draw order, F1 30, F2 3, F3 3, F4 3, F5 OK, B2 OK. One recalculation produced Leo (Dayton), Hugo (Provo), Rosa (Laredo); do not quote it as fixed.
- Edge findings (215 workbooks: 200 sample recalculations + 15 edited variants):
  - 200 recalculations of the sample: draw order always exactly 1..30 with no repeats, always 3 winners, all different rows, all names in the list, Winner rows = the 3 largest random numbers, in order; F5 OK 200 of 200; no error values. Every one of the 30 names won at least once (wins per name between 9 and 32 of 600 draws; expected 20; rough spread only, not a statistical test). 196 distinct winner sets in 200 runs.
  - Winners 1 and 10 with 30 entries: correct counts (1 and 10 different rows), F5 OK.
  - Fewer entries than winners: 5 entries and 10 winners -> all 5 listed, B2 "Fewer entries (5) than winners: all of them win", F5 OK. 1 entry and 3 winners -> 1 winner. 10 entries and 10 winners -> all 10, B2 OK.
  - 0 entries: no winners, B2 "Add names on the Entries tab", F5 CHECK (no error values).
  - Duplicate names ("Sam","Sam","Sam","Ana","Ana","Ben"): treated as separate rows; the draw returned 'Sam', 'Ana', 'Sam' (two different Sam rows), F5 OK. 30 entries all named "Pat" with 10 winners: 10 Pat rows, F5 OK (the CHECK works on rows, not names).
  - Blank rows in the middle (names in rows 5-8, 12-16, 18-20): skipped, order is 1..12, 3 winners. Numeric names (101..106): winners listed as text ('106', '103', '102').
  - Winners number 0, 11, 2.5 or text "three": B2 message, no winners listed, F5 CHECK. A name typed in A40 (below the list): not drawn, F5 CHECK.
- Limits: RAND is pseudo-random with no audit trail; no advice on running a giveaway, rules, prizes or announcing results. Names containing * or ? can confuse COUNTIF in the CHECK (not run). Not run: more than 30 entries, RANDBETWEEN, SORTBY or RANDARRAY. Behavior in Excel and Sheets (RAND recalculation rules, RANK) not tested.

---

## 4. Average that ignores zeros, blanks and text (`templates/tidy-tabs-average-ignore-zeros.xlsx`, function `average_ignore_zeros`)
- Sheets: Daily sales (landscape, print area A1:G32, header row 4 repeats), How to use.
- Inputs (yellow): B5:B32 daily sales (28 fictional market days, A = day number 1 to 28 is not an input). Sample: 6 days at 0, 1 blank (day 18), 1 text "rain" (day 14), 20 days with sales. Error demo block G5:G9 is also yellow (40, #N/A via `=NA()`, 0, 55.5, 30), kept apart from the main range.
- Results (labels D, values E): E5 AVERAGE, E6 AVERAGEIF, E7 SUMIF / COUNTIF, E8 guarded, E9 days counted, E10 days at 0, E11 negative values, E12 text entries, E13 blank cells, E14 CHECK, E17 and E18 error demo.
- Formulas to quote:
  - E5 `=AVERAGE(B5:B32)`
  - E6 `=AVERAGEIF(B5:B32,">0")`
  - E7 `=IF(E9=0,"No days",SUMIF(B5:B32,">0")/E9)` (the topics file's `SUM(range)/COUNTIF(range,">0")` would also add negatives; SUMIF keeps the two the same)
  - E8 `=IF(COUNTIF(B5:B32,">0")=0,"No days",AVERAGEIF(B5:B32,">0"))`
  - E9 `=COUNTIF(B5:B32,">0")`; E10 `=COUNTIF(B5:B32,0)` (the topics file's `COUNT(range)-COUNTIF(range,">0")` would lump negatives in with zeros; this template counts them separately in E11 `=COUNTIF(B5:B32,"<0")`)
  - E12 `=COUNTA(B5:B32)-COUNT(B5:B32)`; E13 `=ROWS(B5:B32)-COUNTA(B5:B32)`
  - E14 CHECK: `=IF(AND(SUMPRODUCT(--ISERROR(B5:B32))=0,E9+E10+E11=COUNT(B5:B32),E12+E13+COUNT(B5:B32)=ROWS(B5:B32),IF(E9=0,E8="No days",AND(ROUND(E7-E6,6)=0,ROUND(E8-E6,6)=0))),"OK","CHECK")`
  - E17 `=AVERAGE(G5:G9)`; E18 `=AVERAGEIF(G5:G9,">0")`
- Errors shown on purpose (sample): G6 `#N/A`, E17 `#N/A`. In other edited copies E5 and E6 may show #DIV/0! (see below).
- Example values (sample): E5 $157.29 (157.2885), E6 $204.48 (204.475), E7 $204.48, E8 $204.48, E9 20, E10 6, E11 0, E12 1, E13 1, E14 OK, E17 #N/A, E18 $41.83. The "too low" gap is $47.19. Days counted 20 + zero days 6 + text 1 + blank 1 = 28.
- Edge findings (17 workbooks, 119 cell comparisons; the only differences from the model are the two error-in-range copies, where the model did not expect an error):
  - All zero: E5 $0.00, E6 `#DIV/0!` (AVERAGEIF with nothing above 0), E7 and E8 "No days", E9 0, E10 28, E14 OK. This is the case that needs the guard. All blank and all text: E5 and E6 `#DIV/0!`, E7 and E8 "No days", E12 28 (text) or E13 28 (blank), E14 OK.
  - One value ($50.50): all four averages $50.50, E9 1.
  - Negatives (10, -5, 20, -30, 0): E5 -1, E6 15, E9 2, E10 1, E11 2, E14 OK. Negatives are skipped by `">0"`.
  - Text in a number cell: counted in E12, skipped by every average. Numbers stored as text ("45", "50"): skipped by AVERAGE, AVERAGEIF and COUNTIF `">0"` and counted as text in E12 (2), so they are left out of the average silently (the post should say so and point to the convert-text template if one exists). The text "0": not counted by COUNTIF(range,0) (E10), counted in E12.
  - Real $0.00 days as open days (0, 0, 100, 200): E5 $75, E6 $150; the reader chooses.
  - Tiny and huge (0.01, 1e9, 0, -1e-9): E6 and E7 agree at 500000000.005. Booleans TRUE/FALSE are treated as numbers by LibreOffice (TRUE = 1): not recommended, not in the sample.
  - An #N/A or a #DIV/0! typed into the main range: E5 shows that error, E12 counts it as non-number, E14 CHECK. AVERAGEIF(">0") skipped the error cell and still returned $204.475 in E6, and E7/E8 too, so only E5 and E14 reveal it. Error demo: AVERAGE returns #N/A, AVERAGEIF(">0") returns $41.83 (40, 55.5, 30) in LibreOffice. Say this is LibreOffice's result only.
  - Two random mixed lists (zeros, blanks, "n", negatives, positives): matched `statistics.mean`.
- Limits: a real $0.00 day may belong in an average; the template does not decide. Excel and Sheets handling of errors and numbers-as-text inside AVERAGEIF was not tested. Not run: more than 28 rows, dates in the range.

---

## 5. Customer tiers (`templates/tidy-tabs-customer-tiers.xlsx`, function `customer_tiers`)
- Sheets: Customers (landscape, print area A1:J40, header row 4 repeats), Cutoffs, How to use.
- Inputs (yellow): Cutoffs B2 Silver from $500.00, B3 Gold from $1,000.00; Customers A5:C104 (customer, city, lifetime spend; 25 fictional customers, 100 slots). Cutoffs B4 = cutoff check (OK or a message).
- Customers columns: D label (IFS), E dollars to next label, F label by nested IF (no IFS). Colors: Gold, Silver, Bronze fills; "Check ..." red.
- Summary (labels H, values I and J): I5:J7 count and spend for Gold, Silver, Bronze (`=COUNTIF` and `=SUMIF`), I8 rows that need a look, I9 customers with a spend entered, I10 CHECK.
- Formulas to quote (row 5; the file stores IFS as `_xlfn.IFS`, shown here as the user types it):
  - D5 label: `=IF(C5="","",IF(Cutoffs!$B$4<>"OK","Check cutoffs",IF(OR(NOT(ISNUMBER(C5)),C5<0),"Check spend",IFS(C5>=Cutoffs!$B$3,"Gold",C5>=Cutoffs!$B$2,"Silver",TRUE,"Bronze"))))`; core `=IFS(C5>=Cutoffs!$B$3,"Gold",C5>=Cutoffs!$B$2,"Silver",TRUE,"Bronze")`
  - F5 nested IF: `=IF(C5="","",IF(Cutoffs!$B$4<>"OK","Check cutoffs",IF(OR(NOT(ISNUMBER(C5)),C5<0),"Check spend",IF(C5>=Cutoffs!$B$3,"Gold",IF(C5>=Cutoffs!$B$2,"Silver","Bronze")))))`; core `=IF(C5>=Cutoffs!$B$3,"Gold",IF(C5>=Cutoffs!$B$2,"Silver","Bronze"))`
  - E5 to next: `=IF(OR(D5="Silver",D5="Bronze"),ROUND(IF(D5="Silver",Cutoffs!$B$3,Cutoffs!$B$2)-C5,2),"")` (Gold and check rows blank; the ROUND removes float tails such as 0.01000000000000227)
  - Cutoffs B4: `=IF(AND(ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B2>0,B2<B3),"OK","Silver must be above 0 and below Gold"),"Type a number in each yellow cell")`
  - I5 `=COUNTIF(D5:D104,H5)`; J5 `=SUMIF(D5:D104,H5,C5:C104)`; I8 `=COUNTIF(D5:D104,"Check*")`; I9 `=COUNTA(C5:C104)`
  - I10 CHECK: `=IF(AND(Cutoffs!B4="OK",I5+I6+I7+I8=I9,SUMPRODUCT(--(D5:D104<>F5:F104))=0,COUNTA(C105:C500)=0),"OK","CHECK")`
- Example table values (sample; every label, to-next and summary cell matches the model):
  | Customer | Spend | Label | To next label |
  | Alma R. | $1,240.50 | Gold | blank |
  | Ben T. | $85.00 | Bronze | $415.00 |
  | Cora L. | $500.00 | Silver | $500.00 |
  | Dev K. | $499.99 | Bronze | $0.01 |
  | Elena M. | $1,000.00 | Gold | blank |
  | Hugo P. | $0.00 | Bronze | $500.00 |
  | Jonah B. | $975.40 | Silver | $24.60 |
  | Opal F. | $999.99 | Silver | $0.01 |
  | Wren I. | $499.00 | Bronze | $1.00 |
  Summary: Gold 6 customers $8,945.75; Silver 8 customers $5,811.74; Bronze 11 customers $2,367.99; I8 0; I9 25; I10 OK.
- Edge findings (14 workbooks, 1,069 cell comparisons, 0 mismatches; IFS and nested IF columns agreed everywhere):
  - Exactly on a cutoff gets the higher label: $500.00 Silver, $1,000.00 Gold. $499.99 Bronze (to next $0.01), $999.99 Silver (to next $0.01), $1,000.01 Gold, 0 and 0.01 Bronze. Cutoffs 100/250 with 99.99, 100, 100.01, 249.99, 250, 250.01: Bronze, Silver, Silver, Silver, Gold, Gold.
  - Cutoffs 200/400 on the sample: 17 Gold, 2 Silver, 6 Bronze. Large values (1e7, 1e9) are Gold; 499.995 Bronze, 500.005 Silver.
  - Blank spend: blank label. Negative spend (-5): "Check spend". Text spend "abc" or the text "750": "Check spend" (not a label; without the ISNUMBER guard "abc" would get Gold). I8 counts these and I10 stays OK when the only issue is a Check row (the row shows in I8; say so).
  - Cutoffs in the wrong order (1000/500), equal (500/500), Silver 0, negative Silver, or text in a cutoff: Cutoffs B4 shows the message, every label shows "Check cutoffs", I10 CHECK, no counts for Gold/Silver/Bronze.
  - 100 customers (two random sets, including exact 500, 1000, 499.99 values): all labels matched; per-label counts and spends matched the model to 1e-6. All blank: counts 0, I10 OK.
  - Without the TRUE catch-all pair (a test formula, not in the template): a spend below Silver gives #N/A. With a lone `IFS(...)` written without the `_xlfn.` prefix in the file XML: #NAME? in LibreOffice.
- Limits: labels only; no advice about discounts or what a customer is owed. Excel 2019 / 365 and current Sheets have IFS and Excel 2016 does not, per the topics file; this was not tested here, so the post says LibreOffice Calc 24.2.7.2 only and offers the nested IF column F as the no-IFS version. Not run: more than 100 customers, a third cutoff (Platinum).
