# Handoff for posts dated 2026-10-17

Status: topics (researcher, tasks/team-topics-2026-10-17.md) done; titles and openings (hook-writer) and the template builds (template-builder) are the next steps for the posts. Templates were built by the new functions in `scripts/build-templates.py` (`overlapping_bookings`, `cash_drawer_count`, `nth_weekday_of_month`, `format_phone_numbers`, `scale_recipe_batch`, registered at the end of the main block) and recalculated in LibreOffice Calc (Linux, headless: `soffice --headless --convert-to xlsx` into a scratch dir, read back with openpyxl `data_only`). Every result below was compared with an independent Python calculation (pairwise date test, Decimal arithmetic, `calendar`, `re`). Edge cases were run by editing a copy of the template with openpyxl, recalculating it in LibreOffice and reading the values. No error values (#NAME?, Err:, #VALUE!) appeared in the committed templates after recalculation; one #VALUE! was found in an early draft of the cash drawer formula (count typed as text) and fixed, see post 2. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. The functions used (COUNTIFS, SUMPRODUCT, COUNTIF, COUNTA, IF, AND, OR, NOT, ISNUMBER, INT, ROUND, ROUNDUP, SUM, MATCH, IFERROR, DATE, EOMONTH, EDATE, WEEKDAY, MOD, MONTH, SUBSTITUTE, LEFT, MID, RIGHT, LEN, FIND) exist in both Excel and Google Sheets by name, but that was not run there. No tasks/search-stats.md exists: no search data. Sample data is fictional (say so once per post). PDF export from LibreOffice gave US Letter pages for every tab (Bookings Letter landscape 792 x 612 pt, all other tabs portrait 612 x 792 pt); each tab prints on one page (2 pages for the two-tab files, 3 for nth weekday).

Only the five new functions were run for the committed files (a small runner that imports the script and calls just those five). Older .xlsx were not rewritten; git status shows only the five new .xlsx plus the edited script and task notes (tasks/topics-seed.md and tasks/hooks-2026-10-17.md changes belong to other roles).

Titles carry "Excel and Google Sheets" and have no colon, dash or hyphen. Cell addresses below are the real ones in the built files.

Deviations from the topic plan are marked "Plan change" in each post.

---

## 1. 2026-10-17-team-01-find-overlapping-bookings-date-ranges-countifs-excel-google-sheets
- Title: (hook-writer) example: How to Find Overlapping Dates and Double Bookings in Excel and Google Sheets
- Template: `templates/tidy-tabs-overlapping-bookings.xlsx` (function `overlapping_bookings`). Sheets: Bookings (landscape), How to use (portrait).
- Inputs (yellow): Bookings A2:E40 (Booking ID, Space with a dropdown Booth A/Booth B/Studio, Client, Start date, End date; 12 sample rows BK-101 to BK-112, October 2026); K2 same-day handoff counts as a conflict = Yes (dropdown Yes/No).
- Formulas (filled to row 40; $B$2:$B$40 space, $D$2:$D$40 start, $E$2:$E$40 end):
  - F2 `=IF(A2="","",IF(OR(B2="",NOT(ISNUMBER(D2)),NOT(ISNUMBER(E2))),"",IF(E2<D2,"",IF($K$2="Yes",COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,"<="&E2,$E$2:$E$40,">="&D2)-1,COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,"<"&E2,$E$2:$E$40,">"&D2)-IF(E2>D2,1,0)))))`.
  - G2 `=IF(A2="","",IF(F2="","Check",IF(F2>0,"Conflict","OK")))`.
  - H2 (cross-check, SUMPRODUCT) `=IF(F2="","",IF($K$2="Yes",SUMPRODUCT(($B$2:$B$40=B2)*($D$2:$D$40<=E2)*($E$2:$E$40>=D2))-1,SUMPRODUCT(($B$2:$B$40=B2)*($D$2:$D$40<E2)*($E$2:$E$40>D2))-IF(E2>D2,1,0)))`.
  - K3 `=COUNTIF(G2:G40,"Conflict")`; K4 `=COUNTA(A2:A40)`; K5 `=COUNTIF(G2:G40,"Check")`; K6 `=IF(SUM(F2:F40)=SUM(H2:H40),"OK","CHECK")`.
- Plan change: formulas run to row 40 (not 13) so rows can be added; the settings sit in J/K (K2 handoff switch), not in the table columns. Plan change: in the "No" branch the row itself is subtracted only when end is after start (`-IF(E2>D2,1,0)`); with a plain `-1` a one-day booking (start = end) would show -1 because `<` and `>` do not match the row itself. The SUMPRODUCT second count (column H) is per row, and K6 compares the two column sums. Rows with a blank space, a date typed as text, or an end before the start get Check.
- Verified values (LibreOffice Calc only), conflicts with handoff Yes: BK-101 1, 102 1 (10/05 end = 10/05 start), 103 1, 104 1 (inside 103), 105 0 (same dates as 101 but Booth B), 106 1, 107 2 (overlaps 106 and 108), 108 1, 109 1, 110 1 (identical dates), 111 0, 112 0 (one day booking). K3 9, K4 12, K5 0, K6 OK. All match a Python pairwise test.
- Edge: K2 = No: BK-101 and 102 become OK (0), the rest unchanged; K3 7; matches Python (`s2 < e and e2 > s`).
- Edge: end before start (BK-105 changed to 10/09 to 10/04): F blank, G Check, K5 1, K3 stays 9; K6 OK.
- Edge: start date typed as text (BK-102 D3 = '10/05/2026'): that row shows Check, and BK-101 silently loses its conflict (shows 0 OK), K3 7, K5 1, K6 still OK. This is the Troubleshooting case: COUNTIFS ignores the text date for the other rows.
- Edge: blank row inside the table (row 7 cleared): stays blank, others recalculated (BK-107 now 1, BK-108 1), K3 8, K4 11.
- Edge: booking with space and ID but no dates: Check. One day booking (start = end) added inside BK-103/BK-104 with K2 = No: it shows 2, and BK-103 and BK-104 each show 2. A one day booking alone on a free day shows 0 OK.
- Limits: date math only, no lease, contract or deposit terms; the space dropdown has three fixed names (edit the validation list for your own); a booking is flagged on every row involved; a date typed as text hides a conflict from the other rows (only the text row shows Check); times of day are not considered (whole dates).

---

## 2. 2026-10-17-team-02-cash-drawer-count-over-short-farmers-market-excel-google-sheets
- Title: (hook-writer) example: How to Count a Cash Drawer and Find Over or Short in Excel and Google Sheets
- Template: `templates/tidy-tabs-cash-drawer-count.xlsx` (function `cash_drawer_count`). Sheets: Count, How to use (portrait).
- Inputs (yellow): Count C2:C11 (count per denomination: $100, $50, $20, $10, $5, $1 bills, quarter, dime, nickel, penny; sample 0, 0, 6, 9, 14, 23, 18, 12, 9, 20); F2 starting float $150.00; F3 cash sales $161.85. B2:B11 hold the values 100, 50, 20, 10, 5, 1, 0.25, 0.10, 0.05, 0.01.
- Formulas:
  - D2 `=IF(C2="","",IF(NOT(ISNUMBER(C2)),"Check",IF(AND(C2>=0,C2=INT(C2)),B2*C2,"Check")))` (filled to D11); C12 `=SUM(C2:C11)` (bills and coins); D12 `=SUM(D2:D11)`.
  - F4 `=F2+F3`; F5 `=ROUND(D12-F4,2)`; F6 `=IF(COUNTIF(D2:D11,"Check")>0,"Check counts",IF(F5=0,"Balanced",IF(F5>0,"Over","Short")))`; F7 `=SUMPRODUCT(B2:B11,C2:C11)`; F8 `=IF(ROUND(F7-D12,2)=0,"OK","CHECK")`.
- Plan change: D2 is wrapped to flag a negative, fractional or text count as Check (plan was `=B2*C2`); the first draft used `AND(ISNUMBER(C2),C2>=0,...)` which returned #VALUE! for a text count in LibreOffice because AND does not skip the other tests, so it was rewritten with nested IFs. Labels sit in column E and values in F (as in the plan's addresses).
- Verified values (LibreOffice Calc only): amounts $120.00, $90.00, $70.00, $23.00, $4.50, $1.20, $0.45, $0.20; D12 $309.35; C12 111; expected F4 $311.85; F5 -2.50, F6 Short; F7 309.35, F8 OK. Python Decimal agrees. Three more fictional drawers: A (float 200, sales 329.10, total 529.10) Balanced, 0; B (float 100, sales 220.52, total 320.51) Short -0.01; C (same counts, sales 220.50) Over 0.01. All match Decimal.
- Edge: balanced sample drawer (sales 159.35): F5 0, Balanced. Float $0.00 with sales equal to the count: Balanced. 3 dimes vs $0.30 sales: Balanced.
- Edge: all counts blank, float 150, sales 0: D12 0, F5 -150, Short (blank is treated as none); float 0 and sales 0: Balanced, 0.
- Edge: count typed as text ('six' in the $20 row): D shows Check, F6 Check counts, D12 and F7 skip it (189.35 both, F8 OK). Negative count (-6): Check in D, F6 Check counts, F8 CHECK (SUMPRODUCT includes the negative, SUM does not). Fraction count (6.5): Check, F8 CHECK.
- Edge on ROUND: with `=D12-F4` (no ROUND) the one cent over case gave 0.00999999999999091 and `=0` comparisons false; ROUND gave 0.01. The balanced cases tried (three drawers) came out exactly 0 in LibreOffice without ROUND too, so the post must not claim that every balanced drawer shows float noise, only that one cent can come out as 0.00999999999999091 (the Troubleshooting case; Excel was not run).
- Limits: counts and arithmetic only, no accounting or reconciliation advice; a float or sales value typed as text makes F4 an error (not guarded); one drawer per file; rolled coins and checks not covered.

---

## 3. 2026-10-17-team-03-date-of-nth-weekday-in-month-second-saturday-excel-google-sheets
- Title: (hook-writer) example: How to Find the Date of the 2nd Saturday (or Any Nth Weekday) of a Month in Excel and Google Sheets
- Template: `templates/tidy-tabs-nth-weekday-of-month.xlsx` (function `nth_weekday_of_month`). Sheets: Dates, Schedule, How to use (portrait).
- Inputs (yellow): Dates B2 year 2026, B3 month 11, B4 weekday Saturday (dropdown from E2:E8, Sunday to Saturday), B5 n = 2.
- Formulas (Dates):
  - B6 `=IFERROR(MATCH(B4,$E$2:$E$8,0),"Check weekday")` (1 = Sunday to 7 = Saturday).
  - B7 `=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B5),ISNUMBER(B6)),IF(AND(B2>=1900,B2<=9999,B2=INT(B2),B3>=1,B3<=12,B3=INT(B3),B5>=1,B5<=5,B5=INT(B5)),"Yes","No"),"No")`.
  - B8 `=IF(B7<>"Yes","Check inputs",IF(MONTH(X)=B3,X,"No such date"))` with X = `DATE(B2,B3,1)+MOD(B6-WEEKDAY(DATE(B2,B3,1),1),7)+7*(B5-1)` written out in full in the cell.
  - B9 (last such weekday) `=IF(B7<>"Yes","Check inputs",EOMONTH(DATE(B2,B3,1),0)-MOD(WEEKDAY(EOMONTH(DATE(B2,B3,1),0),1)-B6,7))`.
  - B10 checks that B8 and B9 fall on the chosen weekday (OK or CHECK).
- Formulas (Schedule, rows 2 to 13): A2 `=IF(AND(ISNUMBER(Dates!$B$2),ISNUMBER(Dates!$B$5),ISNUMBER(Dates!$B$6),Dates!$B$2>=1900,Dates!$B$2<=9999,Dates!$B$2=INT(Dates!$B$2),Dates!$B$5>=1,Dates!$B$5<=5,Dates!$B$5=INT(Dates!$B$5)),DATE(Dates!$B$2,1,1),"")` (A3 `=IF(A2="","",EDATE(A2,1))`, shown as mmmm yyyy); B2 `=IF(A2="","",IF(MONTH(Y)=MONTH(A2),Y,"No such date"))` with Y = `A2+MOD(Dates!$B$6-WEEKDAY(A2,1),7)+7*(Dates!$B$5-1)`; C2 `=IF(ISNUMBER(B2),Dates!$B$4,"")`.
- Plan change: an added Inputs OK? cell (B7) guards month 13, n outside 1 to 5, text and fractions before DATE is called, so the results say Check inputs instead of rolling into the next year. The Schedule follows the weekday and n set on Dates (default 2nd Saturday) and ignores the month cell. Dropdown list sits in E2:E8.
- Verified values (LibreOffice Calc only): 2nd Saturday of 11/2026 is 11/14/2026; last Saturday 11/28/2026; B7 Yes; B10 OK. Schedule for 2nd Saturday of 2026: 01/10, 02/14, 03/14, 04/11, 05/09, 06/13, 07/11, 08/08, 09/12, 10/10, 11/14, 12/12/2026; all match Python `calendar`.
- Verified by exhaustive run: the Dates formulas (B8 and B9 copied into a scratch sheet) for 2026, 2027 and 2028, all 12 months, all 7 weekdays, n = 1 to 5 (1,260 combinations): 0 mismatches against Python `calendar` for both the Nth weekday (or No such date) and the last weekday. The Schedule tab for 2026 and 2027 with n = 1, 2, 4 and 5 (96 cells): 0 mismatches. 5th Saturday does not exist in 2026 for months 2, 3, 4, 6, 7, 9, 11, 12 and in 2027 for 2, 3, 4, 6, 8, 9, 11, 12.
- Edge: first of the month already the weekday (Thursday 10/01/2026, n = 1): 10/01/2026. 5th Saturday of 11/2026: No such date (Schedule row for November also No such date), last Saturday 11/28/2026. February 2028 (leap year), 5th Tuesday: 02/29/2028; 4th Tuesday 02/22/2028; 1st Tuesday 02/01/2028. February 2027 5th Monday: No such date, last Monday 02/22/2027.
- Edge: month typed as 13 or 0: B7 No, B8 and B9 Check inputs, B10 blank (this is the Troubleshooting case; DATE(2026,13,1) would otherwise roll to 01/01/2027). n = 6, 0 or 2.5: Check inputs. A year typed as text '2026': Check inputs and the Schedule is blank. A misspelled weekday: B6 Check weekday, Check inputs.
- Limits: date math only, no holidays or market rules; years 1900 to 9999 only; n above 4 often gives No such date; weekday numbering follows WEEKDAY(...,1) (Sunday = 1) and the weekday must be spelled as in the dropdown list.

---

## 4. 2026-10-17-team-04-format-phone-numbers-substitute-left-mid-right-excel-google-sheets
- Title: (hook-writer) example: How to Format Phone Numbers as (555) 123-4567 in Excel and Google Sheets
- Template: `templates/tidy-tabs-format-phone-numbers.xlsx` (function `format_phone_numbers`). Sheets: Phones, How to use (portrait).
- Inputs (yellow): Phones A2:A40 (phone as typed; columns A to C formatted as text; 12 sample rows, all 212 and 415 numbers in the 555-01xx style).
- Formulas (filled to row 40):
  - B2 `=IF(A2="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2," ",""),"-",""),"(",""),")",""),".",""),"+",""))`.
  - C2 `=IF(B2="","",IF(AND(LEN(B2)=11,LEFT(B2,1)="1"),MID(B2,2,10),B2))`.
  - D2 `=IF(A2="","",IF(AND(LEN(C2)=10,SUMPRODUCT(--ISNUMBER(FIND(MID(C2,{1,2,3,4,5,6,7,8,9,10},1),"0123456789")))=10),"("&LEFT(C2,3)&") "&MID(C2,4,3)&"-"&RIGHT(C2,4),"Check"))`.
  - G2 `=COUNTA(A2:A40)`; G3 `=SUMPRODUCT(--(LEN(D2:D40)=14))`; G4 `=COUNTIF(D2:D40,"Check")`; G5 `=IF(G3+G4=G2,"OK","CHECK")`.
- Plan change: the all-digits test is `SUMPRODUCT(--ISNUMBER(FIND(MID(C2,{1..10},1),"0123456789")))=10` instead of `ISNUMBER(--C2)`, because `--"5551e34567"` converts to a number (scientific notation) and would be wrongly formatted; the new test rejects it. Formatted count uses length 14 instead of a wildcard. Check cell is G5.
- Verified values (LibreOffice Calc only), A to D: (212) 555-0101 and 212.555.0102 and +1 212-555-0103 and the number 2125550104 (typed as a number) and 212 555 0105 and 1-212-555-0106 and (415)555-0107 and 415.555.0108 all give (212) 555-01xx or (415) 555-01xx; 212-555-010 (9 digits), 212-555-0111 x12, 21255501130 (11 digits, not starting with 1) and "call after 5" give Check. G2 12, G3 8, G4 4, G5 OK. All match a Python regex clean-up.
- Edge (18 more cases, all match Python): 5551234567 (text), (555)123-4567, 1 (555) 123-4567, 1-555-123-4567, 15551234567, " 555 123 4567 " (extra spaces) all give (555) 123-4567. Check: 555-123-456, 555-123-4567 x12, 555 1234 5678, 555-ABC-4567, 5551e34567, +44 20 7946 0958, 25551234567, 555/123/4567 (slash is not stripped), a cell with only spaces. A blank row stays blank (not counted). G2 17, G3 8, G4 9, G5 OK.
- Edge: leading zero area code typed as text "0551234567" gives (055) 123-4567; the same number typed as a real number loses the zero (551234567), 9 digits, Check. A non-breaking space (CHAR(160)) inside the number is not removed: Check (the post can link the non-breaking space template).
- Limits: format only, no check that a number exists or can be called; US 10 digit numbers (with an optional leading 1) only; extensions, letters and other countries show Check; slashes and non-breaking spaces are not stripped.

---

## 5. 2026-10-17-team-05-scale-recipe-batch-quantities-absolute-references-excel-google-sheets
- Title: (hook-writer) example: How to Scale a Recipe or Batch Formula Up or Down in Excel and Google Sheets
- Template: `templates/tidy-tabs-scale-recipe-batch.xlsx` (function `scale_recipe_batch`). Sheets: Recipe, How to use (portrait).
- Inputs (yellow): Recipe C1 base yield 12 (jars); C2 target yield 30; A6:C30 ingredient, base quantity, unit (unit dropdown oz/cup/tsp/each); 8 sample rows 6 to 13: Strawberries 64 oz, Granulated sugar 6 cup, Lemon juice 0.25 cup, Powdered pectin 4 tsp, Lemons 3 each, Vanilla extract 1.5 tsp, Salt 0.25 tsp, Jar lids 12 each.
- Formulas (filled to row 30):
  - C3 `=IF(AND(ISNUMBER(C1),C1>0),IF(ISNUMBER(C2),C2/C1,0),"")` (shown as 2.50x).
  - C4 `=IF(C3="","Check base yield",IF(C3=0,"Check target yield",IF(SUMPRODUCT((A6:A30<>"")*(1-ISNUMBER(B6:B30)))>0,"Check quantities","OK")))`.
  - D6 `=IF(OR(C$3="",NOT(ISNUMBER(B6))),"",ROUND(B6*$C$3,2))`; E6 `=IF(D6="","",IF(C6="each",ROUNDUP(ROUND(B6*$C$3,6),0),ROUND(B6*$C$3,2)))`.
- Plan change: the layout is Ingredient A, Base quantity B, Unit C, Scaled quantity D, Buy whole units E (the plan's `D6="each"` would not match its own layout, so the unit is tested in C6); E6 uses `ROUNDUP(ROUND(B6*$C$3,6),0)` so float noise such as 3 x 1.1 = 3.3000000000000003 does not round up to 4 incorrectly (extra ROUND, same result for the sample); D6 mixes `C$3` only in the blank test and `$C$3` in the product. A target of 0 or text is shown as factor 0 with C4 Check target yield (the Troubleshooting case), instead of an error.
- Verified values (LibreOffice Calc only), factor 2.5: Strawberries 160, sugar 15, lemon juice 0.63 (0.625 rounded), pectin 10, lemons 7.5 with Buy whole units 8, vanilla 3.75, salt 0.63, jar lids 30; C3 2.5, C4 OK. Python Decimal (round half up, ceiling for each) agrees on all 8 rows and both columns.
- Edge: factors 0.5 (target 6), 1 (target 12), 5 (60) and 10 (120): all 8 rows match Python in both columns, C4 OK. Target 11 on base 10 (factor 1.1): lemons 3.3 with Buy whole units 4, jar lids 13.2 and 14; matches Python.
- Edge: target blank or 0: C3 0, every quantity 0, C4 Check target yield. Target typed as text ('thirty'): same, factor 0. Base yield 0 or blank: C3 blank, D and E blank, C4 Check base yield, no #DIV/0!.
- Edge: a quantity typed as text ('six' for sugar): that row has no scaled quantity, the other rows are right, C4 Check quantities.
- Limits: fractions typed as decimals (1/3 cup as 0.33); linear scaling only, so cooking time, baking chemistry, set points and food safety at larger sizes are not covered; Buy whole units only rounds up the unit each; row 6 to 30 only.
