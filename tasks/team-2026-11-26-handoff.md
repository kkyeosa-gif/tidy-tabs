# Handoff 2026-11-26: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-26.md (researcher and hook stages were done in the same run: tasks/team-topics-2026-11-26.md, tasks/topics-seed.md, tasks/hooks-2026-11-26.md). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`spend_and_save_tiers` with `SPEND_TIERS_LAST`/`SPEND_LAST`, `holiday_hours_schedule` with `HOURS_EXC_LAST`/`HOURS_SCHED_LAST`/`HOURS_CHECK_LAST` and helper `_day_times`, `average_daily_balance` with `ADB_TX_LAST`/`ADB_DAYS_LAST`, `daily_sales_year_heatmap` with `HEAT_LOG_LAST`, `price_volume_change` with `PV_ROWS_LAST`/`PV_LOG_LAST`, helper `_fill_cells`; registered in `__main__` after `fiscal_calendar_445()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-26 템플릿 제작/검증 (LibreOffice Calc 24.2.7.2)"). Heroes are not made (designer; trial renders in scratch only, calls below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 1,171 recalculated workbooks, 3,861,313 cell comparisons, 0 mismatches against independent Python models in the final runs, 0 error values in any recalculated workbook. A mutation run (one wrong formula per template, 5 random sets each plus the edge cases, run on the drafts before the last layout and column C fixes) produced mismatches in all five (401 / 3,783 / 40 / 3,537 / 410: tier `"<"&` instead of `"<="&`; minute `INT(C5*1440)`; ADB replaced by (opening + ending)/2; grid date without the -1; volume effect valued at this period's price), so the harness catches wrong formulas. All five files exported to PDF: every page US Letter (792 x 612 landscape data tabs, 612 x 792 How to use), no error values, Err: or ### in the PDF text.

Repo note: the whole script was **not** run. The five files were built by importing scripts/build-templates.py with `sys.dont_write_bytecode = True` and calling only the five functions, so no older template and no scripts/__pycache__/*.pyc was rewritten (nothing to restore; `git status` shows no modified .xlsx or .pyc). `git status` after all three stages: scripts/build-templates.py (837 lines added, 0 removed), tasks/topics-seed.md (7 lines added), new tasks/team-topics-2026-11-26.md, tasks/hooks-2026-11-26.md, this handoff, tasks/briefs.md, five new .xlsx.

Conventions in all five: How to use tab last; US Letter, fit to one page wide, header row repeated on the data tabs; data tabs landscape; yellow = inputs, gray = formulas; red = any Status starting "Check" and any CHECK / settings cell that is not OK; light orange = warnings that are not errors (Outside promo dates, Closed, Below zero, Other year, Outside both periods, New product, Dropped); light green = highlights (Nudge Yes, Special hours, Open, the ADB, the three effects). Functions used in the five files: AND, COUNT, COUNTA, COUNTIF, COUNTIFS, DATE, DAY, IF, INDEX, INT, ISNUMBER, LEFT, MATCH, MAX, MIN, MOD, MONTH, N, NOT, OR, ROUND, ROWS, SUM, SUMIFS, SUMPRODUCT, TEXT, TRIM, WEEKDAY, YEAR. No TODAY, FILTER, UNIQUE, SORT, LET, LAMBDA, REDUCE, XLOOKUP, TEXTSPLIT, MINIFS/MAXIFS, SEQUENCE; 0 `_xlfn` in the five saved files and 0 after the LibreOffice round trip.

Help links: checked 2026-11-26 with `curl -sIL -o /dev/null -w '%{http_code}'`. **Every Microsoft link returned 200. Every Google link returned 404 to that HEAD check**, but 200 to a plain GET (`curl -sL -o /dev/null -w '%{http_code}'`), and the GET page title names the function (for example "SUMIFS function - Google Docs Editors Help"). The publisher checklist says `curl -sIL` must give 200, so the publisher should decide whether a GET 200 is enough for the Google pages (earlier posts already link the same Google pages).

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-26-team-01-spend-and-save-tiered-discount-amount-to-reach-next-tier-countif-index-excel-google-sheets | tidy-tabs-spend-and-save-tiers.xlsx | 228 workbooks, 1,150,260 comparisons |
| 2 | 2026-11-26-team-02-holiday-hours-schedule-open-or-closed-at-a-date-and-time-weekday-index-match-excel-google-sheets | tidy-tabs-holiday-hours-schedule.xlsx | 280 workbooks, 408,520 comparisons |
| 3 | 2026-11-26-team-03-average-daily-balance-from-transaction-list-end-of-day-balance-sumifs-excel-google-sheets | tidy-tabs-average-daily-balance.xlsx | 226 workbooks, 334,480 comparisons |
| 4 | 2026-11-26-team-04-year-at-a-glance-daily-sales-heatmap-calendar-grid-sumifs-color-scale-excel-google-sheets | tidy-tabs-daily-sales-year-heatmap.xlsx | 217 workbooks, 1,182,433 comparisons |
| 5 | 2026-11-26-team-05-sales-change-vs-last-year-price-effect-and-volume-effect-by-product-sumifs-excel-google-sheets | tidy-tabs-sales-change-price-volume.xlsx | 220 workbooks, 785,620 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets." For the holiday hours post add: "The MOD time-of-day rounding problem was reproduced in LibreOffice Calc (and in Python's float math) only." For the heatmap post add: "The color scale was seen only in a LibreOffice PDF export."

Hooks check (tasks/hooks-2026-11-26.md): the templates were built first, so every number and formula in the openings was taken from the recalculated templates; all confirmed. Final titles in the hooks file differ from the topic titles for posts 1, 2, 3 and 5 (listed per post below). No "to confirm" items are open.

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=..., name=...)` and `short_sheet_copy(src, sheet, last_row, name)`). Trial renders at 1600x900 were made under /usr/bin/python3 (PYTHONDONTWRITEBYTECODE=1) into scratch only (/tmp/claude-0/-home-user-tidy-tabs/b45b4175-3143-56f9-81ce-cc5c1349391b/scratchpad/hero/); nothing was written to images/. All calls below rendered legibly.

---
## 1. spend-and-save-tiered-discount-amount-to-reach-next-tier-countif-index-excel-google-sheets
Post: posts/ready/2026-11-26-team-01-spend-and-save-tiered-discount-amount-to-reach-next-tier-countif-index-excel-google-sheets.md
Title (hooks final): How to Calculate a Tiered Discount Based on Order Total in Excel and Google Sheets
Template: templates/tidy-tabs-spend-and-save-tiers.xlsx (tabs: Orders, Tiers, How to use)

Layout
- Orders: A1 Nudge if within / B1 (input, $20.00), C1 note, A2 note. Header row 4, rows 5:504: A Order #, B Order date, C Merchandise subtotal, D Not eligible (gift cards etc.) (inputs); E Eligible subtotal, F Tier reached, G Dollars off, H Total after discount, I Discount % of order, J Next tier at, K Add to reach next tier, L Next tier takes off, M Nudge?, N Status. Summary P4:Q13: Orders in the promo (Q4), Orders that reached a tier (Q5), Merchandise in the promo (Q6), Total dollars off (Q7), Discount % of promo merchandise (Q8), Orders within the nudge amount (Q9), Outside promo dates (Q10), Rows showing Check (Q11), Typed below row 504 (Q12), CHECK (Q13). Row 4 repeated.
- Tiers: A1 Promo starts / B1 (11/27/2026), A2 Promo ends / B2 (12/01/2026) (inputs). Header row 4, rows 5:14: A Spend at least, B Dollars off (inputs), C Tier, D Off at the threshold, E Status. Summary G4:H8: Tiers (H4), Rows showing Check (H5), Promo dates (H6), Typed below row 14 (H7), CHECK (H8).

Key formulas (exact text in the file)
- Orders E5 `=IF(OR(N5="",LEFT(N5,5)="Check"),"",ROUND(C5-N(D5),2))`
- F5 `=IF(E5="","",IF(N5="OK",COUNTIF(Tiers!$A$5:$A$14,"<="&E5),0))`
- G5 `=IF(F5="","",IF(F5=0,0,INDEX(Tiers!$B$5:$B$14,F5)))`; H5 `=IF(G5="","",ROUND(C5-G5,2))`; I5 `=IF(G5="","",IF(C5=0,0,G5/C5))`
- J5 `=IF(OR(N5<>"OK",F5>=Tiers!$H$4),"",INDEX(Tiers!$A$5:$A$14,F5+1))`; K5 `=IF(J5="","",ROUND(J5-E5,2))`; L5 `=IF(J5="","",INDEX(Tiers!$B$5:$B$14,F5+1))`; M5 `=IF(K5="","",IF(K5<=$B$1,"Yes",""))`
- N5 `=IF(AND(TRIM(A5)="",B5="",C5="",D5=""),"",IF(Tiers!$H$8<>"OK","Check: Tiers tab",IF(NOT(ISNUMBER(B5)),"Check: date blank or text",IF(NOT(ISNUMBER(C5)),"Check: subtotal blank or text",IF(AND(D5<>"",NOT(ISNUMBER(D5))),"Check: not-eligible amount is text",IF(OR(C5<0,N(D5)<0),"Check: negative amount",IF(N(D5)>C5,"Check: not-eligible is more than the subtotal",IF(OR(INT(B5)<Tiers!$B$1,INT(B5)>Tiers!$B$2),"Outside promo dates","OK"))))))))`
- Tiers E6 (rows 6 to 14; E5 has no row-above tests) `=IF(AND(A6="",B6=""),"",IF(OR(NOT(ISNUMBER(A6)),NOT(ISNUMBER(B6))),"Check: amount blank or text",IF(A6<=0,"Check: threshold must be above 0",IF(OR(B6<=0,B6>=A6),"Check: dollars off must be above 0 and below the threshold",IF(OR(ROUND(A6,2)<>A6,ROUND(B6,2)<>B6),"Check: more than 2 decimals",IF(OR(A5="",B5=""),"Check: blank row above",IF(A6<=A5,"Check: thresholds must go up",IF(B6<=B5,"Check: dollars off must go up","OK"))))))))`; D5 `=IF(E5<>"OK","",B5/A5)`; H8 `=IF(AND(H4>0,H5=0,H6="OK",H7=0),"OK","Check")`
- Q8 `=IF(Q6=0,"",Q7/Q6)`; Q13 `=IF(AND(Tiers!H8="OK",ISNUMBER(B1),Q11=0,Q12=0),"OK","Check")`

Sample values (fictional; recalculated). Tiers $10.00 off $50.00 (20.0% at the threshold), $25.00 off $100.00 (25.0%), $40.00 off $150.00 (26.7%), $75.00 off $250.00 (30.0%); promo 11/27/2026 to 12/01/2026; nudge $20.00.
| Order | Date | Subtotal | Not eligible | Eligible | Tier | Off | After | Next at | Add | Next off | Nudge | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1101 | 11/26/2026 | $112.40 | | $112.40 | 0 | $0.00 | $112.40 | | | | | Outside promo dates |
| 1102 | 11/27/2026 | $49.99 | | $49.99 | 0 | $0.00 | $49.99 | $50.00 | $0.01 | $10.00 | Yes | OK |
| 1103 | 11/27/2026 | $50.00 | | $50.00 | 1 | $10.00 | $40.00 | $100.00 | $50.00 | $25.00 | | OK |
| 1104 | 11/27/2026 | $99.99 | | $99.99 | 1 | $10.00 | $89.99 | $100.00 | $0.01 | $25.00 | Yes | OK |
| 1105 | 11/27/2026 | $100.00 | | $100.00 | 2 | $25.00 | $75.00 | $150.00 | $50.00 | $40.00 | | OK |
| 1106 | 11/28/2026 | $120.10 | $20.10 | $100.00 | 2 | $25.00 | $95.10 | $150.00 | $50.00 | $40.00 | | OK |
| 1107 | 11/28/2026 | $136.50 | | $136.50 | 2 | $25.00 | $111.50 | $150.00 | $13.50 | $40.00 | Yes | OK |
| 1108 | 11/28/2026 | $148.25 | | $148.25 | 2 | $25.00 | $123.25 | $150.00 | $1.75 | $40.00 | Yes | OK |
| 1109 | 11/29/2026 | $175.00 | $50.00 | $125.00 | 2 | $25.00 | $150.00 | $150.00 | $25.00 | $40.00 | | OK |
| 1110 | 11/29/2026 | $236.80 | | $236.80 | 3 | $40.00 | $196.80 | $250.00 | $13.20 | $75.00 | Yes | OK |
| 1111 | 11/30/2026 | $262.15 | | $262.15 | 4 | $75.00 | $187.15 | | | | | OK |
| 1112 | 11/30/2026 | $38.00 | | $38.00 | 0 | $0.00 | $38.00 | $50.00 | $12.00 | $10.00 | Yes | OK |
| 1113 | 12/01/2026 | $84.60 | | $84.60 | 1 | $10.00 | $74.60 | $100.00 | $15.40 | $25.00 | Yes | OK |
| 1114 | 12/01/2026 | $310.00 | $25.00 | $285.00 | 4 | $75.00 | $235.00 | | | | | OK |
| 1115 | 12/02/2026 | $160.00 | | $160.00 | 0 | $0.00 | $160.00 | | | | | Outside promo dates |
Discount % of order (I): 1104 10.0%, 1105 25.0%, 1106 20.8% (of $120.10), 1111 28.6%, 1114 24.2%. Q4:Q13: 13, 11, $1,811.38, $345.00, 19.0%, 7, 2, 0, 0, OK.

Edge cases tested (all matched the Python model)
- Every threshold minus $0.01, exactly, plus $0.01: $49.99 tier 0 ($0.01 to add), $50.00 tier 1, $99.99 tier 1 ($0.01), $100.00 tier 2, $149.99 tier 2 ($0.01), $150.00 tier 3, $249.99 tier 3 ($0.01), $250.00 and $250.01 tier 4 (no next tier), $0.00 and $0.01 tier 0, $99,999.99 tier 4; $100.00 with $100.00 not eligible = eligible $0.00 tier 0; $100.10 minus $0.10 and $250.30 minus $0.30 land exactly on the tier.
- Nudge 0, blank and "20" as text (text makes every row Yes and CHECK Check because B1 must be a number); one tier; ten tiers; a value typed on Tiers row 15 (H7 1, every order Check: Tiers tab).
- Tier table checks: unsorted, same threshold twice, discount not going up, discount equal to the threshold, discount 0, negative threshold, "50" as text, blank discount, a blank row between tiers, 3 decimals, no tiers at all (each its own Check; Orders all Check: Tiers tab).
- Promo one day, reversed, start as text, start with a time (Check); order dates with times at the ends (12/01/2026 11:59 PM counted, 11/26/2026 11:59 PM outside).
- Bad orders: text date, blank date, text subtotal, no subtotal, text not-eligible, negative subtotal, negative not-eligible, not-eligible above the subtotal, an order with no number (counted), a number only, spaces only (blank row), not-eligible equal to the subtotal (eligible $0.00); a row typed on row 505 and on D3000 (Q12 1, CHECK Check).
- A full sheet of 500 orders, and 200 random sets (1 to 10 tiers with .00/.50/.99 thresholds, some bad tier cells, random promo windows, 1 to 120 orders with orders exactly at thresholds, excluded amounts, text dates and excluded above subtotal).
- LibreOffice probe (scratch, not the template): `=A1-B1` with 120.10 and 20.10 gave exactly 100, so no floating-point tier miss was reproduced; the ROUND in E5 is a safeguard only. Do not claim a rounding failure.

Pitfalls the post must state
- Dollars off come off the merchandise subtotal; shipping and sales tax are not in the sheet (no tax advice). Sample orders are fictional.
- An order exactly at a threshold reaches it ("Spend at least"). Promo dates include both days.
- Thresholds must go up from row 5 with no gaps, or every order shows Check: Tiers tab.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Every order says Check: Tiers tab: a tier row is out of order, blank in the middle, text, or the promo dates are reversed.
- An order gets no discount: Outside promo dates (1101 on 11/26, 1115 on 12/02), or the excluded amount took it under $50.00.
- Nudge? is Yes on every row: B1 typed as text.

Hero: `render(short_sheet_copy(t+"tidy-tabs-spend-and-save-tiers.xlsx","Orders",19,"ss"),"Orders",hide_cols="I",name="ss")` (the 15 orders with tiers, dollars off, next tier, add amount and the green Yes cells, plus the summary; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): COUNTIF https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34 and https://support.google.com/docs/answer/3093480; INDEX https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd and https://support.google.com/docs/answer/3098242; ROUND https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c and https://support.google.com/docs/answer/3093440.

---
## 2. holiday-hours-schedule-open-or-closed-at-a-date-and-time-weekday-index-match-excel-google-sheets
Post: posts/ready/2026-11-26-team-02-holiday-hours-schedule-open-or-closed-at-a-date-and-time-weekday-index-match-excel-google-sheets.md
Title (hooks final): How to Make a Holiday Hours Schedule in Excel and Google Sheets
Template: templates/tidy-tabs-holiday-hours-schedule.xlsx (tabs: Hours, Schedule, Check a time, How to use)

Layout
- Hours: A1:A2 notes. Header row 4. Weekly A5:A11 Sunday to Saturday (labels), B Opens, C Closes (inputs), D Status. Dated exceptions rows 5:44: F Date, G Note, H Opens, I Closes (inputs), J Status. Summary L4:M7: Weekly rows showing Check (M4), Date rows showing Check (M5), Typed below row 44 (M6), CHECK (M7).
- Schedule: A1 First date / B1 (11/23/2026), A2 Days to list / B2 (42; 1 to 62) (inputs). Header row 4, rows 5:66: A Date, B Day, C Note, D Opens, E Closes, F Hours open, G Status, H Date row on Hours. Summary J4:K10: Settings (K4), Days listed (K5), Open days (K6), Closed days (K7), Days with special hours or closed for a date (K8), Hours open in the list (K9), CHECK (K10).
- Check a time: A1:A2 notes. Header row 4, rows 5:104: A Date and time (input), B Date, C Time, D Opens, E Closes, F Open?, G Minutes until it opens or closes, H Status, I Date row on Hours, J Minute of the day. Summary L4:M8: Times checked (M4), Open (M5), Rows showing Check (M6), Typed below row 104 (M7), CHECK (M8).
- Builder decisions: same-day hours only (closing at or before opening shows Check); a blank time pair means closed; an Opens/Closes cell that is text is shown blank (and the Hours tab says Check). Times are compared in whole minutes after rounding to whole seconds.

Key formulas (exact text in the file)
- Schedule H5 `=IF(A5="","",IF(COUNTIF(Hours!$F$5:$F$44,A5)=0,0,MATCH(A5,Hours!$F$5:$F$44,0)))`
- Schedule D5 `=IF(A5="","",IF(H5>0,IF(ISNUMBER(INDEX(Hours!$H$5:$H$44,H5)),INDEX(Hours!$H$5:$H$44,H5),""),IF(ISNUMBER(INDEX(Hours!$B$5:$B$11,WEEKDAY(A5))),INDEX(Hours!$B$5:$B$11,WEEKDAY(A5)),"")))`; E5 the same with Hours!$I and Hours!$C
- Schedule A5 `=IF($K$4<>"OK","",IF(ROWS($A$5:A5)>$B$2,"",$B$1+ROWS($A$5:A5)-1))`; B5 `=IF(A5="","",TEXT(A5,"dddd"))`; C5 `=IF(A5="","",IF(H5=0,"",INDEX(Hours!$G$5:$G$44,H5)&""))`; F5 `=IF(OR(D5="",E5=""),"",ROUND((E5-D5)*24,2))`
- Schedule G5 `=IF(A5="","",IF(Hours!$M$7<>"OK","Check: Hours tab",IF(H5>0,IF(D5="","Closed (holiday)","Special hours"),IF(D5="","Closed","Regular hours"))))`
- Check a time C5 `=IF(H5<>"OK","",ROUND(MOD(A5,1)*86400,0)/86400)`; J5 `=IF(B5="","",INT(ROUND(C5*86400,0)/60))`; B5 `=IF(H5<>"OK","",INT(A5))`; I5 like Schedule H5 with B5; D5/E5 like Schedule with B5 and I5
- F5 `=IF(B5="","",IF(D5="",IF(I5>0,"Closed (holiday)","Closed today"),IF(J5<ROUND(D5*1440,0),"Not open yet",IF(J5<ROUND(E5*1440,0),"Open","Closed for the day"))))`
- G5 `=IF(OR(F5="",D5=""),"",IF(F5="Not open yet",ROUND(D5*1440,0)-J5,IF(F5="Open",ROUND(E5*1440,0)-J5,"")))`; H5 `=IF(A5="","",IF(Hours!$M$7<>"OK","Check: Hours tab",IF(NOT(ISNUMBER(A5)),"Check: date and time is text",IF(A5<1,"Check: no date, time only","OK"))))`
- Hours D5 `=IF(AND(B5="",C5=""),"Closed",IF(OR(B5="",C5=""),"Check: give both times or leave both blank",IF(OR(NOT(ISNUMBER(B5)),NOT(ISNUMBER(C5))),"Check: time is text",IF(OR(B5<0,B5>=1,C5<0,C5>=1),"Check: time only, no date",IF(ROUND(C5*1440,0)<=ROUND(B5*1440,0),"Check: closes before it opens","Open")))))`; J5 adds `IF(F5<>INT(F5),"Check: date has a time"` and `IF(COUNTIF($F$5:$F$44,F5)>1,"Check: date listed twice"` and ends in "Closed" / "Special hours"
- Schedule K6 `=COUNTIF(G5:G66,"*hours")`; K7 `=COUNTIF(G5:G66,"Closed*")`; K9 `=SUM(F5:F66)`

Sample values (fictional; recalculated). Weekly: Sunday 11:00 AM to 4:00 PM, Monday closed, Tuesday to Friday 10:00 AM to 6:00 PM, Saturday 9:00 AM to 5:00 PM. Dates: 11/26/2026 Thanksgiving closed; 11/27/2026 Black Friday 8:00 AM to 8:00 PM; 11/28/2026 Small Business Saturday 9:00 AM to 7:00 PM; 12/21/2026 Holiday week Monday 10:00 AM to 6:00 PM; 12/24/2026 Christmas Eve 10:00 AM to 2:00 PM; 12/25/2026 Christmas Day closed; 12/31/2026 New Year's Eve 10:00 AM to 3:00 PM; 01/01/2027 New Year's Day closed.
- Schedule 11/23/2026 to 01/03/2027: 42 days, 34 open, 8 closed (Mondays 11/23, 11/30, 12/07, 12/14, 12/28 plus the three holidays), 8 days from the date list, 253.00 hours, CHECK OK. Rows: 11/26 Thursday Thanksgiving Closed (holiday); 11/27 Friday Black Friday 8:00 AM 8:00 PM 12.00 Special hours; 11/28 10.00 Special; 11/29 Sunday 5.00 Regular; 11/30 Monday Closed; 12/21 Monday 8.00 Special; 12/24 4.00 Special; 12/25 Closed (holiday); 12/31 5.00 Special; 01/01/2027 Closed (holiday); 01/03/2027 Sunday 5.00 Regular.
| Date and time | Open? | Minutes |
|---|---|---|
| 11/26/2026 12:00 PM | Closed (holiday) | |
| 11/27/2026 7:59 AM | Not open yet | 1 |
| 11/27/2026 8:00 AM | Open | 720 |
| 11/27/2026 7:59 PM | Open | 1 |
| 11/27/2026 8:00 PM | Closed for the day | |
| 11/30/2026 11:00 AM | Closed today | |
| 12/21/2026 5:30 PM | Open | 30 |
| 12/24/2026 1:30 PM | Open | 30 |
| 12/24/2026 2:00 PM | Closed for the day | |
| 12/25/2026 10:00 AM | Closed (holiday) | |
| 12/26/2026 8:45 AM | Not open yet | 15 |
| 12/31/2026 2:45 PM | Open | 15 |
| 01/01/2027 12:00 PM | Closed (holiday) | |
| 01/02/2027 4:59 PM | Open | 1 |
M4:M8: 14, 6, 0, 0, OK.

Edge cases tested (all matched the Python model)
- Every minute of 11/27/2026 (special hours), 12/05/2026 (a regular Saturday) and 12/24/2026 (short day): 4,320 date-times.
- Seconds around the edges: 7:59:29 AM Not open yet, 7:59:30 AM and 7:59:31 AM count as minute 479 too (seconds are rounded to whole seconds, then minutes floored); 7:59:55 PM Open; 8:00:00 PM to 8:00:59 PM Closed for the day.
- Schedule 1, 62, 63 (Check), 0 (Check), 2.5 (Check) days; start as text, with a time, blank (Check); a 62-day list from 02/01/2028 (leap day 02/29/2028); a list across the new year.
- Weekly checks: one time only, text time, a time with a date, closing before opening, the same time twice, every day closed, 12:00 AM to 11:59 PM, half hours.
- Date list checks: a date twice, a text date, a date with a time, one time only, closing before opening, a note only, a text time, no dates at all, a blank row in the middle (allowed), no note, 40 rows; a date typed on row 45 (M6 1, CHECK Check, every Schedule and Check row Check: Hours tab).
- Check a time: text, a time with no date (Check), a blank, a date with no time (12:00 AM, Not open yet), a date-time typed on row 105 (M7 1); 100 rows across 11/23 to 12/22.
- 200 random sets (random weekly hours with closed days, 0 to 26 dated exceptions with some duplicates, 1 to 62 days, 1 to 100 random date-times with and without seconds).
- Trap (LibreOffice probe and a first draft of the template): for date-times 11/27/2026 00:00 to 23:59 written by openpyxl, `=MOD(A,1)<TIME(h,m,0)` was TRUE for 704 of 1,440 minutes and `=INT(MOD(A,1)*1440)` gave the minute before (10:00 AM gave 599); Python's float math gives the same 704. With plain `=MOD(A5,1)` in column C, a LibreOffice PDF render showed 12/21/2026 5:30 PM as 5:29 PM and 12/25/2026 10:00 AM as 9:59 AM. A mutation using `INT(C5*1440)` for J (run on the draft that still had plain MOD in C) gave 3,783 mismatches. The final template rounds to whole seconds first. These are LibreOffice and Python results only.

Pitfalls the post must state
- The holidays are the reader's own list; the sheet knows no holiday dates or rules. Sample shop is fictional.
- Same-day hours only; overnight hours are not supported.
- Open means opening time <= time < closing time, so at closing time the shop shows closed.
- Updating a website or Google Business Profile is done by hand (no import claim).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Every Schedule row says Check: Hours tab: a time pair with one blank, a text time, a date listed twice or a date with a time on the Hours tab.
- A time shows one minute early or "Not open yet" right at opening: the time came from plain MOD (probe above); use the rounded C5.
- A holiday shows regular hours: the date on the Hours tab has a time or is text (Check: date has a time / Check: date blank or text), so it is not matched.

Hero: `render(short_sheet_copy(t+"tidy-tabs-holiday-hours-schedule.xlsx","Schedule",30,"hh"),"Schedule",hide_cols="H",name="hh")` (11/23 to 12/18 with orange closed days, green special hours and the summary; legible). Second card: `render(short_sheet_copy(t+"tidy-tabs-holiday-hours-schedule.xlsx","Check a time",18,"hc"),"Check a time",hide_cols="IJ",name="hc")` (the 14 checks with Open / Not open yet / Closed; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): WEEKDAY https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a and https://support.google.com/docs/answer/3092985; MATCH https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a and https://support.google.com/docs/answer/3093378; INDEX https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd and https://support.google.com/docs/answer/3098242; MOD https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3 and https://support.google.com/docs/answer/3093497; ROUND https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c and https://support.google.com/docs/answer/3093440; TEXT https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c and https://support.google.com/docs/answer/3094139.

---
## 3. average-daily-balance-from-transaction-list-end-of-day-balance-sumifs-excel-google-sheets
Post: posts/ready/2026-11-26-team-03-average-daily-balance-from-transaction-list-end-of-day-balance-sumifs-excel-google-sheets.md
Title (hooks final): How to Calculate Average Daily Balance in Excel and Google Sheets
Template: templates/tidy-tabs-average-daily-balance.xlsx (tabs: Daily, Transactions, How to use)

Layout
- Daily: inputs B1 Period starts (11/01/2026), B2 Period ends (11/30/2026), B3 Opening balance ($3,250.00), B4 Annual rate (optional, 2.00%), B5 Days in a year for the rate (365); C1:C5 notes. Header row 7, rows 8:99 (92 days): A Date, B Day's net amount, C Entries, D End-of-day balance, E Below zero?. Summary J4:K17: Settings (K4), Days in the period (K5), Sum of end-of-day balances (K6), Average daily balance (K7, green), Ending balance (K8), Lowest end-of-day balance (K9), First day at the lowest (K10), Days below zero (K11), Shortcut (opening + ending)/2 (K12), Shortcut days with entries only (K13), Entries counted (K14), Rough interest (K15), Transactions CHECK (K16), CHECK (K17). Row 7 repeated.
- Transactions: A1:A2 notes. Header row 4, rows 5:1004: A Date, B Description, C Amount (+ in, - out) (inputs), D Status. Summary F4:G9: Rows OK, Total of OK rows, Outside the period, Rows showing Check, Typed below row 1004, CHECK.

Key formulas (exact text in the file)
- Daily A8 `=IF($K$4<>"OK","",IF($B$1+ROWS($A$8:A8)-1>$B$2,"",$B$1+ROWS($A$8:A8)-1))`
- B8 `=IF(A8="","",SUMIFS(Transactions!$C$5:$C$1004,Transactions!$A$5:$A$1004,A8,Transactions!$D$5:$D$1004,"OK"))`; C8 `=IF(A8="","",COUNTIFS(Transactions!$A$5:$A$1004,A8,Transactions!$D$5:$D$1004,"OK"))`
- D8 `=IF(A8="","",ROUND($B$3+B8,2))`; D9 `=IF(A9="","",ROUND(D8+B9,2))` (and down); E8 `=IF(A8="","",IF(D8<0,"Below zero",""))`
- K5 `=IF(K4<>"OK","",B2-B1+1)`; K6 `=SUM(D8:D99)`; K7 `=IF(K4<>"OK","",ROUND(K6/K5,2))`; K8 `=IF(K4<>"OK","",INDEX(D8:D99,K5))`; K9 `=IF(K4<>"OK","",MIN(D8:D99))`; K10 `=IF(K4<>"OK","",INDEX(A8:A99,MATCH(K9,D8:D99,0)))`
- K12 `=IF(K4<>"OK","",ROUND((B3+K8)/2,2))`; K13 `=IF(K4<>"OK","",IF(COUNTIF(C8:C99,">0")=0,"",ROUND(SUMIFS(D8:D99,C8:C99,">0")/COUNTIF(C8:C99,">0"),2)))`; K15 `=IF(OR(K4<>"OK",B4=""),"",ROUND(K7*B4/B5*K5,2))`
- K4 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B5)),IF(AND(B1=INT(B1),B2=INT(B2),B2>=B1,B2-B1+1<=92,OR(B4="",AND(ISNUMBER(B4),B4>=0,B4<1)),OR(B5=360,B5=365,B5=366)),"OK","Check"),"Check")`; K17 `=IF(K4<>"OK","Check",IF(AND(K16="OK",ROUND(K8-B3-SUMIFS(Transactions!$C$5:$C$1004,Transactions!$D$5:$D$1004,"OK"),2)=0),"OK","Check"))`
- Transactions D5 `=IF(AND(A5="",TRIM(B5)="",C5=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(C5="","Check: no amount",IF(NOT(ISNUMBER(C5)),"Check: amount is text",IF(OR(A5<Daily!$B$1,A5>Daily!$B$2),"Outside the period (not counted)","OK"))))))`

Sample values (fictional; recalculated). 16 entries, 11/02/2026 to 11/30/2026 (two on 11/18): Etsy payouts $612.40, $534.85, $688.10, $954.30, $1,095.60; market deposits $418.00, $502.25, $726.50, $1,180.75; supplier -$845.17, shipping labels -$96.30, rent -$1,400.00, jars -$1,312.60, card fees -$41.92, packaging -$1,205.40, loan payment -$350.00.
| Date | Net | End-of-day balance |
|---|---|---|
| 11/01/2026 | $0.00 | $3,250.00 |
| 11/02/2026 | $612.40 | $3,862.40 |
| 11/03/2026 | -$845.17 | $3,017.23 |
| 11/10/2026 | -$1,400.00 | $2,473.78 |
| 11/12/2026 | -$1,312.60 | $1,161.18 |
| 11/18/2026 | -$1,247.32 (2 entries) | $1,104.21 |
| 11/21/2026 | $726.50 | $1,830.71 |
| 11/25/2026 | -$350.00 | $2,435.01 |
| 11/28/2026 | $1,180.75 | $3,615.76 |
| 11/30/2026 | $1,095.60 | $4,711.36 |
K4:K17: OK, 30, $76,632.18, **$2,554.41**, $4,711.36, $1,104.21, 11/18/2026, 0, **$3,980.68** (opening + ending)/2, **$2,743.02** (15 days with entries), 16, **$4.20**, OK, OK. Transactions: 16 OK, $1,461.36, 0, 0, 0, OK.

Edge cases tested (all matched the Python model)
- Entries in reverse order (same results); no entries (ADB = opening); a one-day period; 92 days (10/01 to 12/31/2026, K5 92, CHECK OK); 93 days (K4 Check, every day blank, CHECK Check); end before start; start as text or with a time; opening blank (Check); opening -$500.00 (ADB -$1,195.59, lowest -$2,645.79, 27 days below zero, CHECK OK); rate blank (K15 blank), text, 1 (100%, Check), 0 ($0.00); 360 and 366 (OK), 364 (Check); February 2028 with 02/29.
- Below zero: opening $200.00 gives ADB -$495.59, lowest -$1,945.79 first on 11/18/2026, 22 days below zero, (opening + ending)/2 $930.68, entries-only -$306.98, CHECK OK.
- Bad rows: text date, date with a time, no amount, "10.00" as text, a description only, an amount with no date (6 rows Check, CHECK Check, none counted); entries the day before and after the period (Outside the period, not counted); a row typed on row 1005 (G8 1); amounts with 3 decimals (10.005, -0.125: the sheet has no 2-decimal check; the running balance rounds each day; LibreOffice and the model agreed only after the model rounded on 15 significant digits first, see briefs).
- A tie for the lowest balance (first date returned); 1,000 rows over 92 days; 200 random sets (periods 1 to 92 days from 2025 to 2027, openings -$2,000.00 to $20,000.00, rates blank to 24.99%, 360/365/366, up to 200 entries including outside dates and a date-time).

Pitfalls the post must state
- Every calendar day counts; days without entries carry the day before's balance. Entries before the period belong in the opening balance.
- K15 is a rough estimate; banks and card issuers use their own method and the statement is the authority. No advice on fees, accounts or credit. Sample account is fictional.
- Dates only; a date with a time is flagged and not counted.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The ADB looks too high: (opening + ending)/2 gave $3,980.68 against $2,554.41.
- Averaging only days with entries gave $2,743.02.
- An entry is missing: its date has a time or is text, or it is outside B1:B2 (Status says which).
- Everything blank and CHECK Check: more than 92 days, end before start, or B5 not 360/365/366.

Hero: `render(short_sheet_copy(t+"tidy-tabs-average-daily-balance.xlsx","Daily",37,"adb"),"Daily",hide_cols="C",name="adb")` (settings, all 30 days and the summary with the green $2,554.41 and both shortcuts; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): SUMIFS https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b and https://support.google.com/docs/answer/3238496; ROUND https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c and https://support.google.com/docs/answer/3093440; MIN https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152 and https://support.google.com/docs/answer/3094017; INDEX and MATCH as in post 2.

---
## 4. year-at-a-glance-daily-sales-heatmap-calendar-grid-sumifs-color-scale-excel-google-sheets
Post: posts/ready/2026-11-26-team-04-year-at-a-glance-daily-sales-heatmap-calendar-grid-sumifs-color-scale-excel-google-sheets.md
Title (hooks final): How to Make a Year-at-a-Glance Daily Sales Heatmap in Excel and Google Sheets
Template: templates/tidy-tabs-daily-sales-year-heatmap.xlsx (tabs: Year grid, Sales log, How to use)

Layout
- Year grid: A1 Year / B1 (input, 2026), E1:E2 notes. Header row 4: A Month, B:AF days 1 to 31 (numbers), AG Month total, AH Days with sales, AI Best day, AJ Best day sales. Rows 5:16 January to December. Color scale on B5:AF16 (0 white, 75th percentile light green, max dark green). Summary AM4:AN14: Year (AN4), Year total grid (AN5), Log total for the year (AN6), Days with sales (AN7), Average per day with sales (AN8), Best day (AN9), Best day sales (AN10), Best month (AN11), Best month total (AN12), Sales log CHECK (AN13), CHECK (AN14). Row 4 repeated; prints on one landscape page.
- Sales log: A1:A2 notes. Header row 4, rows 5:5004: A Date, B Amount, C Note (inputs), D Status. Summary G4:H9: Rows OK, Total of OK rows, Rows from other years, Rows showing Check, Typed below row 5004, CHECK.
- Builder decisions: the grid shows whole dollars (`#,##0;-#,##0;"-"`); totals keep cents. A negative day shows white.

Key formulas (exact text in the file)
- A5 `=IF($AN$4<>"OK","",DATE($B$1,1,1))` (rows 6 to 16 months 2 to 12), number format mmmm
- B5 (every grid cell) `=IF($A5="","",IF(B$4>DAY(DATE(YEAR($A5),MONTH($A5)+1,1)-1),"",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$A$5:$A$5004,$A5+B$4-1,'Sales log'!$D$5:$D$5004,"OK")))`
- AG5 `=IF($A5="","",SUM(B5:AF5))`; AH5 `=IF($A5="","",COUNTIF(B5:AF5,">0"))`; AI5 `=IF(OR($A5="",COUNTIF(B5:AF5,">0")=0),"",$A5+MATCH(MAX(B5:AF5),B5:AF5,0)-1)`; AJ5 `=IF(AI5="","",MAX(B5:AF5))`
- AN6 `=IF(AN4<>"OK","",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$D$5:$D$5004,"OK",'Sales log'!$A$5:$A$5004,">="&DATE(B1,1,1),'Sales log'!$A$5:$A$5004,"<"&DATE(B1+1,1,1)))`
- AN8 `=IF(OR(AN4<>"OK",N(AN7)=0),"",ROUND(AN5/AN7,2))`; AN9 `=IF(OR(AN4<>"OK",COUNT(AJ5:AJ16)=0),"",INDEX(AI5:AI16,MATCH(MAX(AJ5:AJ16),AJ5:AJ16,0)))`; AN11 `=IF(OR(AN4<>"OK",MAX(AG5:AG16)<=0),"",INDEX(A5:A16,MATCH(MAX(AG5:AG16),AG5:AG16,0)))`
- AN14 `=IF(AN4<>"OK","Check",IF(AND(AN13="OK",ROUND(AN5-AN6,2)=0),"OK","Check"))`
- Sales log D5 `=IF(AND(A5="",B5="",TRIM(C5)=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(B5="","Check: no amount",IF(NOT(ISNUMBER(B5)),"Check: amount is text",IF(YEAR(A5)<>'Year grid'!$B$1,"Other year (not in the grid)","OK"))))))`

Sample values (fictional; recalculated). Log 329 rows, 01/01/2026 to 11/25/2026: one Etsy row per open day (closed Mondays and 07/04/2026), a second Market row every Saturday, a $640.00 Holiday market row on 11/21/2026, a -$38.50 refund on 03/17/2026.
| Month | Total | Days with sales | Best day | Best day sales |
|---|---|---|---|---|
| January | $2,878.94 | 27 | 01/03/2026 | $331.36 |
| February | $2,662.24 | 24 | 02/07/2026 | $318.65 |
| March | $3,432.97 | 26 | 03/21/2026 | $426.38 |
| April | $3,960.34 | 26 | 04/11/2026 | $473.88 |
| May | $5,279.40 | 27 | 05/23/2026 | $640.67 |
| June | $4,234.09 | 25 | 06/27/2026 | $587.36 |
| July | $4,003.81 | 26 | 07/25/2026 | $536.49 |
| August | $4,550.29 | 26 | 08/22/2026 | $552.06 |
| September | $4,939.65 | 26 | 09/19/2026 | $546.97 |
| October | $6,292.00 | 27 | 10/10/2026 | $680.76 |
| November | $7,641.00 | 21 | 11/21/2026 | $1,650.50 |
| December | $0.00 | 0 | | |
AN4:AN14: OK, $49,874.73, $49,874.73, 281, $177.49, 11/21/2026, $1,650.50, November, $7,641.00, OK, OK. Grid examples: 01/01 $57, 01/03 $331, 02/02 "-" (Monday), 02/29 to 02/31 blank, 11/21 1,651 (darkest cell).

Edge cases tested (all matched the Python model)
- The sample shown as 2025 (every row Other year, grid all dashes, CHECK OK); year "2026" as text, blank, 2026.5, 1900 (Check, grid blank); 9999 (OK).
- 2028 with every day logged (02/29 filled, 366 days with sales), the same log shown as 2027, 2100 (not a leap year: 02/29 column blank).
- Empty log; refunds only (best day blank, best month blank); a tie for the best day (01/05 and 06/05 at $100.00: the first, 01/05, is returned).
- Bad rows: text date, a date with a time (Check, not in the grid: AN5 = AN6 = $4,518.56 for the good rows but CHECK Check because the log has Check rows), no amount, "10" as text, a note only, a 2025 date (Other year); a row on row 5005 (H8 1).
- Other years mixed in (12/31/2025 and 01/01/2027 counted as Other year, not in the grid); a full log of 5,000 rows; 200 random logs (years 1990 to 2040, 0 to 600 rows including dates 10 days outside the year, refunds and text dates).

Pitfalls the post must state
- The grid shows whole dollars; totals keep cents. A dash is a day with $0.00 or nothing logged; a blank is a date that does not exist.
- The color scale was set as a rule in the file and seen only in a LibreOffice PDF export; give the Excel and Sheets color scale steps from the help pages.
- A sale with a time of day in its date is not matched by any cell; the Sales log flags it and CHECK turns red. Sample sales are fictional.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- CHECK is red with dates that look right: a date has a time (Check: date has a time) or is text.
- The whole grid is blank: B1 is text, blank, not a whole number or before 1901.
- A month looks empty: the log rows are from another year (Other year).

Hero: `render(short_sheet_copy(t+"tidy-tabs-daily-sales-year-heatmap.xlsx","Year grid",16,"heat"),"Year grid",hide_cols=["AK","AL","AM","AN"],name="heat")` (the full 12 x 31 grid with the color scale, month totals and best days; legible). Note: pass the hidden columns as a list, since `hide_cols` iterates over its items and "AKALAMAN" as a string would hide single letters.

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): SUMIFS as in post 3; DATE https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349 and https://support.google.com/docs/answer/3092969; DAY https://support.microsoft.com/en-us/office/day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101 and https://support.google.com/docs/answer/3093040; MAX https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098 and https://support.google.com/docs/answer/3094013; conditional formatting (color scale steps) https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-in-excel-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f and https://support.google.com/docs/answer/78413; MATCH and INDEX as in post 2.

---
## 5. sales-change-vs-last-year-price-effect-and-volume-effect-by-product-sumifs-excel-google-sheets
Post: posts/ready/2026-11-26-team-05-sales-change-vs-last-year-price-effect-and-volume-effect-by-product-sumifs-excel-google-sheets.md
Title (hooks final): How to See If Sales Grew From Price or Volume in Excel and Google Sheets
Template: templates/tidy-tabs-sales-change-price-volume.xlsx (tabs: Compare, Sales log, How to use)

Layout
- Compare: inputs B1 This period from (01/01/2026), B2 This period to (10/31/2026), B3 Last period from (01/01/2025), B4 Last period to (10/31/2025); C1:C4 notes. Header row 6, rows 7:56: A Product (input), B Last units, C Last revenue, D Last avg price, E This units, F This revenue, G This avg price, H Revenue change, I Volume effect, J Price effect, K New or dropped, L Status. Summary N7:O21: Periods (O7), Last period revenue (O8), This period revenue (O9), Revenue change (O10), Revenue change % (O11), Units last (O12), Units this (O13), Volume effect (O14), Price effect (O15), New and dropped (O16), Change not explained (O17), Log revenue not in the list (O18), Rows showing Check (O19), Sales log CHECK (O20), CHECK (O21). Row 6 repeated.
- Sales log: A1:A2 notes. Header row 4, rows 5:3004: A Date, B Product, C Units, D Revenue (inputs), E Status. Summary H4:I9: Rows OK, Revenue of OK rows, Outside both periods, Rows showing Check, Typed below row 3004, CHECK.
- Builder decisions: price effect is the remainder H-I so the parts add up to the cent; New product / Dropped rows put their whole change in K; no mix line.

Key formulas (exact text in the file)
- B7 `=IF(TRIM(A7)="","",SUMIFS('Sales log'!$C$5:$C$3004,'Sales log'!$B$5:$B$3004,A7,'Sales log'!$E$5:$E$3004,"OK",'Sales log'!$A$5:$A$3004,">="&$B$3,'Sales log'!$A$5:$A$3004,"<="&$B$4))`; C7 the same with column D; E7 and F7 the same with $B$1 and $B$2
- D7 `=IF(OR(L7="",LEFT(L7,5)="Check",N(B7)<=0),"",C7/B7)`; G7 `=IF(OR(L7="",LEFT(L7,5)="Check",N(E7)<=0),"",F7/E7)`
- H7 `=IF(OR(L7="",LEFT(L7,5)="Check"),"",ROUND(F7-C7,2))`; I7 `=IF(L7<>"OK","",ROUND((E7-B7)*D7,2))`; J7 `=IF(L7<>"OK","",ROUND(H7-I7,2))`; K7 `=IF(L7="New product",H7,IF(L7="Dropped",H7,IF(OR(L7="OK",L7="No sales in either period"),0,"")))`
- L7 `=IF(TRIM(A7)="","",IF($O$7<>"OK","Check: period dates (B1:B4)",IF(COUNTIF($A$7:$A$56,A7)>1,"Check: product listed twice",IF(OR(B7<0,E7<0),"Check: net units below 0",IF(AND(B7=0,C7<>0),"Check: last revenue with 0 units",IF(AND(E7=0,F7<>0),"Check: this revenue with 0 units",IF(AND(B7=0,E7=0),"No sales in either period",IF(B7=0,"New product",IF(E7=0,"Dropped","OK")))))))))`
- O7 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B4)),IF(AND(B1<=B2,B3<=B4,OR(B3>B2,B1>B4)),"OK","Check"),"Check")`; O8 `=SUMIFS(C7:C56,L7:L56,"<>Check*")`; O17 `=ROUND(O10-O14-O15-O16,2)`; O21 `=IF(AND(O7="OK",O17=0,O18=0,O19=0,O20="OK"),"OK","Check")`
- Sales log E5 `=IF(AND(A5="",TRIM(B5)="",C5="",D5=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(TRIM(B5)="","Check: no product",IF(COUNTIF(Compare!$A$7:$A$56,B5)=0,"Check: product not on the Compare tab",IF(OR(NOT(ISNUMBER(C5)),NOT(ISNUMBER(D5))),"Check: units or revenue blank or text",IF(OR(AND(A5>=Compare!$B$1,A5<=Compare!$B$2),AND(A5>=Compare!$B$3,A5<=Compare!$B$4)),"OK","Outside both periods")))))))`

Sample values (fictional; recalculated). Log 106 rows: one row per product per month (15th) for January to October of 2025 and 2026, some lines with a 10% or 20% discount, a -2 unit / -$48.00 Lavender return on 04/02/2026, and 5 rows on 11/15/2025 (Outside both periods).
| Product | Last units | Last revenue | Last avg | This units | This revenue | This avg | Change | Volume | Price | New/dropped | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Cedar candle 8 oz | 410 | $9,009.00 | $21.97 | 388 | $9,302.40 | $23.98 | $293.40 | -$483.41 | $776.81 | $0.00 | OK |
| Lavender candle 8 oz | 352 | $7,737.40 | $21.98 | 369 | $8,836.80 | $23.95 | $1,099.40 | $373.68 | $725.72 | $0.00 | OK |
| Pumpkin candle 4 oz | 260 | $3,637.20 | $13.99 | 301 | $4,211.20 | $13.99 | $574.00 | $573.56 | $0.44 | $0.00 | OK |
| Wax melts 6-pack | 180 | $1,700.50 | $9.45 | 150 | $1,494.00 | $9.96 | -$206.50 | -$283.42 | $76.92 | $0.00 | OK |
| Holiday gift set | 0 | $0.00 | | 64 | $3,038.40 | $47.48 | $3,038.40 | | | $3,038.40 | New product |
| Oat soap bar | 95 | $755.20 | $7.95 | 0 | $0.00 | | -$755.20 | | | -$755.20 | Dropped |
O7:O21: OK, $22,839.30, $26,882.80, $4,043.50, 17.7%, 1,297, 1,272, **$180.41**, **$1,579.89**, **$2,283.20**, $0.00, $0.00, 0, OK, OK. Sales log: 101 OK, $49,722.10, 5 outside, 0 Check.

Edge cases tested (all matched the Python model)
- Full calendar years; periods swapped (all signs flip); overlapping periods, a reversed period, a text date (O7 Check, every row Check: period dates); one-day periods.
- "cedar candle 8 oz" added as a second row (COUNTIF is not case-sensitive: both rows Check: product listed twice, CHECK Check); only 4 products listed (the 21 log rows of the two unlisted products show Check: product not on the Compare tab, O18 stays $0.00, CHECK Check through the Sales log CHECK); an extra product with no sales (No sales in either period, 0 in K); no products at all.
- Bad log rows: text date, date with a time, product of spaces, a trailing space (no match, Check), "1" as text, blank revenue (each Check); "CEDAR CANDLE 8 OZ" matched (not case-sensitive).
- A 0-unit line with $15.00 revenue last period (Check: last revenue with 0 units); net -3 units (Check: net units below 0); returns that cancel a sale to 0 units and $0.00 (Dropped); price only (100 units both years, +$200.00 all price), volume only (100 to 130 units at $22.00, +$660.00 all volume); odd prices (3 units for $10.00 vs 7 for $23.33).
- 50 products with 3,000 rows; 200 random sets (1 to 50 products, 0 to 400 rows with returns, discounts and unlisted products, periods of 1 to 301 days from 2009 to 2035). Separately, in all 3,162 OK rows of the 220 recalculated workbooks, J7 differed from (G7-D7)*E7 (computed from the LibreOffice values) by at most $0.005, never $0.01 or more.

Pitfalls the post must state
- Average price is revenue per unit; discounts and returns move it. A price effect is not a list price change.
- New and dropped products have no old or new price, so their whole change is its own line. No mix line; the per-product split is shown instead.
- Product names must match the Compare tab exactly (case does not matter, spaces do). Sample sales are fictional; no pricing advice.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- CHECK red and Sales log rows say Check: product not on the Compare tab: the product is missing from column A, or the log name has a trailing space. (O18 is a safeguard; it stayed $0.00 in every test because such rows are already Check.)
- Every row says Check: period dates: the periods overlap, one is reversed, or a date is text.
- A product shows Check: revenue with 0 units: a $0-unit line with revenue (fix the units).

Hero: `render(short_sheet_copy(t+"tidy-tabs-sales-change-price-volume.xlsx","Compare",21,"pv"),"Compare",name="pv")` (periods, the six products with volume and price effects, New product and Dropped rows, and the summary with the three green parts; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): SUMIFS as in post 3; ROUND as in post 1; COUNTIF as in post 1.

---
Not built or tested (so not for the posts): anything in Excel or Google Sheets; how Excel or Sheets draw the color scale; typing by hand in the GUI (all inputs were written with openpyxl and recalculated headless, including the date-times); stacked or percent-off tiers, coupons, shipping and sales tax; overnight hours, split shifts within a day, time zones, holiday rules or holiday lists; posting hours to a website or Google Business Profile; a bank's or card issuer's own balance or interest method, compounding, fees or minimum-balance rules; sales by week or weekday heatmaps, multiple years on one grid; price-volume-mix with a mix line, cost or margin effects, forecasting or pricing advice.
