# Handoff 2026-11-25: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-25.md. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`black_friday_vs_last_year` with `BF_LOG_LAST`/`BF_ROWS_LAST`, `petty_cash_log` with `PETTY_LAST`/`PETTY_CATS_LAST`, `purchase_order_tracker` with `PO_LAST`/`PO_REC_LAST`/`PO_SUP_LAST`, `stockout_lost_sales` with `STOCK_LOG_LAST`/`STOCK_SKU_LAST`, `fiscal_calendar_445` with `FISCAL_YEARS_LAST`/`FISCAL_DATES_LAST`, helper `_thanksgiving`; registered in `__main__` after `find_transposed_digits()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-25 템플릿 제작/검증 (LibreOffice Calc 24.2.7.2)"). Heroes are not made (designer; trial renders in scratch only, calls below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 1,063 recalculated workbooks, 5,999,269 cell comparisons, 0 mismatches against independent Python models in the final runs, 0 error values in any recalculated workbook. A mutation run (one wrong formula per template) produced mismatches in all five (2,519 / 20 / 618 / 103 / 5,233), so the harness catches wrong formulas. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values, Err: or ### in the PDF text (the only "#DIV/0!" is the How to use sentence of the stockout file).

Repo note: `python3 scripts/build-templates.py` (whole script, /usr/bin/python3 3.13 with openpyxl 3.1.5) ran without errors, but it rewrites every older template byte-wise (zip timestamps), so all older .xlsx were restored with `git checkout`; the five new files were then rebuilt by importing the module and calling only the five functions. Importing the scripts also rewrote the tracked scripts/__pycache__/*.pyc; those were restored too. `git status` after the builder: scripts/build-templates.py (873 lines added, 0 removed), tasks/briefs.md, this handoff, five new .xlsx (tasks/topics-seed.md, tasks/team-topics-2026-11-25.md and tasks/hooks-2026-11-25.md were changed or added by other stages, not by the builder).

Conventions in all five: How to use tab last; US Letter, fit to one page wide, header row repeated on the data tabs; data tabs landscape; yellow = inputs, gray = formulas; red = any Status starting "Check" and any CHECK / settings cell that is not OK; light orange = warnings that are not errors (No entries..., Late, Out, 53-week year, No days logged); light green = highlights (Black Friday row, Replenish rows, the replenish amount, Received lines, the lost-sales cells). Functions used in the five files: AND, COUNT, COUNTA, COUNTIF, COUNTIFS, DATE, EDATE, IF, INDEX, INT, ISNUMBER, LEFT, MATCH, MAX, MID, MIN, MOD, MONTH, N, NOT, OR, ROUND, ROW, SUM, SUMIFS, SUMPRODUCT, TEXT, TRIM, VALUE, WEEKDAY, YEAR. No TODAY, FILTER, UNIQUE, SORT, LET, LAMBDA, REDUCE, XLOOKUP, TEXTSPLIT, MINIFS/MAXIFS, SEQUENCE; 0 `_xlfn` in the five saved files and 0 after the LibreOffice round trip. Help links were not checked by the builder (the researcher's list is in the topics file).

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-25-team-01-compare-black-friday-weekend-sales-to-last-year-same-sale-day-thanksgiving-weekday-sumifs-excel-google-sheets | tidy-tabs-black-friday-vs-last-year.xlsx | 360 workbooks, 1,501,037 comparisons |
| 2 | 2026-11-25-team-02-petty-cash-log-with-running-balance-and-replenishment-by-category-sumifs-countif-excel-google-sheets | tidy-tabs-petty-cash-log.xlsx | 217 workbooks, 338,737 comparisons |
| 3 | 2026-11-25-team-03-purchase-order-tracker-partial-deliveries-open-quantity-supplier-on-time-rate-sumifs-sumproduct-excel-google-sheets | tidy-tabs-purchase-order-tracker.xlsx | 215 workbooks, 871,180 comparisons |
| 4 | 2026-11-25-team-04-stockout-days-and-estimated-lost-sales-from-daily-inventory-log-countifs-sumifs-excel-google-sheets | tidy-tabs-stockout-lost-sales.xlsx | 207 workbooks, 2,640,699 comparisons |
| 5 | 2026-11-25-team-05-retail-4-4-5-fiscal-calendar-week-period-quarter-from-date-int-match-excel-google-sheets | tidy-tabs-4-4-5-fiscal-calendar.xlsx | 64 workbooks, 647,616 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets." For the Black Friday post add: "Thanksgiving dates were checked against Python's datetime for every year from 1990 to 2045." For the PO post add: "The COUNTIFS(range,\"<>\") counting problem was reproduced in LibreOffice Calc only."

Hooks check (tasks/hooks-2026-11-25.md, every "to confirm by builder" number, checked against the built templates in LibreOffice 24.2.7.2):
- 1 chosen: confirmed (Thanksgiving formula is B3 exactly, inside a settings guard; Black Friday 11/28/2025 $705.21 vs 11/29/2024 $579.27, +21.7%). Opening 2: confirmed (2024's Black Friday minus 364 = 12/01/2023, $180.24 instead of $222.00, template trap block with B1 = 2024; EDATE(Black Friday,-12) was a Wednesday or Thursday for every year 2000 to 2040 in Python and in every template year tested). Opening 3: confirmed with one wording fix: the template shows the day's sales as $0.00 next to "No entries this year" and leaves Change $ and Change % blank; the -100% is what E/F-1 would give without the flag (not shown by the template). Write "shows No entries instead of a -100% drop", not "instead of $0.00".
- 2 chosen: confirmed ($158.37 spent, expected $41.63, $41.37 counted, short $0.26, refill $158.63). Over/short in the file is H11 `=IF(H10="","",ROUND(H10-H9,2))` = counted minus (fund - spent). Opening 2: confirmed ($199.74, Replenish H13). Opening 3: the running balance is H5 `=ROUND(SUM($E$5:E5)-SUM($D$5:D5),2)` (core as in the hook); the **cycle formula differs**: the file counts the "Added as" column, I5 core `=COUNTIF($F$5:F5,"Replenish")`, not column B. Postage $82.05, Supplies $61.32, Travel $15.00 confirmed.
- 3 chosen: confirmed (PO-1002 tin lids 120 of 240, Partial, Late 15 days, $48.00 open as of 11/20/2026). Opening 2: confirmed (Wick Co fill rate 90%, Suppliers I5 = H5/G5 with H = SUMIFS of "Received up to ordered" and G = SUMIFS of ordered). Opening 3: confirmed (Soy Wax Direct uncapped 125% in Suppliers L7, capped 100%, the line "Over by 1").
- 4 chosen: confirmed (CED-8 out 10 of 28 days, 1.278 a day in stock, 13 lost units, $312.00 at $24.00). Day status in the file: H6 `=IF(F6="","",IF(F6<=0,"Out","In stock"))` with F6 = on hand + received. Opening 2: confirmed (0.821 over 28 days, 1.278 on 18 in-stock days); the shortcut's 8 lost units is now also a LibreOffice result (Summary L6), no longer Python only. Opening 3: confirmed (cedar out 10 days from 11/05/2026 to 11/14/2026, pumpkin out 1 day, 11/22/2026).
- 5 chosen: confirmed (11/27/2026 FY2026 W43 P10 Q4). The file's period formula is the general one, E5 core `=MATCH(week,first_weeks,1)` with a second first-week column for 53-week years, **not** `MATCH(MIN(week,52),...)`; with week 53 in P12 both give the same period (checked for every day FY2023 to FY2027), but quote the file's version. Opening 2: confirmed (01/30/2027 MONTH 1, FY2026 W52 P12). Opening 3: confirmed (FY2023 53 weeks from Calendar C6 `=(B7-B6)/7`; 02/03/2024 W53 P12). The builder did not check the NRF page.

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=..., name=...)` and `short_sheet_copy(src, sheet, last_row, name)`; every data tab fills gray formulas far down, so use short_sheet_copy). Trial renders at 1600x900 were made under /usr/bin/python3 into scratch only (/tmp/claude-0/-home-user-tidy-tabs/ecfe5f88-7e1f-5333-84c8-1933f5219d51/scratchpad/hero/); nothing was written to images/. All calls below rendered legibly.

---
## 1. compare-black-friday-weekend-sales-to-last-year-same-sale-day-thanksgiving-weekday-sumifs-excel-google-sheets
Title (topics): How to Compare Black Friday Sales to Last Year's Black Friday (Not the Same Date) in Excel and Google Sheets
Template: templates/tidy-tabs-black-friday-vs-last-year.xlsx (tabs: Compare, Sales log, How to use)

Layout
- Compare: A1 This year / B1 (input, 2025 in the sample), A2 Compare with / B2 (input, `=B1-1` by default, can be typed over), A3 Thanksgiving this year / B3, A4 Thanksgiving compare year / B4 (formulas); C1:C4 notes. Header row 6, rows 7:26 (20 rows): A Sale day (input, label only), B Days from Thanksgiving (input offset), C This year date, D Compare year date, E This year sales, F Compare year sales, G Change $, H Change %, I Status. Summary K6:L15: Settings (B1:B2) (L6), Days compared (L7), This year total (L8), Compare year total (L9), Change $ (L10), Change % (L11), Days with no entries (counted as $0.00) (L12), Rows showing Check (L13), Sales log CHECK (L14), CHECK (L15). Trap block K17:N22 "Why not date - 364 or EDATE (one year back)": header row 18 Method / Date / Day / Sales; row 19 Black Friday this year (B3+1), 20 Black Friday compare year (B4+1), 21 This year's Black Friday minus 364, 22 EDATE(this year's Black Friday,-12). Landscape, row 6 repeated. The Black Friday row is light green.
- Sales log: A1:A2 notes, header row 4, rows 5:2004: A Date, B Amount, C Channel (optional) (inputs), D Days from Thanksgiving, E Status. Summary G4:H8: Rows OK (H4), Total of OK rows (H5), Rows showing Check (H6), Typed below row 2004 (not read) (H7), CHECK (H8). Row 4 repeated.
- Deviations from the topic: the sample shows **2025 vs 2024** (B1 = 2025), because the fictional log covers 2023 to 2025 only; with B1 = 2026 every row says "No entries this year" (the note in C1 tells readers to type 2026 once this year's sales are in). Totals sit in the K:L summary, not in a totals row. Change $ and Change % stay blank on a "No entries" row (the topic only asked for the flag); the totals still count that day as $0.00 and L12 counts such days. Added: duplicate-offset check, settings check (years 1901 to 9999, whole, B2 not equal to B1), the trap block with TEXT(date,"dddd") day names, a "date has a time" check on the log (SUMIFS matches whole dates only).

Key formulas (exact text in the file)
- B3 `=IF($L$6<>"OK","",DATE(B1,11,1)+MOD(5-WEEKDAY(DATE(B1,11,1)),7)+21)`; B4 the same with B2
- C7 `=IF(OR($L$6<>"OK",NOT(ISNUMBER(B7))),"",IF(B7<>INT(B7),"",$B$3+B7))`; D7 `=IF(C7="","",$B$4+B7)`
- E7 `=IF(C7="","",SUMIFS('Sales log'!$B$5:$B$2004,'Sales log'!$A$5:$A$2004,C7))`; F7 the same with D7
- G7 `=IF(OR(C7="",LEFT(I7,2)="No"),"",ROUND(E7-F7,2))`; H7 `=IF(OR(C7="",LEFT(I7,2)="No"),"",IF(F7=0,"",E7/F7-1))`
- I7 `=IF(AND(TRIM(A7)="",B7=""),"",IF($L$6<>"OK","Check: settings (B1:B2)",IF(NOT(ISNUMBER(B7)),"Check: offset blank or text",IF(B7<>INT(B7),"Check: offset must be a whole number",IF(COUNTIF($B$7:$B$26,B7)>1,"Check: same offset twice",IF(AND(COUNTIF('Sales log'!$A$5:$A$2004,C7)=0,COUNTIF('Sales log'!$A$5:$A$2004,D7)=0),"No entries either year",IF(COUNTIF('Sales log'!$A$5:$A$2004,C7)=0,"No entries this year",IF(COUNTIF('Sales log'!$A$5:$A$2004,D7)=0,"No entries compare year","OK"))))))))`
- L6 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),IF(AND(B1>=1901,B1<=9999,B1=INT(B1),B2>=1901,B2<=9999,B2=INT(B2),B2<>B1),"OK","Check"),"Check")`; L8 `=SUM(E7:E26)`; L10 `=ROUND(L8-L9,2)`; L11 `=IF(L9=0,"",L8/L9-1)`; L12 `=COUNTIF(I7:I26,"No entries*")`; L15 `=IF(AND(L6="OK",L7>0,L13=0,L14="OK"),"OK","Check")`
- Trap block: L19 `=IF($L$6<>"OK","",$B$3+1)`; L21 `=IF(L19="","",L19-364)`; L22 `=IF(L19="","",EDATE(L19,-12))`; M19 `=IF(L19="","",TEXT(L19,"dddd"))`; N19 `=IF(L19="","",SUMIFS('Sales log'!$B$5:$B$2004,'Sales log'!$A$5:$A$2004,L19))`
- Sales log D5 `=IF(E5<>"OK","",A5-(DATE(YEAR(A5),11,1)+MOD(5-WEEKDAY(DATE(YEAR(A5),11,1)),7)+21))`; E5 `=IF(AND(A5="",B5="",TRIM(C5)=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(B5="","Check: no amount",IF(NOT(ISNUMBER(B5)),"Check: amount is text","OK")))))`

Sample values (fictional; recalculated). Log: one Etsy row per day 11/15 to 12/06 of 2023, 2024 and 2025 (amounts $180.xx to $300.xx, the researcher's seed), plus a second "Market booth" row on each Black Friday ($412.50 on 11/28/2025, $388.25 on 11/29/2024); 11/30/2025 left out. 67 rows, total $16,224.92. B1 2025, B2 2024: Thanksgiving 11/27/2025 and 11/28/2024.
| Sale day | Offset | 2025 date | 2024 date | 2025 sales | 2024 sales | Change $ | Change % | Status |
|---|---|---|---|---|---|---|---|---|
| Monday before Thanksgiving | -3 | 11/24/2025 | 11/25/2024 | $274.45 | $181.60 | $92.85 | 51.1% | OK |
| Tuesday | -2 | 11/25/2025 | 11/26/2024 | $183.96 | $211.91 | -$27.95 | -13.2% | OK |
| Wednesday | -1 | 11/26/2025 | 11/27/2024 | $259.84 | $239.53 | $20.31 | 8.5% | OK |
| Thanksgiving | 0 | 11/27/2025 | 11/28/2024 | $246.15 | $230.71 | $15.44 | 6.7% | OK |
| Black Friday | 1 | 11/28/2025 | 11/29/2024 | $705.21 | $579.27 | $125.94 | 21.7% | OK |
| Small Business Saturday | 2 | 11/29/2025 | 11/30/2024 | $277.73 | $219.67 | $58.06 | 26.4% | OK |
| Sunday | 3 | 11/30/2025 | 12/01/2024 | $0.00 | $245.93 | (blank) | (blank) | No entries this year |
| Cyber Monday | 4 | 12/01/2025 | 12/02/2024 | $187.71 | $223.35 | -$35.64 | -16.0% | OK |
L6:L15: OK, 8, $2,135.05, $2,131.97, $3.08, 0.1%, 1, 0, OK, OK. Trap block: 11/28/2025 Friday $705.21; 11/29/2024 Friday $579.27; minus 364 = 11/29/2024 Friday $579.27 (right this year); EDATE = 11/28/2024 Thursday $230.71 (Thanksgiving 2024).
- B1 = 2024 (vs 2023): Thanksgiving 11/28/2024 and 11/23/2023; Black Friday $579.27 vs $222.00 (+$357.27, 160.9%); window $2,131.97 vs $1,773.85 (+$358.12, 20.2%). Trap block: minus 364 = **12/01/2023** (a Friday, one week after Black Friday 2023) pulling **$180.24** instead of $222.00; EDATE = 11/29/2023, a **Wednesday**, $272.77.
- B1 = 2026: B3 11/26/2026, B4 11/27/2025, every row "No entries this year", L12 8, CHECK OK.

Edge cases tested (all matched the Python model)
- B1 every year 2000 to 2040 (41 workbooks, log one row per day 11/18 to 12/05 of every year 1990 to 2045, so the Sales log D column checked Thanksgiving for all 56 years against Python's datetime): all matched. Thanksgiving on 11/22 (2001, 2007, 2012, 2018, 2029, 2035, 2040) and on 11/28 (2002, 2013, 2019, 2024, 2030); Cyber Monday in December: 12/02/2024, 12/02/2019, 12/02/2030.
- Date - 364 (Python, Black Friday this year minus 364 vs last year's Black Friday, 2000 to 2040): wrong in 2002, 2008, 2013, 2019, 2024, 2030, 2036, right in the other years; the template trap row showed the same for every B1 tested. EDATE(Black Friday,-12): Wednesday or Thursday for every year 2000 to 2040 (Python), never a Friday.
- Offsets widened to -7 .. +11 (19 rows), for 2025 and 2024: 2025 vs 2024 gives 19 days, $4,282.38 vs $4,025.86 (+$256.52, 6.4%) with 4 "No entries" days (11/30/2025 and the days past the sample log's 12/06 end). Several entries per day; refunds as negative amounts (-$45.00 and -$12.50 on Black Friday 2025 make it $647.71; $19.99, $20.01 and $0.00 more on Black Friday 2024 make it $619.27; a -$187.71 refund makes Cyber Monday 2025 $0.00 with entries: Status OK, change -100.0%).
- Missing day in either year ("No entries this year" / "compare year" / "either year"); a log holding only 2025 (every row "No entries compare year").
- Sales log checks: date typed as text "11/28/2025" (Check: date blank or text, not counted, COUNTIF does not see it), date with a time (Check: date has a time, not matched), blank amount, amount "40.00" as text, a channel with no date, an amount with no date; a row typed on row 2005 (H7 1, CHECK Check).
- Settings: year "2025" as text, 2025.5, B2 = B1, B1 blank, 1900 (all L6 Check, every row Check: settings). Offsets: "two" (Check: offset blank or text), 1.5 (whole number), 1 twice (same offset twice, both rows), a label with no offset, an offset with no label (computed, OK).
- A full log of 2,000 rows, and 300 random logs (1 to 400 rows over 1 to 4 random years 2005 to 2035, random compare years, random windows).

Pitfalls the post must state
- The Thanksgiving rule is the US one (fourth Thursday of November); say once. "Small Business Saturday" is just a label.
- Differences are sales, not profit (no ad spend or discounts). The sample sales are fictional.
- Dates must be real dates without times; text dates and date-times are flagged and not counted.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Comparing 11/28 with 11/28: Black Friday 2025 ($705.21) against Thanksgiving 2024 ($230.71), which is what EDATE(-12) also gives.
- Date - 364 picks the wrong week: 2024's Black Friday minus 364 = 12/01/2023 ($180.24, not $222.00).
- A day shows $0.00: Status says No entries this year (11/30/2025); log that day or ignore it.
- A sale is missing from the total: its date is text or has a time (Sales log Status).

Hero: `render(short_sheet_copy(t+"tidy-tabs-black-friday-vs-last-year.xlsx","Compare",22,"bf"),"Compare",name="bf")` (settings, the 8 sale days with the green Black Friday row and the orange No entries flag, the summary and the trap block; legible).

---
## 2. petty-cash-log-with-running-balance-and-replenishment-by-category-sumifs-countif-excel-google-sheets
Title (topics): How to Keep a Petty Cash Log and Work Out the Replenishment in Excel and Google Sheets
Template: templates/tidy-tabs-petty-cash-log.xlsx (tabs: Log, Replenish, How to use)

Layout
- Log: A1 Fund amount / B1 (input, $200.00), A2 Cash counted now / B2 (input, $41.37), C1:C2 notes. Header row 4, rows 5:504: A Date, B Paid to / note, C Category (dropdown from Replenish!A5:A14), D Paid out, E Added, F Added as (dropdown Start / Replenish / Fund change), G Receipt? (Y/N) (dropdown) (inputs); H Balance, I Cycle, J Status. Mini summary L4:M7 (links to the Replenish tab): Cash expected in the box (M4), Over (+) or short (-) (M5), Replenish amount (M6), CHECK (M7). Replenish rows are light green. Row 4 repeated.
- Replenish: A1:A2 notes, header row 4, rows 5:14: A Category (input; sample Postage, Supplies, Meals, Travel, Other, Over/short), B Spent this cycle, C Payouts this cycle, D Spent in all cycles, E Status. Summary G4:H20: Fund amount (Log B1) (H4), Current cycle (Replenish rows so far) (H5), This cycle started on (H6), Spent this cycle (H7), Spent outside the listed categories (H8), Cash expected in the box (fund - spent) (H9), Cash counted (Log B2) (H10), Over (+) or short (-) (H11), Replenish amount (fund - counted) (H12, green), Box after a refill of the receipts only (H13), Payouts without a receipt this cycle (H14), Log balance (added - paid out) (H15), Log balance minus cash expected (H16), Log rows showing Check (H17), Category rows showing Check (H18), Typed below row 504 on the Log (not read) (H19), CHECK (H20).
- Builder decisions (the topic left them open): a separate "Added as" column (F) instead of reading "Replenish" from the note; the cycle counts Replenish rows in F. "Fund change" is the label for cash that grows the fund (also change B1). A sixth category **Over/short** is in the list: when refilling, log the short as an Over/short payout (or a negative payout for cash over; negative payouts are allowed only in Over/short) so the log balance stays equal to the box. CHECK needs "Log balance minus cash expected" = $0.00 (H16), which catches a Fund change without updating B1 and an earlier short that was never logged. Payouts without a receipt (H14) leave out Over/short rows.

Key formulas (exact text in the file)
- Log H5 `=IF(AND(D5="",E5=""),"",ROUND(SUM($E$5:E5)-SUM($D$5:D5),2))`; I5 `=IF(AND(D5="",E5=""),"",COUNTIF($F$5:F5,"Replenish"))`
- Log J5 `=IF(AND(A5="",TRIM(B5)="",TRIM(C5)="",D5="",E5="",TRIM(F5)="",TRIM(G5)=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(AND(D5="",E5=""),"Check: no amount",IF(AND(D5<>"",E5<>""),"Check: paid out and added on one row",IF(OR(AND(D5<>"",NOT(ISNUMBER(D5))),AND(E5<>"",NOT(ISNUMBER(E5)))),"Check: amount is text",IF(OR(ROUND(N(D5),2)<>N(D5),ROUND(N(E5),2)<>N(E5)),"Check: more than 2 decimals",IF(D5<>"",IF(AND(D5<=0,C5<>"Over/short"),"Check: paid out must be above 0",IF(TRIM(C5)="","Check: no category",IF(COUNTIF(Replenish!$A$5:$A$14,C5)=0,"Check: category not on the Replenish tab",IF(TRIM(F5)<>"","Check: Added as is only for cash added",IF(AND(G5<>"",G5<>"Y",G5<>"N"),"Check: Receipt? must be Y or N",IF(H5<0,"Check: balance below 0","OK")))))),IF(AND(F5<>"Start",F5<>"Replenish",F5<>"Fund change"),"Check: Added as must be Start, Replenish or Fund change",IF(TRIM(C5)<>"","Check: category is only for payouts",IF(AND(E5<=0,F5<>"Fund change"),"Check: added must be above 0",IF(H5<0,"Check: balance below 0","OK")))))))))))`
- Replenish B5 `=IF(TRIM(A5)="","",SUMIFS(Log!$D$5:$D$504,Log!$C$5:$C$504,A5,Log!$I$5:$I$504,$H$5))`; C5 `=IF(TRIM(A5)="","",COUNTIFS(Log!$C$5:$C$504,A5,Log!$I$5:$I$504,$H$5))`; D5 `=IF(TRIM(A5)="","",SUMIFS(Log!$D$5:$D$504,Log!$C$5:$C$504,A5))`
- H5 `=COUNTIF(Log!$F$5:$F$504,"Replenish")`; H6 `=IF(H5=0,IF(COUNT(Log!$A$5:$A$504)=0,"",MIN(Log!$A$5:$A$504)),INDEX(Log!$A$5:$A$504,SUMPRODUCT(MAX((Log!$F$5:$F$504="Replenish")*(ROW(Log!$F$5:$F$504)-4)))))`; H7 `=SUMIFS(Log!$D$5:$D$504,Log!$I$5:$I$504,H5)`; H8 `=ROUND(H7-SUM(B5:B14),2)`
- H9 `=ROUND(H4-H7,2)`; H10 `=IF(ISNUMBER(Log!B2),Log!B2,"")`; H11 `=IF(H10="","",ROUND(H10-H9,2))`; H12 `=IF(H10="","",ROUND(H4-H10,2))`; H13 `=IF(H10="","",ROUND(H10+H7,2))`
- H14 `=COUNTIFS(Log!$I$5:$I$504,H5,Log!$D$5:$D$504,"<>",Log!$G$5:$G$504,"<>Y",Log!$C$5:$C$504,"<>Over/short")`; H15 `=ROUND(SUM(Log!$E$5:$E$504)-SUM(Log!$D$5:$D$504),2)`; H16 `=ROUND(H15-H9,2)`
- H20 `=IF(AND(ISNUMBER(Log!B1),ISNUMBER(Log!B2),H8=0,H16=0,H17=0,H18=0,H19=0),IF(AND(Log!B1>0,Log!B2>=0,Log!B2<=Log!B1),"OK","Check: fund or count (Log B1:B2)"),"Check")`

Sample values (fictional; recalculated). Fund $200.00, counted $41.37. 13 rows 11/02/2026 to 11/24/2026:
| Date | Paid to / note | Category | Paid out | Added | Added as | Receipt? | Balance | Cycle |
|---|---|---|---|---|---|---|---|---|
| 11/02/2026 | Start fund | | | $200.00 | Start | | $200.00 | 0 |
| 11/03/2026 | Post office | Postage | $18.40 | | | Y | $181.60 | 0 |
| 11/05/2026 | Hardware store | Supplies | $23.17 | | | Y | $158.43 | 0 |
| 11/09/2026 | Coffee for market crew | Meals | $14.60 | | | N | $143.83 | 0 |
| 11/12/2026 | Post office | Postage | $9.85 | | | Y | $133.98 | 0 |
| 11/14/2026 | Parking | Travel | $12.00 | | | Y | $121.98 | 0 |
| 11/16/2026 | Refill from bank | | | $78.02 | Replenish | | $200.00 | 1 |
| 11/18/2026 | Craft store | Supplies | $31.49 | | | Y | $168.51 | 1 |
| 11/19/2026 | Post office | Postage | $27.30 | | | Y | $141.21 | 1 |
| 11/20/2026 | Ice for cooler | Supplies | $6.99 | | | N | $134.22 | 1 |
| 11/21/2026 | Parking | Travel | $15.00 | | | Y | $119.22 | 1 |
| 11/23/2026 | Tape and labels | Supplies | $22.84 | | | Y | $96.38 | 1 |
| 11/24/2026 | Post office | Postage | $54.75 | | | Y | $41.63 | 1 |
Replenish (current cycle 1, started 11/16/2026): Postage $82.05 (2 payouts; all cycles $110.30), Supplies $61.32 (3; $84.49), Meals $0.00 (0; $14.60), Travel $15.00 (1; $27.00), Other $0.00, Over/short $0.00. Spent this cycle **$158.37**, outside the list $0.00, cash expected **$41.63**, counted $41.37, over/short **-$0.26** (short), **replenish $158.63** (not $158.37), box after a refill of receipts only **$199.74**, payouts without a receipt 1 (ice), log balance $41.63, difference $0.00, CHECK OK.

Edge cases tested (all matched the Python model)
- No refill yet (log to 11/14, counted $121.98): cycle 0, started 11/02/2026, spent $78.02, replenish $78.02, CHECK OK.
- Two refills (an Over/short payout of $0.26, a Replenish of $158.63 on 11/25, then $11.10 + $8.25): cycle 2, spent $19.35, expected $180.65, counted $180.65, over/short $0.00, replenish $19.35, CHECK OK. Cash over logged as Over/short -$0.40 with counted $42.03: over/short $0.00, replenish $157.97, OK.
- Refill typed as "Refill" in Added as: Check, not counted as a refill (cycle 0, spent $236.39, expected -$36.39, log difference $78.02, CHECK Check). Typed "replenish" (lower case): counted as a refill (LibreOffice text compare is not case-sensitive), OK.
- Payout with no category / with "Snacks" (not listed): row Check, Spent outside the listed categories $27.30, CHECK Check.
- Fund change +$50.00 with B1 updated to $250.00 (counted $91.37): expected $91.63, short $0.26, replenish $158.63, OK. The same without updating B1: log difference $50.00, CHECK Check.
- Counted blank: H10:H13 blank, CHECK Check. Counted "41.37" as text: same. Counted $250.00 (above the fund): replenish -$50.00, CHECK "Check: fund or count (Log B1:B2)". Counted exactly $41.63: over/short $0.00.
- Row checks, each its own message: 1.005 (more than 2 decimals), paid and added on one row, -$2.00 payout outside Over/short, "2.00" as text, added with no Added as, added with a category, Receipt? "maybe", a text date, no amount, a payout that takes the balance below $0.00, Added as on a payout, added $0.00, a negative Fund change; a row on row 600 (H19); the same category twice ("postage" next to "Postage", both Check).
- A full log of 500 rows (refills whenever the box fell below $60.00), and 200 random logs (1 to 120 rows, random funds $100.00 to $500.00, refills off by -$0.25 to +$0.50, mislabeled refills, Over/short rows, bad rows).
- Rounding note (not a template change): in a first random run, 49 of 217 workbooks held a 3-decimal amount (1.005, flagged Check), and the running balance after it was an exact half cent; LibreOffice rounded those half cents down in that run (242 cells, each $0.01 from a decimal half-up model) but up in the hand-typed bad-rows case (40.625 -> $40.63). The final 200 random logs use 2-decimal amounts only. Do not claim anything about 3-decimal amounts except that the sheet flags them.

Pitfalls the post must state
- Categories are the reader's own labels, not tax categories; no word on what is deductible (link IRS Pub 583 at most, no judgment). Payees and amounts are fictional.
- Refill by fund minus count, not by the receipts total, and log the short (Over/short) when you refill.
- "Replenish" in Added as starts a new cycle; another word does not (the row shows Check).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The box never gets back to the fund: refilling by receipts ($158.37) leaves $199.74 (H13).
- CHECK says Check after raising the fund: B1 not changed after a Fund change row (H16 $50.00).
- A payout is not in the category totals: no category or a category not on the list (H8 shows the amount).
- The cycle did not restart after a refill: Added as typed as "Refill" (Check).

Hero: `render(short_sheet_copy(t+"tidy-tabs-petty-cash-log.xlsx","Replenish",20,"pc"),"Replenish",name="pc")` (categories and the whole summary, the green $158.63; very legible). Optional second card: `render(short_sheet_copy(t+"tidy-tabs-petty-cash-log.xlsx","Log",17,"pcl"),"Log",hide_cols="I",name="pcl")` (the 13 rows with the green Replenish row and the mini summary).

---
## 3. purchase-order-tracker-partial-deliveries-open-quantity-supplier-on-time-rate-sumifs-sumproduct-excel-google-sheets
Title (topics): How to Track Purchase Orders With Partial Deliveries in Excel and Google Sheets
Template: templates/tidy-tabs-purchase-order-tracker.xlsx (tabs: PO lines, Receipts, Suppliers, How to use)

Layout
- PO lines: A1 As-of date / B1 (input, 11/20/2026), C1 and A2 notes. Header row 4, rows 5:204: A PO #, B Supplier, C Item, D Qty ordered, E Unit cost, F Order date, G Expected date (inputs); H Received, I Open, J Receipt status, K Late as of B1, L Completed on, M On time?, N Received up to ordered, O Open value, P Status. Summary R4:S11: PO lines (S4), Open lines (S5), Late open lines as of B1 (S6), Open value (S7), Rows showing Check (S8), Receipts rows showing Check (S9), Typed below the last row (not read) (S10), CHECK (S11). Row 4 repeated.
- Receipts: A1:A2 notes, header row 4, rows 5:1004: A Date received, B PO #, C Item, D Qty received (inputs); E Status; F Date used (the date as a number for column L, 0 shown blank when the date is blank or text).
- Suppliers: A1:A2 notes, header row 4, rows 5:24: A Supplier (input), B PO lines, C Completed lines, D On time, E Late, F On-time rate, G Qty ordered, H Received up to ordered, I Fill rate, J Open lines, K Open value, L Fill rate if not capped, M Status. Summary O4:P7: PO lines with a supplier not listed here (P4), Supplier rows showing Check (P5), PO lines CHECK (P6), CHECK (P7).
- Deviations from the topic: the topic's "Status" column J is named Receipt status and the checks are in P (Status). K says "Late 1 day" (singular) and "Late 15 days" (the topic wrote "Late 1 days"). Completed on reads the helper Receipts!F (so a text date in the receipts cannot make SUMPRODUCT fail). Extra checks: qty 0 or not whole, expected before order, a receipt dated before the order date, PO # / item / supplier blank; Suppliers adds the uncapped fill rate (L) and open lines (J).

Key formulas (exact text in the file)
- H5 `=IF(OR(TRIM(A5)="",TRIM(C5)="",NOT(ISNUMBER(D5)),NOT(ISNUMBER(E5))),"",IF(D5<=0,"",SUMIFS(Receipts!$D$5:$D$1004,Receipts!$B$5:$B$1004,A5,Receipts!$C$5:$C$1004,C5)))`
- I5 `=IF(H5="","",MAX(0,D5-H5))`; J5 `=IF(H5="","",IF(H5=0,"Not received",IF(H5<D5,"Partial",IF(H5=D5,"Received","Over by "&(H5-D5)))))`
- K5 `=IF(OR(I5="",NOT(ISNUMBER($B$1)),NOT(ISNUMBER(G5))),"",IF(I5=0,"",IF($B$1>G5,"Late "&($B$1-G5)&IF($B$1-G5=1," day"," days"),"")))`
- L5 `=IF(I5="","",IF(I5>0,"",SUMPRODUCT(MAX((Receipts!$B$5:$B$1004=A5)*(Receipts!$C$5:$C$1004=C5)*Receipts!$F$5:$F$1004))))`; M5 `=IF(OR(L5="",NOT(ISNUMBER(G5))),"",IF(L5<=G5,"On time","Late"))`
- N5 `=IF(H5="","",MIN(H5,D5))`; O5 `=IF(I5="","",ROUND(I5*E5,2))`
- P5 `=IF(AND(TRIM(A5)="",TRIM(B5)="",TRIM(C5)="",D5="",E5="",F5="",G5=""),"",IF(OR(TRIM(A5)="",TRIM(C5)=""),"Check: PO # or item blank",IF(TRIM(B5)="","Check: no supplier",IF(OR(NOT(ISNUMBER(D5)),NOT(ISNUMBER(E5))),"Check: qty or cost blank or text",IF(OR(D5<=0,D5<>INT(D5)),"Check: qty must be a whole number above 0",IF(OR(NOT(ISNUMBER(F5)),NOT(ISNUMBER(G5))),"Check: order or expected date blank or text",IF(G5<F5,"Check: expected before order date",IF(COUNTIFS($A$5:$A$204,A5,$C$5:$C$204,C5)>1,"Check: same item twice on this PO",IF(SUMPRODUCT((Receipts!$B$5:$B$1004=A5)*(Receipts!$C$5:$C$1004=C5)*(Receipts!$F$5:$F$1004>0)*(Receipts!$F$5:$F$1004<F5))>0,"Check: a receipt is dated before the order date",IF(NOT(ISNUMBER($B$1)),"Check: as-of date (B1)","OK"))))))))))`
- Receipts E5 `=IF(AND(A5="",TRIM(B5)="",TRIM(C5)="",D5=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(OR(TRIM(B5)="",TRIM(C5)=""),"Check: PO # or item blank",IF(NOT(ISNUMBER(D5)),"Check: qty blank or text",IF(D5<=0,"Check: qty must be above 0",IF(COUNTIFS('PO lines'!$A$5:$A$204,B5,'PO lines'!$C$5:$C$204,C5)=0,"Check: no matching PO line","OK"))))))`; F5 `=IF(ISNUMBER(A5),A5,0)`
- Suppliers D5 `=IF(TRIM(A5)="","",COUNTIFS('PO lines'!$B$5:$B$204,A5,'PO lines'!$M$5:$M$204,"On time"))`; E5 the same with "Late"; C5 `=IF(TRIM(A5)="","",D5+E5)`; F5 `=IF(TRIM(A5)="","",IF(C5=0,"",D5/C5))`
- Suppliers G5 `=IF(TRIM(A5)="","",SUMIFS('PO lines'!$D$5:$D$204,'PO lines'!$B$5:$B$204,A5))`; H5 the same on column N; I5 `=IF(TRIM(A5)="","",IF(G5=0,"",H5/G5))`; L5 `=IF(TRIM(A5)="","",IF(G5=0,"",SUMIFS('PO lines'!$H$5:$H$204,'PO lines'!$B$5:$B$204,A5)/G5))`; K5 SUMIFS on column O; J5 `=IF(TRIM(A5)="","",COUNTIFS('PO lines'!$B$5:$B$204,A5,'PO lines'!$I$5:$I$204,">0"))`
- Suppliers M5 `=IF(TRIM(A5)="","",IF(COUNTIF($A$5:$A$24,A5)>1,"Check: same supplier twice",IF(B5=0,"No PO lines",IF(COUNTIFS('PO lines'!$B$5:$B$204,A5,'PO lines'!$P$5:$P$204,"Check*")>0,"Check: a PO line shows Check","OK"))))`; P4 `=SUMPRODUCT((TRIM('PO lines'!$B$5:$B$204)<>"")*(COUNTIF(A5:A24,'PO lines'!$B$5:$B$204)=0))`

Sample values (fictional; recalculated), as of 11/20/2026:
| PO # | Supplier | Item | Ordered | Unit cost | Expected | Received | Open | Receipt status | Late as of B1 | Completed on | On time? | Open value |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| PO-1001 | Wick Co | Cotton wick 6in | 500 | $0.08 | 10/30/2026 | 500 | 0 | Received | | 10/29/2026 | On time | $0.00 |
| PO-1001 | Wick Co | Wood wick 5in | 200 | $0.35 | 10/30/2026 | 200 (150 + 50) | 0 | Received | | 11/06/2026 | Late | $0.00 |
| PO-1002 | Jar Supply | 8 oz amber jar | 240 | $1.10 | 11/05/2026 | 240 | 0 | Received | | 11/04/2026 | On time | $0.00 |
| PO-1002 | Jar Supply | Black tin lid | 240 | $0.40 | 11/05/2026 | 120 | 120 | Partial | Late 15 days | | | $48.00 |
| PO-1003 | Soy Wax Direct | Soy wax 45 lb | 4 | $62.00 | 11/12/2026 | 5 | 0 | Over by 1 | | 11/11/2026 | On time | $0.00 |
| PO-1004 | Jar Supply | 4 oz amber jar | 120 | $0.85 | 11/24/2026 | 0 | 120 | Not received | | | | $102.00 |
| PO-1005 | Wick Co | Cotton wick 6in | 300 | $0.08 | 11/19/2026 | 200 | 100 | Partial | Late 1 day | | | $8.00 |
Order dates: 10/20, 10/20, 10/25, 10/25, 11/02, 11/10, 11/12/2026. Receipts: 10/29 PO-1001 cotton 500, 10/29 PO-1001 wood 150, 11/06 PO-1001 wood 50, 11/04 PO-1002 jars 240, 11/04 PO-1002 lids 120, 11/11 PO-1003 wax 5, 11/18 PO-1005 cotton 200. S4:S11: 7, 3, 2, $158.00, 0, 0, 0, OK.
| Supplier | Lines | Completed | On time | Late | On-time rate | Ordered | Received up to ordered | Fill rate | Open lines | Open value | Fill rate if not capped |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Wick Co | 3 | 2 | 1 | 1 | 50% | 1,000 | 900 | 90% | 1 | $8.00 | 90% |
| Jar Supply | 3 | 1 | 1 | 0 | 100% | 600 | 360 | 60% | 2 | $150.00 | 60% |
| Soy Wax Direct | 1 | 1 | 1 | 0 | 100% | 4 | 4 | 100% | 0 | $0.00 | 125% |

Trap reproduced (probe copy of the template, extra cells on the Suppliers tab): `=COUNTIFS('PO lines'!$B$5:$B$204,A5,'PO lines'!$M$5:$M$204,"<>")` counted 3 completed lines for Wick Co (correct 2) and 3 for Jar Supply (correct 1), because the open lines' formula blanks ("") counted as filled; on-time rate 33% instead of 50% and 33% instead of 100%. `SUMPRODUCT((B=A5)*(M<>""))` gave 2 and 1. The template counts On time + Late instead.

Edge cases tested (all matched the Python model)
- As-of 11/25/2026: Late 20 days, PO-1004 Late 1 day, PO-1005 Late 6 days, S6 3. As-of 10/30/2026 and 11/20/2026. As-of blank or "11/20/2026" as text: Late column blank, every line "Check: as-of date (B1)", CHECK Check.
- Receipt "Black tin lid " (trailing space): Receipts Check: no matching PO line, not added to the line. Trailing space on the PO line instead: the line shows 0 received (Not received, Late 15 days, open value $96.00) and its receipt shows Check.
- Same item twice on one PO (a second Black tin lid line on PO-1002): both lines Check: same item twice on this PO, and each line counts the same 120 (the reason it is a Check).
- Receipt dated 11/01/2026 for PO-1003 (ordered 11/02/2026): Check: a receipt is dated before the order date.
- Qty 0 line: Check: qty must be a whole number above 0, results blank. Qty 2.5, "5" as text, blank PO #, blank supplier, text order date, expected before order: each its own Check.
- Over-receipt on a late line (10 more wood wicks on 11/09): Over by 10, Completed on 11/09/2026, Late, uncapped fill rate 91%, capped 90%. 130 more lids on 11/19: Over by 10, completed 11/19/2026, Late; Jar Supply on-time rate 50%, fill rate 80% (uncapped 82%).
- "po-1005" / "COTTON WICK 6IN" in the receipts matched the line (LibreOffice SUMIFS and = are not case-sensitive): PO-1005 Received, On time.
- Receipts checks: text date, blank PO #, "1" as text, qty 0, PO-7777 (no matching line); Suppliers: "wick co" next to "Wick Co" (Check: same supplier twice), a supplier with no lines (No PO lines), a line whose supplier is not listed (P4 1); rows typed below the last row (S10).
- 200 lines with 1,000 receipts, and 200 random sets (1 to 40 lines, 0 to 80 receipts, random as-of dates, some trailing spaces, qty 0 / 2.5 lines).
- Rounding note (model, not template): three random sets had a 2.5 qty (flagged Check) whose open value was an exact half cent; LibreOffice gave $33.23 for 2.5 x $13.29 (= $33.225), which is the half-up result; the first Python model multiplied binary floats and was fixed.

Pitfalls the post must state
- No payment terms, invoices or three-way match (link invoice-due-date-payment-terms); no claim about how any real supplier ships. Sample suppliers, items and prices are fictional.
- PO # and item must match exactly (a trailing space breaks it; upper/lower case did not matter in LibreOffice).
- Receipts dated after the as-of date still count as received; the as-of date is for lateness only.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- On-time rate too low: COUNTIFS(...,"<>") counted formula blanks (Wick Co 33% instead of 50%).
- A delivery does not show on its line: trailing space in the item (Receipts Status: Check: no matching PO line).
- Fill rate above 100%: the uncapped version with an over-delivery (Soy Wax Direct 125%); MIN caps it at 100%.
- Open quantity negative: without MAX(0,...) the wax line showed -1 and the open value total dropped from $158.00 to $96.00 (mutation run, template copy).

Hero: `render(short_sheet_copy(t+"tidy-tabs-purchase-order-tracker.xlsx","PO lines",11,"po"),"PO lines",hide_cols="EFNQRS",name="po")` (as-of date, the 7 lines with Received / Partial / Over by 1 / Late 15 days; legible; do not hide column B, it holds the as-of date in B1). Second card: `render(short_sheet_copy(t+"tidy-tabs-purchase-order-tracker.xlsx","Suppliers",7,"sp"),"Suppliers",hide_cols="NOP",name="sp")` (on-time 50% / 100% / 100%, fill 90% / 60% / 100%, uncapped 125%).

---
## 4. stockout-days-and-estimated-lost-sales-from-daily-inventory-log-countifs-sumifs-excel-google-sheets
Title (topics): How to Count Stockout Days and Estimate Lost Sales in Excel and Google Sheets
Template: templates/tidy-tabs-stockout-lost-sales.xlsx (tabs: Daily log, Products, Summary, How to use)

Layout
- Daily log: A1:A2 notes, header row 5, rows 6:3005: A Date, B SKU, C Units on hand at open, D Units received, E Units sold (inputs; D and E may be blank = 0); F Available, G Left at close, H Day status (Out light orange), I Status. Row 5 repeated.
- Products: A1 note, header row 4, rows 5:54: A SKU, B Name, C Price (inputs), D Status.
- Summary: A1:A2 notes, header row 4, rows 5:54 (one per Products row, nothing to type): A SKU, B Name, C Price, D Days logged, E Out days, F In-stock days, G Sold on in-stock days, H Sold per in-stock day, I Sold per day, all days (shortcut), J Estimated lost units, K Estimated lost sales (J and K light green when above 0), L Lost units by the shortcut, M In-stock days that ended at 0, N Status. Summary P4:Q10: Out days (all products) (Q4), Estimated lost units (Q5), Estimated lost sales (Q6), Log rows showing Check (Q7), Product rows showing Check (Q8), Typed below row 3005 on the log (not read) (Q9), CHECK (Q10). Prints on one page.
- Deviations from the topic: column order on the log is F Available, G Left at close (added), H Day status, I Status (the topic had G Day status, H Status). Days logged = Out days + In-stock days (COUNTIFS on the two words, not COUNTIF of the SKU, so rows with Check are left out). Added: the shortcut's lost units (L) for contrast, Status per product (No days logged, Out every logged day, Check: see Products).

Key formulas (exact text in the file)
- Daily log F6 `=IF(I6<>"OK","",C6+N(D6))`; G6 `=IF(F6="","",F6-N(E6))`; H6 `=IF(F6="","",IF(F6<=0,"Out","In stock"))`
- Daily log I6 `=IF(AND(A6="",TRIM(B6)="",C6="",D6="",E6=""),"",IF(NOT(ISNUMBER(A6)),"Check: date blank or text",IF(TRIM(B6)="","Check: no SKU",IF(COUNTIF(Products!$A$5:$A$54,B6)=0,"Check: SKU not on Products",IF(OR(NOT(ISNUMBER(C6)),AND(D6<>"",NOT(ISNUMBER(D6))),AND(E6<>"",NOT(ISNUMBER(E6)))),"Check: number blank or text",IF(OR(C6<0,N(D6)<0,N(E6)<0),"Check: negative number",IF(N(E6)>C6+N(D6),"Check: sold more than available",IF(COUNTIFS($A$6:$A$3005,A6,$B$6:$B$3005,B6)>1,"Check: SKU and date typed twice","OK"))))))))`
- Summary E5 `=IF(A5="","",COUNTIFS('Daily log'!$B$6:$B$3005,A5,'Daily log'!$H$6:$H$3005,"Out"))`; F5 the same with "In stock"; D5 `=IF(A5="","",E5+F5)`
- G5 `=IF(A5="","",SUMIFS('Daily log'!$E$6:$E$3005,'Daily log'!$B$6:$B$3005,A5,'Daily log'!$H$6:$H$3005,"In stock"))`; H5 `=IF(A5="","",IF(F5=0,"No in-stock days",G5/F5))`; I5 `=IF(A5="","",IF(D5=0,"",G5/D5))`
- J5 `=IF(A5="","",IF(E5=0,0,IF(F5=0,"No in-stock days",ROUND(E5*H5,0))))`; K5 `=IF(A5="","",IF(OR(NOT(ISNUMBER(J5)),C5=""),"",ROUND(J5*C5,2)))`; L5 `=IF(A5="","",IF(OR(E5=0,F5=0),J5,ROUND(E5*I5,0)))`
- M5 `=IF(A5="","",COUNTIFS('Daily log'!$B$6:$B$3005,A5,'Daily log'!$H$6:$H$3005,"In stock",'Daily log'!$G$6:$G$3005,0))`; N5 `=IF(A5="","",IF(LEFT(Products!D5,5)="Check","Check: see Products",IF(D5=0,"No days logged",IF(F5=0,"Out every logged day","OK"))))`
- Q10 `=IF(AND(SUM(D5:D54)>0,Q7=0,Q8=0,Q9=0),"OK","Check")`

Sample values (fictional; recalculated). Products LAV-8 Lavender candle 8 oz $24.00, CED-8 Cedar candle 8 oz $24.00, PUMP-4 Pumpkin candle 4 oz $16.00. Log 84 rows, 10/26/2026 to 11/22/2026 (28 days per SKU; LAV-8 starts with 20, CED-8 12, PUMP-4 30; 24 LAV-8 received 11/09/2026, 24 CED-8 received 11/15/2026).
| SKU | Days logged | Out days | In-stock days | Sold on in-stock days | Per in-stock day | Per day, all days | Lost units | Lost sales | Lost units by the shortcut | Ended at 0 |
|---|---|---|---|---|---|---|---|---|---|---|
| LAV-8 | 28 | 0 | 28 | 33 | 1.179 | 1.179 | 0 | $0.00 | 0 | 0 |
| CED-8 | 28 | 10 | 18 | 23 | 1.278 | 0.821 | 13 | $312.00 | 8 | 1 |
| PUMP-4 | 28 | 1 | 27 | 30 | 1.111 | 1.071 | 1 | $16.00 | 1 | 1 |
CED-8 out 11/05 to 11/14/2026 (sold its last unit 11/04, restocked 11/15); PUMP-4 out 11/22/2026 (ended at 0 on 11/21). Q4:Q10: 11, 14, $328.00, 0, 0, 0, OK. The all-days rate (0.821) is 36% below the in-stock rate (1.278) for CED-8; the shortcut gives 8 lost units instead of 13.

Edge cases tested (all matched the Python model)
- A SKU out every logged day (FIG-8, 28 days at 0): H and J "No in-stock days" (no #DIV/0!), K blank, Status Out every logged day, CHECK OK.
- A SKU never out (LAV-8): 0 lost units, $0.00.
- Missing dates (CED-8 11/07 to 11/09 deleted): 25 days logged, 7 out days, 9 lost units $216.00, shortcut 0.92 a day and 6 units; missing days are simply not counted.
- Received on an out day (CED-8 11/15: 0 on hand + 24 received) counts as in stock (sample).
- A product on Products with no log rows: No days logged, 0 lost.
- Log checks, each its own message: on hand -1, on hand "3" as text, unknown SKU, sold 3 of 2, the same SKU and date twice (both rows), a text date, no SKU, received "1" as text, received -1; a row on row 3010 (Q9). "lav-8" in lower case matched LAV-8 (LibreOffice COUNTIF/COUNTIFS are not case-sensitive) and counted as OK.
- Products checks: "cED-8" next to "CED-8" (both Check: same SKU twice, Summary Check: see Products), price "16.00" as text, a name with no SKU.
- A full log of 3,000 rows (50 SKUs x 60 days), and 200 random logs (1 to 8 SKUs, 1 to 60 days, about 5% of days missing, random restocks, about 1% bad rows).

Pitfalls the post must state
- An estimate, not a forecast: shoppers who bought another scent and rising or falling demand are not in it. No accounting treatment of lost sales.
- Days that sold out during the day count as in stock, so the estimate is on the low side (column M).
- One row per SKU per day, out days included; a missing day is not an out day.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The lost-sales estimate looks small: the all-days average (0.821 vs 1.278 for CED-8, 8 vs 13 units).
- #DIV/0! for a product that was always out: the template shows No in-stock days instead.
- A day is left out: Status Check (sold more than available, SKU typed twice, unknown SKU).

Hero: `render(short_sheet_copy(t+"tidy-tabs-stockout-lost-sales.xlsx","Summary",7,"so"),"Summary",hide_cols="OPQ",name="so")` (the three products with 1.278 vs 0.821 and 13 vs 8 lost units, $312.00 in green; legible).

---
## 5. retail-4-4-5-fiscal-calendar-week-period-quarter-from-date-int-match-excel-google-sheets
Title (topics): How to Get the 4-4-5 Fiscal Week and Period for Any Date in Excel and Google Sheets
Template: templates/tidy-tabs-4-4-5-fiscal-calendar.xlsx (tabs: Calendar, Dates, How to use)

Layout
- Calendar: A1 Pattern / B1 (input dropdown 4-4-5, 4-5-4, 5-4-4), A2 Week 53 goes in period / B2 (input whole 1 to 12, 12 in the sample), C1:C2 and A3 notes. Header row 5. Fiscal years rows 6:25: A Fiscal year (label), B Start date (inputs; one more row with the next year's start), C Weeks (53 light orange), D Last day, E Status. Period table rows 6:17 (formulas only): G Period (1 to 12, shown P1 to P12), H Weeks (52-week year), I First week (52-week year), J First week (53-week year), K Weeks (53-week year). Summary M4:N9: Settings (B1:B2) (N4), Fiscal years ready (N5), Year rows showing Check (N6), 53-week years (N7), Weeks in the period table (N8), CHECK (N9). Row 5 repeated.
- Dates: A1:A2 notes, header row 4, rows 5:1004: A Date (input); B Fiscal year row, C Fiscal year, D Fiscal week, E Period, F Quarter, G Week starts, H Period starts, I Label, J Calendar month (contrast), K Status. Summary M4:N8: Dates typed (N4), Dates placed (N5), Rows showing Check (N6), Typed below row 1004 (not read) (N7), CHECK (N8). Row 4 repeated.
- Deviations from the topic: B2 is a period number (1 to 12) rather than a fixed P12 rule; the period table holds two first-week columns (I for 52-week years, J for 53-week years, where every period after B2 starts a week later), so the period MATCH needs no MIN(D,52). The sample has 6 start rows (FY2023 to FY2027 plus FY2028 as the end marker). Added: checks for a blank row inside the table, a start with a time, no label, the next start as text; a calendar-month column for contrast.

Key formulas (exact text in the file)
- Calendar C6 `=IF(OR(NOT(ISNUMBER(B6)),NOT(ISNUMBER(B7))),"",(B7-B6)/7)`; D6 `=IF(C6="","",B7-1)`
- Calendar E7 `=IF(AND(TRIM(A7)="",B7=""),"",IF(NOT(ISNUMBER(B7)),"Check: start blank or text",IF(B7<>INT(B7),"Check: start has a time",IF(B6="","Check: blank row above",IF(TRIM(A7)="","Check: no label",IF(B8="","Next year start (end of the table)",IF(NOT(ISNUMBER(B8)),"Check: next start is text",IF(B8<=B7,"Check: starts out of order",IF(OR(C7<>INT(C7),AND(C7<>52,C7<>53)),"Check: not 52 or 53 whole weeks","OK")))))))))` (E6 the same without the blank-row test)
- Calendar H6 `=IF($N$4<>"OK","",VALUE(MID($B$1,2*MOD(G6-1,3)+1,1)))`; I6 `=IF($N$4<>"OK","",1)`; I7 `=IF($N$4<>"OK","",I6+H6)`; J6 `=IF($N$4<>"OK","",I6+IF(G6>$B$2,1,0))`; K6 `=IF($N$4<>"OK","",H6+IF(G6=$B$2,1,0))`
- N4 `=IF(AND(OR(B1="4-4-5",B1="4-5-4",B1="5-4-4"),ISNUMBER(B2)),IF(AND(B2>=1,B2<=12,B2=INT(B2)),"OK","Check"),"Check")`; N9 `=IF(AND(N4="OK",N5>0,N6=0,N8=52),"OK","Check")`
- Dates B5 `=IF(OR(NOT(ISNUMBER(A5)),Calendar!$N$9<>"OK"),"",IF(A5<Calendar!$B$6,"",MATCH(A5,Calendar!$B$6:$B$25,1)))`
- C5 `=IF(B5="","",IF(INDEX(Calendar!$C$6:$C$25,B5)="","",INDEX(Calendar!$A$6:$A$25,B5)))`; D5 `=IF(C5="","",INT((A5-INDEX(Calendar!$B$6:$B$25,B5))/7)+1)`
- E5 `=IF(D5="","",IF(INDEX(Calendar!$C$6:$C$25,B5)=53,MATCH(D5,Calendar!$J$6:$J$17,1),MATCH(D5,Calendar!$I$6:$I$17,1)))`; F5 `=IF(E5="","",INT((E5-1)/3)+1)`
- G5 `=IF(D5="","",INDEX(Calendar!$B$6:$B$25,B5)+(D5-1)*7)`; H5 `=IF(E5="","",INDEX(Calendar!$B$6:$B$25,B5)+(IF(INDEX(Calendar!$C$6:$C$25,B5)=53,INDEX(Calendar!$J$6:$J$17,E5),INDEX(Calendar!$I$6:$I$17,E5))-1)*7)`
- I5 `=IF(E5="","",C5&" P"&E5&" W"&D5)`; J5 `=IF(ISNUMBER(A5),MONTH(A5),"")`; K5 `=IF(A5="","",IF(NOT(ISNUMBER(A5)),"Check: date is text",IF(Calendar!$N$9<>"OK","Check: Calendar tab",IF(A5<Calendar!$B$6,"Check: before the first fiscal year",IF(C5="","Check: after the last fiscal year","OK")))))`

Sample values (fictional example start dates, all Sundays: FY2023 01/29/2023, FY2024 02/04/2024, FY2025 02/02/2025, FY2026 02/01/2026, FY2027 01/31/2027, FY2028 01/30/2028 as the end marker; 4-4-5, week 53 in P12; recalculated). Weeks: FY2023 53 (last day 02/03/2024), FY2024 to FY2027 52 (last days 02/01/2025, 01/31/2026, 01/30/2027, 01/29/2028). Period table 4-4-5: weeks 4,4,5 x 4; first weeks 1, 5, 9, 14, 18, 22, 27, 31, 35, 40, 44, 48; P12 has 6 weeks in a 53-week year. N4:N9 OK, 5, 0, 1, 52, OK.
| Date | Fiscal year | Week | Period | Quarter | Week starts | Period starts | Label | Calendar month |
|---|---|---|---|---|---|---|---|---|
| 02/01/2026 | FY2026 | 1 | 1 | 1 | 02/01/2026 | 02/01/2026 | FY2026 P1 W1 | 2 |
| 02/28/2026 | FY2026 | 4 | 1 | 1 | 02/22/2026 | 02/01/2026 | FY2026 P1 W4 | 2 |
| 03/01/2026 | FY2026 | 5 | 2 | 1 | 03/01/2026 | 03/01/2026 | FY2026 P2 W5 | 3 |
| 04/04/2026 | FY2026 | 9 | 3 | 1 | 03/29/2026 | 03/29/2026 | FY2026 P3 W9 | 4 |
| 04/05/2026 | FY2026 | 10 | 3 | 1 | 04/05/2026 | 03/29/2026 | FY2026 P3 W10 | 4 |
| 11/25/2026 | FY2026 | 43 | 10 | 4 | 11/22/2026 | 11/01/2026 | FY2026 P10 W43 | 11 |
| 11/27/2026 | FY2026 | 43 | 10 | 4 | 11/22/2026 | 11/01/2026 | FY2026 P10 W43 | 11 |
| 12/31/2026 | FY2026 | 48 | 12 | 4 | 12/27/2026 | 12/27/2026 | FY2026 P12 W48 | 12 |
| 01/30/2027 | FY2026 | 52 | 12 | 4 | 01/24/2027 | 12/27/2026 | FY2026 P12 W52 | 1 |
| 01/31/2027 | FY2027 | 1 | 1 | 1 | 01/31/2027 | 01/31/2027 | FY2027 P1 W1 | 1 |
| 01/28/2024 | FY2023 | 53 | 12 | 4 | 01/28/2024 | 12/24/2023 | FY2023 P12 W53 | 1 |
| 02/03/2024 | FY2023 | 53 | 12 | 4 | 01/28/2024 | 12/24/2023 | FY2023 P12 W53 | 2 |
| 02/04/2024 | FY2024 | 1 | 1 | 1 | 02/04/2024 | 02/04/2024 | FY2024 P1 W1 | 2 |
| 12/31/2023 | FY2023 | 49 | 12 | 4 | 12/31/2023 | 12/24/2023 | FY2023 P12 W49 | 12 |
| 01/29/2023 | FY2023 | 1 | 1 | 1 | 01/29/2023 | 01/29/2023 | FY2023 P1 W1 | 1 |
| 02/01/2025 | FY2024 | 52 | 12 | 4 | 01/26/2025 | 12/29/2024 | FY2024 P12 W52 | 2 |
Dates N4:N8: 16, 16, 0, 0, OK.

Edge cases tested (all matched the Python model; the Dates results were also checked against a separate walk-the-periods model)
- Every day from 01/29/2023 to 01/29/2028 (1,827 days) for each of 4-4-5, 4-5-4 and 5-4-4 (6 workbooks). Period tables: 4-5-4 first weeks 1, 5, 10, 14, 18, 23, 27, 31, 36, 40, 44, 49; 5-4-4 first weeks 1, 6, 10, 14, 19, 23, 27, 32, 36, 40, 45, 49.
- Week 53 sent to P1, P3, P6 and P11 for each pattern (FY2023, 371 days each): with P1 in 4-4-5, P1 has 5 weeks and 02/26/2023 to 03/01/2023 are FY2023 P1 W5; later periods start one week later (P2 at week 6).
- First and last day of each fiscal year (above); 01/28/2023 and 01/01/2020 (Check: before the first fiscal year); 01/30/2028 and 06/01/2030 (Check: after the last fiscal year); "02/01/2026" as text (Check: date is text); 11/25/2026 3:45 PM (a date with a time: FY2026 P10 W43, OK).
- Saturday starts (01/28/2023 ...) and Monday starts with a 53-week year (09/30/2024 ..., week 53 in P1): both work (allowed; the post says some calendars start on Saturday or Monday).
- Calendar checks: starts out of order, a blank row inside the table, 54 weeks, a start 1 day off (not whole weeks), a start typed as text, a start with a time, no label, pattern "4-4-4", week-53 period 13 or "12" as text, only one start row: each shows its Check, CHECK Check, and every date shows Check: Calendar tab.
- A date typed on row 1005 (N7 1). 30 random calendars (random start month and day, 2 to 19 years, random 52/53-week years, random pattern and week-53 period) with 1,000 random dates each, including dates before and after the table.

Pitfalls the post must state
- The reader types their own start dates; the sample dates are examples, not the NRF's or any company's calendar (link the NRF page only if someone confirms it loads).
- Where week 53 goes is the reader's calendar rule; the sheet does not recommend one. No claim that any business must use a 52-53 week year (tax year elections are gate 3).
- Add the next year's start, or dates after the last year show Check.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- 01/30/2027 shows month 1 but is FY2026 P12 W52; 02/28/2026 is month 2 but P1 (calendar month vs fiscal period).
- A typed 4-4-5 table used with 4-5-4 puts dates in the wrong period (mutation run: first weeks 1, 5, 9 ... instead of 1, 5, 10 ...).
- Every date says Check: Calendar tab: a start out of order, a blank row, or a year that is not 52 or 53 weeks.
- Dates after the last year: add the next start date.

Hero: `render(short_sheet_copy(t+"tidy-tabs-4-4-5-fiscal-calendar.xlsx","Dates",20,"fy"),"Dates",hide_cols="BLMN",name="fy")` (16 sample dates with week, period, quarter, label and calendar month; very legible). Optional second card: `render(short_sheet_copy(t+"tidy-tabs-4-4-5-fiscal-calendar.xlsx","Calendar",17,"cal"),"Calendar",name="cal")` (not trial-rendered).

---
Not built or tested (so not for the posts): anything in Excel or Google Sheets; non-US holidays or a Thanksgiving other than the US rule; ad spend, discounts or profit; hourly sales; time zones of online orders; amounts typed with 3 decimals beyond the fact that the sheet flags them; tax categories or deductibility; bank deposits or bookkeeping entries; more than one petty cash box; payment terms, invoices, three-way match, partial cancellations, units of measure or currency; supplier lead-time forecasts; substitution between products, demand trends, seasonality, forecasting, safety stock or reorder points; accounting treatment of lost sales; the NRF or any company's real fiscal calendar; 4-4-4/13-period calendars; fiscal years longer than 53 weeks; WEEKNUM/ISOWEEKNUM versions; SEQUENCE, LET, LAMBDA or Power Query versions; typing by hand in the GUI (all inputs were written with openpyxl and recalculated headless).
