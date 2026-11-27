# Handoff 2026-11-28: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-28.md (researcher and hook stages: tasks/team-topics-2026-11-28.md, tasks/hooks-2026-11-28.md). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`december_daily_sales_targets` with `DST_LOG_LAST`, `year_end_sales_projection` with `YEP_LOG_LAST`, `unbilled_hours_by_client` with `UNB_LOG_LAST`/`UNB_CLIENTS_LAST`, `processing_time_percentile` with `PT_ORDERS_LAST`/`PT_HOL_LAST` and helper `_pt_sample_orders`, `recurring_charges_finder` with `RC_CHARGES_LAST`/`RC_ROWS_LAST` and helper `_rc_sample_charges`, shared helper `_settings_rows`; registered in `__main__` after `price_volume_change()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-28 템플릿 제작/검증 (LibreOffice Calc 24.2.7.2)"), the hooks file updated (no "to confirm" left). Heroes are not made (designer; trial renders in scratch only, calls below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 1,133 recalculated workbooks, 802,387 cell comparisons, 0 mismatches against independent Python models in the final runs, 0 error values in any recalculated workbook. A mutation run (one wrong formula per template, the edge cases plus 5 random sets each) produced mismatches in all five (199 / 73 / 49 / 1,665 / 224: each day's own share rounded instead of the running total; the projection divided by the rounded share in B9; `D=""` without TRIM; NETWORKDAYS without the -1; the COUNTIF key without the wildcard escape), so the harness catches wrong formulas. All five files exported to PDF: every page US Letter (792 x 612 landscape data tabs, 612 x 792 How to use), no error values, Err: or ### in the PDF text.

Repo note: the whole script was **not** run. The five files were built by importing scripts/build-templates.py with `sys.dont_write_bytecode = True` and calling only the five functions, so no older template and no scripts/__pycache__/*.pyc was rewritten (nothing to restore; `git status` shows no modified .xlsx or .pyc). `git status` after this stage: scripts/build-templates.py (899 lines added, 0 removed), tasks/briefs.md, new tasks/hooks-2026-11-28.md (untracked from the hook stage, now updated), this handoff, five new .xlsx. (tasks/team-topics-2026-11-28.md and tasks/topics-seed.md were committed by someone else during this run as "Add 2026-11-28 research topics".)

Conventions in all five: How to use tab last; US Letter, fit to one page wide, header row repeated on the data tabs; data tabs landscape; yellow = inputs, gray = formulas; red = any Status starting "Check" and any CHECK / settings cell that is not OK; light orange = warnings that are not errors (closed day weight 0, behind target, Outside the month, Before last year / After the as-of date, Yes: spaces only, older than B2 days, Not shipped yet, Outside the date range, Refund, Price change, Not seen); light green = highlights (at or above target, the pattern projection, unbilled totals, the promise and its coverage row, Subscription, yearly total). Functions used in the five files: AND, AVERAGE, COUNT, COUNTA, COUNTIF, COUNTIFS, DATE, DAY, IF, INDEX, INT, ISNUMBER, ISTEXT, LEFT, LEN, MATCH, MAX, MEDIAN, MIN, MONTH, N, NETWORKDAYS, NOT, OR, PERCENTILE, ROUND, ROUNDUP, SUBSTITUTE, SUM, SUMIF, SUMIFS, SUMPRODUCT, TRIM, WEEKDAY, YEAR. No TODAY, FILTER, UNIQUE, SORT, LET, LAMBDA, REDUCE, XLOOKUP, TEXTSPLIT, MINIFS/MAXIFS, SEQUENCE, CHAR or UNICHAR; 0 `_xlfn` in the five saved files and 0 after the LibreOffice round trip (PERCENTILE stays PERCENTILE).

Help links: re-checked 2026-11-27 with `curl -sIL -o /dev/null -w '%{http_code}'`. **Every Microsoft link in the topics file returned 200 (19). Every Google link returned 404 to that HEAD check (17)**, but 200 to a plain GET (`curl -sL -o /dev/null -w '%{http_code}'`). Same situation as 11-26: the publisher decides whether a GET 200 is enough.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-28-team-01-december-daily-sales-targets-from-monthly-goal-weighted-by-weekday-ahead-or-behind-weekday-index-sumifs-excel-google-sheets | tidy-tabs-december-daily-sales-targets.xlsx | 235 workbooks, 80,535 comparisons |
| 2 | 2026-11-28-team-02-project-year-end-sales-from-year-to-date-run-rate-vs-last-year-pattern-sumifs-date-excel-google-sheets | tidy-tabs-year-end-sales-projection.xlsx | 227 workbooks, 75,750 comparisons |
| 3 | 2026-11-28-team-03-unbilled-hours-and-amount-by-client-blank-invoice-number-sumifs-sumproduct-excel-google-sheets | tidy-tabs-unbilled-hours-by-client.xlsx | 222 workbooks, 167,740 comparisons |
| 4 | 2026-11-28-team-04-set-order-processing-time-from-90th-percentile-business-days-networkdays-percentile-excel-google-sheets | tidy-tabs-processing-time-percentile.xlsx | 225 workbooks, 44,764 comparisons |
| 5 | 2026-11-28-team-05-find-recurring-charges-and-subscriptions-in-bank-export-distinct-months-countifs-sumif-excel-google-sheets | tidy-tabs-recurring-charges-finder.xlsx | 224 workbooks, 433,598 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets." For the unbilled hours post add: "The SUMIFS blank-criterion result and the CHAR(160) difference were seen in LibreOffice Calc only." For the recurring charges post add: "The wildcard miscount was reproduced in LibreOffice Calc only."

Hooks check (tasks/hooks-2026-11-28.md): every "to confirm" was replaced. Posts 1, 2, 3 and 5: every number confirmed as the probe had it; formula text updated where the built cell differs (post 1 `$D$36` and `$M$5:$M$11`; post 3 the SUMPRODUCT with the Status test and the hours as a separate argument; post 4 `Orders!$D$5:$D$504`). Post 4: the sample was rebuilt as the topics file asked (the 90th percentile now lands between two values), so its numbers changed and the sentences were rewritten: 68 shipped (not 46), average 2.72 (not 2.80), 90th percentile 4.3 rounded up to 5, 65 of 68 (95.59%, not 97.83%), 3 days 67.65%, 4 days 89.71%. Final titles in the hooks file: posts 1, 2 and 5 keep the topic title; post 3 is "How to Total Unbilled Hours by Client in Excel and Google Sheets"; post 4 is "How to Calculate Order Processing Time in Excel and Google Sheets".

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=..., name=...)`, `short_sheet_copy(src, sheet, last_row, name)` and `compact_copy(src, sheet, last_row, name, hide_rows=..., labels=...)`, with `t = "templates/"`). Trial renders at 1600x900 were made under /usr/bin/python3 (PYTHONDONTWRITEBYTECODE=1, importing make-heroes.py and calling `save` with a scratch path) into /tmp/claude-0/-home-user-tidy-tabs/9df4f143-e162-5575-9090-17282c79a118/scratchpad/build/hero/ only; nothing was written to images/. All five calls below rendered legibly (looked at one by one).

---
## 1. december-daily-sales-targets-from-monthly-goal-weighted-by-weekday-ahead-or-behind-weekday-index-sumifs-excel-google-sheets
Post: posts/ready/2026-11-28-team-01-december-daily-sales-targets-from-monthly-goal-weighted-by-weekday-ahead-or-behind-weekday-index-sumifs-excel-google-sheets.md
Title (hooks final): How to Set Daily Sales Targets From a Monthly Goal by Day of the Week in Excel and Google Sheets
Template: templates/tidy-tabs-december-daily-sales-targets.xlsx (tabs: Targets, Sales log, How to use)

Layout
- Targets: A1 Month starts / B1 (12/01/2026), A2 Monthly goal / B2 ($9,000.00), A3 Sales counted through / B3 (12/12/2026) (inputs), C1:C3 notes. Header row 4, rows 5:35 (one per day; blank past the month's last day): A Date (shown as "Tue 12/01/2026"), B Override weight (input), C Weight, D Cumulative weight, E Cumulative target, F Daily target, G Actual sales, H Actual to date, I Cumulative still needed, J Still needed per day. Row 36 Month total: D36 total weight, F36 sum of targets, G36 actual, J36 sum of still needed. Weekday weights L4:M11 (L Weekday Sunday to Saturday, M Weight, inputs; L12 "0 = closed"). Summary O4:P20: Settings (P4), Weights (P5), Days in the month (P6), Open days (P7), Closed days (P8), Even split per day (P9), Even split per open day (P10), Target to date (P11), Sold to date (P12), Ahead (+) or behind (-) (P13), Ahead or behind % (P14), Still needed for the month (P15), Weight left after B3 (P16), Days left after B3 (P17), Sum of still needed per day (P18), Sales log CHECK (P19), CHECK (P20). Row 4 repeated; the Targets tab prints on one landscape page.
- Sales log: A1:A2 notes. Header row 4, rows 5:1004: A Date, B Amount, C Note (inputs), D Status. Summary F4:G9: Rows OK, Total of OK rows, Outside the month, Rows showing Check, Typed below row 1004, CHECK.

Key formulas (exact text in the file)
- A5 `=IF($P$4<>"OK","",$B$1)`; A6 (and down) `=IF(A5="","",IF(DAY(A5+1)=1,"",A5+1))`
- C5 `=IF(A5="","",IF(B5<>"",B5,INDEX($M$5:$M$11,WEEKDAY(A5))))`; D5 `=IF(A5="","",SUM($C$5:C5))`; D36 `=SUM(C5:C35)`
- E5 `=IF(OR(A5="",$P$5<>"OK"),"",ROUND($B$2*D5/$D$36,2))`; F5 `=IF(E5="","",ROUND(E5-N(E4),2))` (N of the header in E4 is 0, so the same formula works on the first row)
- G5 `=IF(A5="","",SUMIFS('Sales log'!$B$5:$B$1004,'Sales log'!$A$5:$A$1004,A5,'Sales log'!$D$5:$D$1004,"OK"))`; H5 `=IF(OR(A5="",A5>$B$3),"",SUM($G$5:G5))`
- I5 `=IF(OR(A5="",A5<=$B$3,$P$5<>"OK",N($P$16)<=0),"",ROUND($P$15*SUMIFS($C$5:$C$35,$A$5:$A$35,">"&$B$3,$A$5:$A$35,"<="&A5)/$P$16,2))`; J5 `=IF(I5="","",ROUND(I5-N(I4),2))`
- P4 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B1=INT(B1),B1>=1,B2>=0,B3=INT(B3)),IF(DAY(B1)=1,"OK","Check"),"Check"),"Check")`
- P5 `=IF(P4<>"OK","Check",IF(AND(COUNT(M5:M11)=7,MIN(M5:M11)>=0,SUMPRODUCT((A5:A35<>"")*ISTEXT(C5:C35))=0,MIN(C5:C35)>=0,D36>0),"OK","Check"))`
- P6 `=IF(P4<>"OK","",DATE(YEAR(B1),MONTH(B1)+1,1)-B1)`; P11 `=IF(P5<>"OK","",IF(B3<B1,0,INDEX(E5:E35,MIN(B3-B1+1,P6))))`; P12 `=IF(P4<>"OK","",SUMIFS('Sales log'!$B$5:$B$1004,'Sales log'!$D$5:$D$1004,"OK",'Sales log'!$A$5:$A$1004,"<="&B3))`; P13 `=IF(P11="","",ROUND(P12-P11,2))`; P14 `=IF(OR(P11="",N(P11)=0),"",P13/P11)`; P15 `=IF(P4<>"OK","",MAX(0,ROUND(B2-P12,2)))`; P16 `=IF(P5<>"OK","",SUMIFS(C5:C35,A5:A35,">"&B3))`
- P20 `=IF(AND(P4="OK",P5="OK",P19="OK"),IF(AND(ROUND(F36-B2,2)=0,ROUND(J36-IF(N(P16)>0,P15,0),2)=0,ROUND(G36-'Sales log'!G5,2)=0),"OK","Check"),"Check")`
- Sales log D5 `=IF(AND(A5="",B5="",TRIM(C5)=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(B5="","Check: no amount",IF(NOT(ISNUMBER(B5)),"Check: amount is text",IF(Targets!$P$4<>"OK","Check: Targets settings",IF(OR(A5<Targets!$B$1,A5>=Targets!$B$1+Targets!$P$6),"Outside the month (not counted)","OK")))))))`

Sample values (fictional; recalculated). Weights Sun 1.5, Mon 0, Tue 1, Wed 1, Thu 1.2, Fri 1.8, Sat 2.5; overrides 12/19 3.5, 12/24 1.0, 12/25 0, 12/31 0.8. Sales log: 15 rows, one 11/28/2026 $455.60 (Outside the month) and 14 in December ($3,704.75).
| Date | Weight | Cumulative target | Daily target | Actual | Actual to date | Still needed per day |
|---|---|---|---|---|---|---|
| Tue 12/01/2026 | 1.0 | $238.10 | $238.10 | $212.40 | $212.40 | |
| Wed 12/02/2026 | 1.0 | $476.19 | $238.09 | $188.15 | $400.55 | |
| Fri 12/04/2026 | 1.8 | $1,190.48 | $428.58 | $397.75 | $1,043.20 | |
| Sat 12/05/2026 | 2.5 | $1,785.71 | $595.23 | $652.75 | $1,695.95 | |
| Mon 12/07/2026 | 0.0 | $2,142.86 | $0.00 | $0.00 | $1,997.05 | |
| Sat 12/12/2026 | 2.5 | $3,928.57 | $595.24 | $679.55 | $3,704.75 | |
| Sun 12/13/2026 | 1.5 | $4,285.71 | $357.14 | $0.00 | | $372.90 |
| Tue 12/15/2026 | 1.0 | $4,523.81 | $238.10 | $0.00 | | $248.61 |
| Wed 12/16/2026 | 1.0 | $4,761.90 | $238.09 | $0.00 | | $248.60 |
| Sat 12/19/2026 | 3.5 | $6,309.52 | $833.33 | $0.00 | | $870.11 |
| Thu 12/24/2026 | 1.0 | $7,380.95 | $238.09 | $0.00 | | $248.61 |
| Fri 12/25/2026 | 0.0 | $7,380.95 | $0.00 | $0.00 | | $0.00 |
| Sat 12/26/2026 | 2.5 | $7,976.19 | $595.24 | $0.00 | | $621.51 |
| Thu 12/31/2026 | 0.8 | $9,000.00 | $190.48 | $0.00 | | $198.88 |
| Month total | 37.8 | | $9,000.00 | $3,704.75 | | $5,295.25 |
P4:P20: OK, OK, 31, 26, 5, $290.32, $346.15, $3,928.57, $3,704.75, -$223.82, -5.7%, $5,295.25, 21.3, 19, $5,295.25, OK, OK. Sales log G4:G9: 14, $3,704.75, 1, 0, 0, OK.

Edge cases tested (all matched the Python model)
- February 2027 (28 days), February 2028 (29), November 2026 (30), January 2027 (31): rows past the month blank, F36 = goal.
- Month start 12/02/2026, text, or with a time (P4 Check, grid blank, CHECK Check). As-of before the month (target to date $0.00, the whole goal spread over every day), on day 1, on the last day, after the month (still needed shown, no days left, J blank, CHECK OK), as text, blank, with a time (Check).
- Weights: all 0 (Weights Check, targets blank, no error), all 0 with an override 0, a negative, text or blank weekday weight, a negative or text override (Check); an override 0 on an open day (closed); an override on every day.
- Goal 0 (all targets $0.00), $0.01 (one day gets $0.01), text or negative (Check), already beaten ($3,000.00 goal: still needed $0.00 and $0.00 a day, CHECK OK), exactly met.
- Sales log: a date with a time (Check: date has a time, CHECK Check), text date, blank amount, text amount, a note only, an amount with no date, a January row (Outside the month); empty log; a row typed on row 1005; 1,000 log rows; 200 random sets (months 2020 to 2035, random weights and overrides, goals $0.00 to $60,000.00, as-of 5 days before to 5 days after the month, up to 60 sales inside and outside the month).
- Scratch copy (not in the template): rounding each remaining day on its own, `ROUND($P$15*C/$P$16,2)`, summed to $5,295.21, 4 cents short of $5,295.25; rounding each day's own share of the month, `ROUND($B$2*C5/$D$36,2)`, summed to $9,000.03.

Pitfalls the post must state
- The weights are the reader's own guess or last year's weekday shares (link sales-by-day-of-week-sumif); 0 means closed. The targets are a plan, not a forecast; no advice on what goal to set. Sample shop and sales are fictional.
- The running total is rounded, not each day, so the days add up to the goal to the cent.
- Sales dates must be plain dates; a date with a time is flagged and not counted.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Daily targets add up to a few cents more or less than the goal: each day was rounded on its own ($9,000.03 and $5,295.21 in the scratch copies); round the running total instead.
- Targets all blank and CHECK Check: every weight is 0, or a weight is text or negative, or B1 is not the first of a month.
- A day's sales are missing and CHECK is red: the Sales log date has a time of day (Check: date has a time).
- Still needed shows $0.00 a day: the goal is already reached.

Hero: `render(short_sheet_copy(t+"tidy-tabs-december-daily-sales-targets.xlsx","Targets",36,"dst"),"Targets",hide_cols="DI",name="dst")` (the full December grid with yellow overrides, orange closed days and behind-target cells, the weight table and the summary with -$223.82; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): WEEKDAY https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a and https://support.google.com/docs/answer/3092985; INDEX https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd and https://support.google.com/docs/answer/3098242; SUMIFS https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b and https://support.google.com/docs/answer/3238496; ROUND https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c and https://support.google.com/docs/answer/3093440.

---
## 2. project-year-end-sales-from-year-to-date-run-rate-vs-last-year-pattern-sumifs-date-excel-google-sheets
Post: posts/ready/2026-11-28-team-02-project-year-end-sales-from-year-to-date-run-rate-vs-last-year-pattern-sumifs-date-excel-google-sheets.md
Title (hooks final): How to Project Year-End Sales From Year-to-Date Sales in Excel and Google Sheets
Template: templates/tidy-tabs-year-end-sales-projection.xlsx (tabs: Projection, Sales log, How to use)

Layout
- Projection (one page): A labels, B values, D notes. Inputs B1 Sales counted through (11/25/2026), B2 Annual goal (optional, $48,000.00), B3 Same date last year (optional, blank). B4 Same date last year, B5 Sales this year to date, B6 This year starts, B7 Last year to the same date, B8 Last year full year, B9 Share of last year made by the same date, B10 Days so far, B11 Days in this year, B12 Share of the year gone. Header row 14 (Year-end projection | Last year's pattern | Straight-line run rate | How it is figured), rows 15 Projected total, 16 Ahead (+) or short (-) of the goal, 17 Expected for the rest of the year, 18 Expected per remaining day (B pattern, C run rate). Header row 20, rows 21 Still needed, 22 Days left, 23 Needed per remaining day, 24 Last year after the same date. Header row 26 Checks: B27 Settings, B28 First date in the Sales log, B29 Log covers last year from its first week, B30 Same date last year (Same month and day / Moved to 03/01 / Typed in B3), B31 Sales log CHECK, B32 CHECK.
- Sales log: A1:A2 notes. Header row 4, rows 5:5004: A Date, B Amount, C Note (inputs), D Status. Summary F4:G9: Rows OK, Total of OK rows, Not counted (before last year or after B1), Rows showing Check, Typed below row 5004, CHECK.

Key formulas (exact text in the file)
- B4 `=IF($B$27<>"OK","",IF(B3<>"",B3,DATE(YEAR(B1)-1,MONTH(B1),DAY(B1))))`; B6 `=IF($B$27<>"OK","",DATE(YEAR(B1),1,1))`
- B5 `=IF($B$27<>"OK","",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$D$5:$D$5004,"OK",'Sales log'!$A$5:$A$5004,">="&B6,'Sales log'!$A$5:$A$5004,"<"&B1+1))`
- B7 `=IF($B$27<>"OK","",SUMIFS('Sales log'!$B$5:$B$5004,'Sales log'!$D$5:$D$5004,"OK",'Sales log'!$A$5:$A$5004,">="&DATE(YEAR(B1)-1,1,1),'Sales log'!$A$5:$A$5004,"<"&B4+1))`; B8 the same with `"<"&B6`
- B9 `=IF(OR(B8="",N(B8)<=0),"",ROUND(B7/B8,4))`; B10 `=IF(B6="","",B1-B6+1)`; B11 `=IF(B6="","",DATE(YEAR(B1),12,31)-B6+1)`; B12 `=IF(B6="","",ROUND(B10/B11,4))`
- B15 `=IF(B6="","",IF(N(B8)<=0,"No sales last year",IF(N(B7)<=0,"No sales last year by this date",ROUND(B5*B8/B7,2))))`; C15 `=IF(B6="","",ROUND(B5/B10*B11,2))`
- B16 `=IF(OR(NOT(ISNUMBER(B15)),B2=""),"",ROUND(B15-$B$2,2))`; B17 `=IF(NOT(ISNUMBER(B15)),"",ROUND(B15-$B$5,2))`; B18 `=IF(OR(B17="",N($B$22)=0),"",ROUND(B17/$B$22,2))` (C16:C18 the same on C15)
- B21 `=IF(OR(B6="",B2=""),"",MAX(0,ROUND(B2-B5,2)))`; B22 `=IF(B6="","",B11-B10)`; B23 `=IF(OR(B21="",N(B22)=0),"",ROUND(B21/B22,2))`; B24 `=IF(B6="","",ROUND(B8-B7,2))`
- B27 `=IF(NOT(ISNUMBER(B1)),"Check",IF(OR(B1<>INT(B1),B1<367),"Check",IF(AND(B2<>"",NOT(ISNUMBER(B2))),"Check",IF(N(B2)<0,"Check",IF(B3="","OK",IF(NOT(ISNUMBER(B3)),"Check",IF(B3<>INT(B3),"Check",IF(YEAR(B3)<>YEAR(B1)-1,"Check","OK"))))))))`
- B29 `=IF(B27<>"OK","Check",IF(B28="","Check: log is empty",IF(B28<=DATE(YEAR(B1)-1,1,7),"OK","Check: log starts after 01/07 of last year")))`; B30 `=IF(B27<>"OK","",IF(B3<>"","Typed in B3",IF(MONTH(B4)<>MONTH(B1),"Moved to 03/01 (no 02/29 last year)","Same month and day")))`; B32 `=IF(AND(B27="OK",B29="OK",B31="OK",ISNUMBER(B15)),"OK","Check")`
- Sales log D5 `=IF(AND(A5="",B5="",TRIM(C5)=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(B5="","Check: no amount",IF(NOT(ISNUMBER(B5)),"Check: amount is text",IF(Projection!$B$27<>"OK","Check: Projection settings",IF(A5<DATE(YEAR(Projection!$B$1)-1,1,1),"Before last year (not counted)",IF(A5>=Projection!$B$1+1,"After the as-of date (not counted)","OK")))))))`

Sample values (fictional gift shop; recalculated). Log 595 rows, 01/01/2025 to 11/25/2026, closed Mondays, December the busiest month.
| Line | Last year's pattern | Straight-line run rate |
|---|---|---|
| Projected total for 2026 | $49,252.37 | $43,599.24 |
| Ahead (+) or short (-) of $48,000.00 | $1,252.37 | -$4,400.76 |
| Expected for the rest of the year | $9,953.33 | $4,300.20 |
| Expected per remaining day | $276.48 | $119.45 |
B4:B12: 11/25/2025, $39,299.04, 01/01/2026, $35,696.16, $44,736.98, 79.79%, 329, 365, 90.14%. B21:B24: $8,700.96, 36, $241.69, $9,040.82. B27:B32: OK, 01/01/2025, OK, Same month and day, OK, OK. Sales log: 595 OK, $84,036.02, 0, 0, 0, OK.

Edge cases tested (all matched the Python model)
- As-of 01/01/2026 and 12/31/2026; 02/29/2028 (B4 03/01/2027, B30 "Moved to 03/01 (no 02/29 last year)") and the same with 02/28/2027 typed in B3 (B30 "Typed in B3"); as-of 11/25/2028 with B11 366.
- No sales last year by the date (log from 06/01/2025, as-of 03/01/2026): B15 "No sales last year by this date", run rate still shown, B29 Check, CHECK Check. No sales last year at all: "No sales last year". Logs starting 05/01/2025 (Check), 01/07/2025 (OK), 01/08/2025 (Check).
- Times of day: 11/25/2026 3:30 PM counted in B5 (YTD $39,419.04 with a $120.00 sale), 11/26/2026 12:30 AM After the as-of date, 11/25/2025 11:59 PM counted in B7.
- As-of text, blank or with a time (Check, everything blank); goal blank (goal lines blank), text or negative (Check), already beaten ($0.00 still needed); B3 a date this year or text (Check); empty log (B29 Check: log is empty); 7 bad or outside rows; refunds; a row on row 5005; 5,000 rows; 200 random logs (years 2001 to 2035, logs starting on 01/01 up to 500 days late, closed weekdays, times on the as-of date, refunds, text dates).
- Mutation: dividing by the rounded share in B9 instead of B7/B8 gave $49,253.09 for the sample ($0.72 off).

Pitfalls the post must state
- The pattern projection assumes this year keeps last year's seasonal shape; a moving holiday (Black Friday 11/28/2025 vs 11/27/2026) shifts the share by a few days (link compare-black-friday-weekend-sales-to-last-year). No promise about the result; no tax or budget advice. Sample sales are fictional.
- The log must cover last year from its first week; a log that starts later makes the share too high and the projection too low.
- 02/29 compares with 03/01 of last year unless the reader types a date in B3.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The projection cell says "No sales last year by this date": nothing in the log before the same date last year (log starts too late).
- CHECK red with "Check: log starts after 01/07 of last year": the log does not reach back to the start of last year.
- Using the rounded share (79.79%) instead of B7/B8 moved the projection by $0.72 in the sample.
- B30 says Moved to 03/01: the as-of date is 02/29.

Hero: `render(short_sheet_copy(t+"tidy-tabs-year-end-sales-projection.xlsx","Projection",24,"yep"),"Projection",name="yep")` (inputs, YTD and last year, the two projections side by side with green and orange gaps, and the still-needed block; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): SUMIFS as in post 1; DATE https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349 and https://support.google.com/docs/answer/3092969; ROUND as in post 1.

---
## 3. unbilled-hours-and-amount-by-client-blank-invoice-number-sumifs-sumproduct-excel-google-sheets
Post: posts/ready/2026-11-28-team-03-unbilled-hours-and-amount-by-client-blank-invoice-number-sumifs-sumproduct-excel-google-sheets.md
Title (hooks final): How to Total Unbilled Hours by Client in Excel and Google Sheets
Template: templates/tidy-tabs-unbilled-hours-by-client.xlsx (tabs: Time log, Clients, How to use)

Layout
- Time log: A1:A2 notes. Header row 4, rows 5:1004: A Date, B Client, C Hours, D Invoice # (inputs), E Amount, F Looks blank?, G Status, H Unbilled age (days) (0 in gray for billed rows). Summary J4:K9: Rows OK, Hours on OK rows, After the as-of date, Rows showing Check, Typed below row 1004, CHECK.
- Clients: A1 As of / B1 (11/27/2026), A2 Flag unbilled entries older than (days) / B2 (30) (inputs), C1:C2 notes. Header row 4, rows 5:54: A Client, B Rate (inputs), C Unbilled hours, D Unbilled $, E Oldest unbilled, F Older than B2 days, G Unbilled entries, H Status. Summary J4:K15: Settings (K4), Unbilled hours (K5), Unbilled $ (K6), Unbilled entries (K7), Oldest unbilled entry (K8), Older than B2 days (K9), Invoice # cells with only spaces (K10), Invoice # cells with a non-breaking space (K11), Unbilled $ in the whole Time log (K12), Client rows showing Check (K13), Time log CHECK (K14), CHECK (K15).

Key formulas (exact text in the file)
- Clients C5 `=IF(H5<>"OK","",SUMPRODUCT(('Time log'!$B$5:$B$1004=A5)*(TRIM('Time log'!$D$5:$D$1004)="")*('Time log'!$G$5:$G$1004="OK"),'Time log'!$C$5:$C$1004))`; D5 the same with `'Time log'!$E$5:$E$1004`; G5 `=IF(H5<>"OK","",SUMPRODUCT(('Time log'!$B$5:$B$1004=A5)*(TRIM('Time log'!$D$5:$D$1004)="")*('Time log'!$G$5:$G$1004="OK")))`
- E5 `=IF(OR(H5<>"OK",N(G5)=0),"",$B$1-SUMPRODUCT(MAX(('Time log'!$B$5:$B$1004=A5)*'Time log'!$H$5:$H$1004)))`; F5 `=IF(H5<>"OK","",SUMPRODUCT(('Time log'!$B$5:$B$1004=A5)*('Time log'!$H$5:$H$1004>$B$2)))`
- H5 `=IF(TRIM(A5)="","",IF($K$4<>"OK","Check: settings (B1:B2)",IF(COUNTIF($A$5:$A$54,A5)>1,"Check: client listed twice",IF(NOT(ISNUMBER(B5)),"Check: rate blank or text",IF(B5<0,"Check: rate below 0","OK")))))`
- K4 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),IF(AND(B1=INT(B1),B2>=0,B2=INT(B2)),"OK","Check"),"Check")`; K12 `=SUMPRODUCT((TRIM('Time log'!$D$5:$D$1004)="")*('Time log'!$G$5:$G$1004="OK"),'Time log'!$E$5:$E$1004)`; K15 `=IF(AND(K4="OK",K11=0,K13=0,K14="OK",ROUND(K12-K6,2)=0),"OK","Check")`
- Time log E5 `=IF(G5<>"OK","",ROUND(C5*INDEX(Clients!$B$5:$B$54,MATCH(B5,Clients!$A$5:$A$54,0)),2))`; H5 `=IF(AND(G5="OK",TRIM(D5)=""),Clients!$B$1-A5,0)`
- Time log F5 `=IF(LEN(D5)=0,"",IF(TRIM(D5)="","Yes: spaces only (not billed)",IF(TRIM(SUBSTITUTE(D5," "," "))="","Check: non-breaking space","")))` where the first quoted character in the SUBSTITUTE is a typed non-breaking space (U+00A0), not CHAR(160); quote it in the post as `SUBSTITUTE(D5,CHAR(160)," ")` only together with the LibreOffice caveat below.
- Time log G5 `=IF(AND(A5="",TRIM(B5)="",C5="",LEN(D5)=0),"",IF(Clients!$K$4<>"OK","Check: Clients settings",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(A5<>INT(A5),"Check: date has a time",IF(TRIM(B5)="","Check: no client",IF(COUNTIF(Clients!$A$5:$A$54,B5)=0,"Check: client not on the Clients tab",IF(NOT(ISNUMBER(INDEX(Clients!$B$5:$B$54,MATCH(B5,Clients!$A$5:$A$54,0)))),"Check: client's rate",IF(NOT(ISNUMBER(C5)),"Check: hours blank or text",IF(C5<0,"Check: hours below 0",IF(A5>Clients!$B$1,"After the as-of date (not counted)","OK"))))))))))`

Sample values (fictional; recalculated). 16 entries 09/14 to 11/24/2026; rates Pine & Co. Realty $55.00, Lakeside Yoga $60.00, Harbor Dental $75.00, Maple Street Bakery $50.00, Cedar Hill Florist $65.00. Invoice cells: 10/28 Lakeside holds one space (Looks blank? = Yes: spaces only (not billed)); 11/24 Lakeside holds `=""`.
| Client | Unbilled hours | Unbilled $ | Oldest unbilled | Older than 30 days | Unbilled entries |
|---|---|---|---|---|---|
| Pine & Co. Realty | 8.25 | $453.75 | 10/09/2026 | 1 | 3 |
| Lakeside Yoga | 6.00 | $360.00 | 10/28/2026 | 0 | 3 |
| Harbor Dental | 4.25 | $318.75 | 10/21/2026 | 1 | 2 |
| Maple Street Bakery | 2.25 | $112.50 | 11/05/2026 | 0 | 2 |
| Cedar Hill Florist | 0.00 | $0.00 | | 0 | 0 |
K4:K15: OK, 20.75, $1,245.00, 10, 10/09/2026, 2, 1, 0, $1,245.00, 0, OK, OK. Time log K4:K9: 16, 36.50, 0, 0, 0, OK. Unbilled ages: 49 (10/09), 37 (10/21), 30 (10/28, exactly 30, not counted as older), 24, 22, 15, 10, 8, 4, 3.

Edge cases tested (all matched the Python model)
- Entries exactly 30 days old (10/28/2026, not counted) and 31 days (counted); a client with no entries; every entry billed (oldest blank, no 0 date).
- Invoice cells: one space and two spaces (unbilled, Yes: spaces only), a non-breaking space alone, before or after a space (billed, Check: non-breaking space, K11 3, CHECK Check), `=""` (unbilled, no flag), "-", "N/A", a number 1046, "inv 7" (billed).
- Client names: "lakeside yoga" (matched, capital letters do not matter), "Lakeside Yoga " and " Harbor Dental" (Check: client not on the Clients tab); a client listed twice (both rows Check, CHECK Check); rate "55" as text or -55 (Check).
- Hours "2" as text, blank, -1.5 (Check); a date with a time, a text date, no date, an invoice number only (Check).
- As-of 10/15/2026 (before most entries: they show After the as-of date (not counted)); as-of text or with a time, B2 -1 or 2.5 (Check); B2 0; empty log; no clients; a row on row 1005; 1,000 rows; 50 clients with 300 entries; 200 random logs (1 to 12 clients, up to 250 entries with every invoice variant above).
- Trap (LibreOffice only, scratch copies of the sample): `=SUMIFS(C:C,B:B,"Lakeside Yoga",D:D,"")` typed on the Time log gave 4.00 (counted the empty cell and the `=""` cell, not the space); `=SUMIFS(C:C,B:B,"Lakeside Yoga",D:D,"=")` gave 3.00 (only the truly empty cell); the TRIM SUMPRODUCT gave 6.00. Mutation `('Time log'!$D$5:$D$1004="")` without TRIM also gave 4.00.
- CHAR(160) (LibreOffice only): with the LibreOffice process under LC_CTYPE=C.UTF-8, `=CODE(CHAR(160))` gave 239 and `=A1=CHAR(160)` was FALSE for a cell holding a non-breaking space; under the POSIX locale they gave 160 and TRUE. A typed non-breaking space inside SUBSTITUTE worked under both, and TRIM left the non-breaking space in place (LEN 1). Not checked in Excel or Sheets.

Pitfalls the post must state
- An Invoice # cell counts as not billed when TRIM leaves it empty; anything else ("-", "N/A", a non-breaking space) counts as billed. Link remove-non-breaking-spaces-clean-substitute for the non-breaking space.
- The SUMIFS "" vs "=" numbers are LibreOffice results only; link the Microsoft and Google SUMIFS pages and do not state Excel or Sheets behavior.
- When to invoice and which tax year income belongs to are not covered. Sample clients and hours are fictional.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A SUMIFS with "" for blank invoice numbers comes up 2 hours short: one cell holds a space (4.00 vs 6.00 for Lakeside Yoga).
- Looks blank? says Check: non-breaking space and CHECK is red: the cell was pasted with a non-breaking space; clear it.
- An entry is missing from a client: its client name has a trailing or leading space (Check: client not on the Clients tab), or its date is after the as-of date.
- Both rows of a client say Check: client listed twice (capital letters do not matter in the duplicate test).

Hero: `render(compact_copy(t+"tidy-tabs-unbilled-hours-by-client.xlsx","Clients",15,"unb",labels={"A2":"Older than (days)"}),"Clients",name="unb")` (the five clients with unbilled hours, dollars, oldest dates, the orange older-than flags and the summary with 20.75 and $1,245.00; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): SUMPRODUCT https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e and https://support.google.com/docs/answer/3094294; TRIM https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9 and https://support.google.com/docs/answer/3094140; SUMIFS as in post 1; LEN https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb and https://support.google.com/docs/answer/3094081; MAX https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098 and https://support.google.com/docs/answer/3094013; SUBSTITUTE https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332 and https://support.google.com/docs/answer/3094215.

---
## 4. set-order-processing-time-from-90th-percentile-business-days-networkdays-percentile-excel-google-sheets
Post: posts/ready/2026-11-28-team-04-set-order-processing-time-from-90th-percentile-business-days-networkdays-percentile-excel-google-sheets.md
Title (hooks final): How to Calculate Order Processing Time in Excel and Google Sheets
Template: templates/tidy-tabs-processing-time-percentile.xlsx (tabs: Orders, Promise, How to use)

Layout
- Orders: A1:A2 notes. Header row 4, rows 5:504: A Order #, B Order date, C Ship date (blank = not shipped) (inputs), D Business days to ship, E Status. G4 Closed days (holidays), G5:G34 inputs, defined name `Holidays` = Orders!$G$5:$G$34. Summary I4:J9: Shipped orders counted, Not shipped yet, Outside the date range, Rows showing Check, Typed below row 504, CHECK.
- Promise (one page): inputs B1 Percentile to promise (90%), B2 Count orders from (optional), B3 Count orders through (optional). Header row 5; B6 Shipped orders counted, B7 Not shipped yet, B8 Average, B9 Median, B10 Fastest, B11 Slowest, B12 Percentile at B1, B13 Processing time to promise (rounded up, shown as "5 business days"), B14 Orders shipped within the promise, B15 Share, B16 95th percentile. Coverage table D5:F16 (D 0 to 10 business days, E Orders shipped within, F Share; the promise row is green), D17 note. Checks header row 18: B19 Settings, B20 Holidays, B21 Orders rows showing Check, B22 Orders CHECK, B23 CHECK.

Key formulas (exact text in the file)
- Orders D5 `=IF(E5<>"OK","",MAX(0,NETWORKDAYS(INT(B5),INT(C5),Holidays)-1))`
- Orders E5 `=IF(AND(TRIM(A5)="",B5="",C5=""),"",IF(OR(Promise!$B$19<>"OK",Promise!$B$20<>"OK"),"Check: Promise settings or holidays",IF(NOT(ISNUMBER(B5)),"Check: order date blank or text",IF(C5="","Not shipped yet",IF(NOT(ISNUMBER(C5)),"Check: ship date is text",IF(INT(C5)<INT(B5),"Check: ship date before order date",IF(OR(AND(Promise!$B$2<>"",B5<Promise!$B$2),AND(Promise!$B$3<>"",B5>=Promise!$B$3+1)),"Outside the date range (not counted)","OK")))))))`
- Promise B6 `=IF(OR($B$19<>"OK",$B$20<>"OK"),"",COUNT(Orders!$D$5:$D$504))`; B8 `=IF(OR($B$19<>"OK",$B$20<>"OK"),"",IF(B6=0,"No shipped orders",ROUND(AVERAGE(Orders!$D$5:$D$504),2)))`; B9 the same with MEDIAN
- B12 `=IF(OR($B$19<>"OK",$B$20<>"OK"),"",IF(B6=0,"No shipped orders",PERCENTILE(Orders!$D$5:$D$504,B1)))`; B13 `=IF(OR($B$19<>"OK",$B$20<>"OK"),"",IF(B6=0,"No shipped orders",ROUNDUP(PERCENTILE(Orders!$D$5:$D$504,B1),0)))`
- B14 `=IF(NOT(ISNUMBER(B13)),"",COUNTIF(Orders!$D$5:$D$504,"<="&B13))`; B15 `=IF(NOT(ISNUMBER(B13)),"",B14/B6)`; B16 PERCENTILE with 0.95
- E6 `=IF(OR($B$19<>"OK",$B$20<>"OK",N($B$6)=0),"",COUNTIF(Orders!$D$5:$D$504,"<="&D6))`; F6 `=IF(E6="","",E6/$B$6)`
- B19 `=IF(NOT(ISNUMBER(B1)),"Check",IF(OR(B1<0,B1>1),"Check",IF(AND(B2<>"",NOT(ISNUMBER(B2))),"Check",IF(AND(B3<>"",NOT(ISNUMBER(B3))),"Check",IF(AND(ISNUMBER(B2),ISNUMBER(B3)),IF(B3<B2,"Check","OK"),"OK")))))`; B20 `=IF(COUNTA(Orders!$G$5:$G$34)=COUNT(Orders!$G$5:$G$34),"OK","Check")`; B23 `=IF(AND(B19="OK",B20="OK",B22="OK",ISNUMBER(B13)),"OK","Check")`

Sample values (fictional; recalculated). 70 orders #3101 to #3170, order dates 10/02 to 11/24/2026, 2 not shipped (11/23 and 11/24), closed days 10/12, 11/11, 11/26 and 11/27/2026. Built so the 90th percentile falls between 4 and 5.
| Order | Order date | Ship date | Business days |
|---|---|---|---|
| #3101 | 10/02/2026 | 10/06/2026 | 2 |
| #3102 | 10/03/2026 (Sat) | 10/06/2026 | 1 |
| #3112 | 10/10/2026 (Sat) | 10/10/2026 | 0 |
| #3114 | 10/12/2026 (closed) | 10/16/2026 | 3 |
| #3149 | 11/05/2026 | 11/17/2026 | 7 (across 11/11) |
| #3155 | 11/11/2026 (closed) | 11/16/2026 | 2 |
| #3169 | 11/23/2026 | | Not shipped yet |
Promise B6:B16: 68, 2, 2.72, 3.0, 0, 7, 4.3, 5 business days, 65, 95.59%, 5.0. Coverage: 0 days 4 (5.88%), 1 16 (23.53%), 2 32 (47.06%), 3 46 (67.65%), 4 61 (89.71%), 5 65 (95.59%, green), 6 67 (98.53%), 7 to 10 68 (100.00%). B19:B23 OK, OK, 0, OK, OK.

Edge cases tested (all matched the Python model)
- Holidays 11/26 and 11/27/2026: Saturday 11/21 to Monday 11/23/2026 0; Friday 11/20 same day 0; Sunday 11/22 same day 0 (NETWORKDAYS 0, minus 1, MAX 0); Friday 11/20 to Monday 11/23 1; Thanksgiving 11/26 to Monday 11/30 0; Wednesday 11/25 to Monday 11/30 1; 11/19 to 12/02/2026 7; 11/23 10:15 PM to 11/24 8:05 AM 1 (INT drops the time); ship 11/23 before order 11/24 (Check); text order or ship date (Check); no order number (counted); no order date (Check). The same rows with no holidays.
- One shipped order; no shipped orders ("No shipped orders" in B8:B13 and B16, B14:B15 blank, CHECK Check); no orders.
- B1 90 (Check, stats blank), 0 (fastest), 1 (slowest), 0.5, 0.85, 0.95, 0.99, "90%" as text, blank, -0.1 (Check).
- Date ranges: October only, from 11/01, through 10/15, reversed (Check), text (Check); a closed day typed as text (B20 Check, every order Check: Promise settings or holidays); no closed days; a row on row 505; 500 orders; 200 random sets (years 2020 to 2030, 0 to 300 orders with 0 to 15 days, unshipped, ship before order, date-times, text dates, random percentiles and ranges).
- Mutation: NETWORKDAYS without the -1 raised the sample average to 3.71.

Pitfalls the post must state
- Processing time here is the reader's handling time (order to shipment, business days, their own closed days); carrier transit is not included. No claim about Etsy, Shopify or carrier rules; the reader sets the number in their own shop. Past orders do not guarantee December. Sample orders are fictional.
- PERCENTILE is the inclusive kind (same as PERCENTILE.INC; link both Microsoft pages), not PERCENTILE.EXC. Blank cells in D (unshipped, not OK, outside the range) are skipped.
- Type the percentile as 90% or 0.9, not 90.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Every stat is blank and CHECK is red: B1 typed as 90, or a closed day typed as text.
- A weekend order shipped Monday shows 0, not 2: weekends are skipped and the -1 makes the next business day 1 only for weekday orders (Saturday to Monday is 0).
- The average (2.72) is shorter than what to promise: 3 days covered only 67.65% of orders.
- "No shipped orders": every ship date is blank or outside the date range.

Hero: `render(short_sheet_copy(t+"tidy-tabs-processing-time-percentile.xlsx","Promise",23,"pt"),"Promise",name="pt")` (stats with the green "5 business days", the coverage table with the green 5-day row, and the checks; legible).

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): NETWORKDAYS https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7 and https://support.google.com/docs/answer/3092979; PERCENTILE https://support.microsoft.com/en-us/office/percentile-function-91b43a53-543c-4708-93de-d626debdddca and PERCENTILE.INC https://support.microsoft.com/en-us/office/percentile-inc-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed, Google https://support.google.com/docs/answer/3094093; ROUNDUP https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7 and https://support.google.com/docs/answer/3093443; COUNTIF https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34 and https://support.google.com/docs/answer/3093480.

---
## 5. find-recurring-charges-and-subscriptions-in-bank-export-distinct-months-countifs-sumif-excel-google-sheets
Post: posts/ready/2026-11-28-team-05-find-recurring-charges-and-subscriptions-in-bank-export-distinct-months-countifs-sumif-excel-google-sheets.md
Title (hooks final): How to Find Recurring Charges and Subscriptions in a Bank Export in Excel and Google Sheets
Template: templates/tidy-tabs-recurring-charges-finder.xlsx (tabs: Charges, Recurring, How to use)

Layout
- Charges: A1:A2 notes. Header row 4, rows 5:2004: A Date, B Merchant, C Amount (inputs, paste), D Key (for COUNTIF), E Month, F First in month, G New merchant #, H Status, I Date (OK rows) (0 shown blank). Summary K4:L10: Rows OK, Total of OK rows, Refund rows, Rows showing Check, Typed below row 2004, Merchants found, CHECK.
- Recurring: A1 Recurring if months at least / C1 (3), A2 Export ends (latest date) / C2 (formula), A3 Not seen in (days) / C3 (45) (C1 and C3 inputs), D1:D3 notes. Header row 6, rows 7:106 (100 merchants): A #, B Merchant, C Key, D Charges, E Months, F Total, G First date, H Last date, I First amount, J Last amount, K Type, L Price change, M Stopped?, N Yearly at last price (x 12). Summary P6:Q18: Merchants (Q6), Subscriptions (Q7), Repeat vendors (Q8), Occasional (Q9), One-time (Q10), Subscriptions not seen lately (Q11), with a price change (Q12), Yearly at last prices (Q13), Spent on subscriptions (Q14), Settings (Q15), Merchants not listed (Q16), Charges CHECK (Q17), CHECK (Q18).

Key formulas (exact text in the file)
- Charges D5 `=IF(H5<>"OK","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B5,"~","~~"),"*","~*"),"?","~?"))`; E5 `=IF(H5<>"OK","",DATE(YEAR(A5),MONTH(A5),1))`
- F5 `=IF(H5<>"OK","",IF(COUNTIFS($B$5:B5,D5,$E$5:E5,E5)=1,1,0))`; G5 `=IF(H5<>"OK","",IF(COUNTIFS($B$5:B5,D5,$H$5:H5,"OK")=1,MAX($G$4:G4)+1,""))`; I5 `=IF(H5="OK",A5,0)`
- H5 `=IF(AND(A5="",TRIM(B5)="",C5=""),"",IF(NOT(ISNUMBER(A5)),"Check: date blank or text",IF(TRIM(B5)="","Check: no merchant",IF(B5&""<>TRIM(B5),"Check: extra spaces in the merchant",IF(NOT(ISNUMBER(C5)),"Check: amount blank or text",IF(C5<0,"Refund (not counted)","OK"))))))`
- Recurring C2 `=IF(COUNT(Charges!$A$5:$A$2004)=0,"",MAX(Charges!$A$5:$A$2004))`; B7 `=IF(A7>MAX(Charges!$G$5:$G$2004),"",INDEX(Charges!$B$5:$B$2004,MATCH(A7,Charges!$G$5:$G$2004,0)))`; C7 `=IF(B7="","",INDEX(Charges!$D$5:$D$2004,MATCH(A7,Charges!$G$5:$G$2004,0)))`
- D7 `=IF(B7="","",COUNTIFS(Charges!$B$5:$B$2004,C7,Charges!$H$5:$H$2004,"OK"))`; E7 `=IF(B7="","",SUMIF(Charges!$B$5:$B$2004,C7,Charges!$F$5:$F$2004))`; F7 `=IF(B7="","",ROUND(SUMIFS(Charges!$C$5:$C$2004,Charges!$B$5:$B$2004,C7,Charges!$H$5:$H$2004,"OK"),2))`
- G7 `=IF(B7="","",$C$2-SUMPRODUCT(MAX((Charges!$B$5:$B$2004=B7)*(Charges!$I$5:$I$2004>0)*($C$2-Charges!$I$5:$I$2004))))`; H7 `=IF(B7="","",SUMPRODUCT(MAX((Charges!$B$5:$B$2004=B7)*Charges!$I$5:$I$2004)))`
- I7 `=IF(B7="","",ROUND(SUMIFS(Charges!$C$5:$C$2004,Charges!$B$5:$B$2004,C7,Charges!$A$5:$A$2004,G7,Charges!$H$5:$H$2004,"OK"),2))`; J7 the same with H7
- K7 `=IF(OR(B7="",$Q$15<>"OK"),"",IF(E7>=$C$1,IF(COUNTIFS(Charges!$B$5:$B$2004,C7,Charges!$C$5:$C$2004,J7,Charges!$H$5:$H$2004,"OK")>=2,"Subscription","Repeat vendor"),IF(D7=1,"One-time","Occasional")))`
- L7 `=IF(K7<>"Subscription","",IF(ROUND(J7-I7,2)<>0,"Changed",""))`; M7 `=IF(K7<>"Subscription","",IF($C$2-H7>=$C$3,"Not seen in "&$C$3&"+ days",""))`; N7 `=IF(AND(K7="Subscription",M7=""),ROUND(J7*12,2),"")`
- Q15 `=IF(AND(ISNUMBER(C1),ISNUMBER(C3)),IF(AND(C1>=1,C1=INT(C1),C3>=1),"OK","Check"),"Check")`; Q18 `=IF(AND(Q15="OK",Q16=0,Q17="OK"),"OK","Check")`

Sample values (fictional card export; recalculated). 58 charges 12/03/2025 to 11/21/2026, 8 merchants; Cloudbox Storage typed in two cases.
| # | Merchant | Charges | Months | Total | First | Last | First amt | Last amt | Type | Price | Stopped? | Yearly |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Cloudbox Storage | 12 | 12 | $119.88 | 12/03/2025 | 11/03/2026 | $9.99 | $9.99 | Subscription | | | $119.88 |
| 2 | MEALKIT CO | 9 | 9 | $711.00 | 12/08/2025 | 08/08/2026 | $79.00 | $79.00 | Subscription | | Not seen in 45+ days | |
| 3 | PIXELPRO SUITE | 12 | 12 | $306.00 | 12/15/2025 | 11/15/2026 | $24.00 | $27.00 | Subscription | Changed | | $324.00 |
| 4 | STREAMLY | 12 | 12 | $185.88 | 12/21/2025 | 11/21/2026 | $15.49 | $15.49 | Subscription | | | $185.88 |
| 5 | OFFICE MART | 5 | 4 | $229.30 | 01/12/2026 | 10/05/2026 | $63.18 | $23.90 | Repeat vendor | | | |
| 6 | SQ *BAKERY | 3 | 2 | $57.25 | 02/06/2026 | 07/09/2026 | $18.50 | $16.00 | Occasional | | | |
| 7 | SITEHOST PRO | 1 | 1 | $120.00 | 03/10/2026 | 03/10/2026 | $120.00 | $120.00 | One-time | | | |
| 8 | SQ *MAIN ST BAKERY | 4 | 4 | $123.55 | 06/04/2026 | 10/22/2026 | $31.20 | $29.85 | Repeat vendor | | | |
C2 11/21/2026. Q6:Q18: 8, 4, 2, 1, 1, 1, 1, $629.76, $1,322.76, OK, 0, OK, OK. Charges L4:L10: 58, $1,852.86, 0, 0, 0, 8, OK. Keys: "SQ ~*BAKERY", "SQ ~*MAIN ST BAKERY".

Edge cases tested (all matched the Python model)
- The export reversed and shuffled (same merchants, numbered by first row).
- Names with wildcards over 4 months: "A*B" vs "A*BC", "A?C" vs "ABC", "X~Y" vs "X~~Y" vs "XY" (each counted on its own).
- "STREAMLY " and " STREAMLY" and "SQ  *BAKERY" (Check: extra spaces in the merchant, not counted, CHECK Check); "streamly" and "Office Mart" (merged with the upper-case names).
- STREAMLY twice on 11/21/2026: last amount $30.98 (both added) and the type turns Repeat vendor; refunds (Refund (not counted)); MEALKIT CO twice in April (10 charges in 9 months, $790.00, still a Subscription); 13 months of data; a blank row between every charge; text amount, text date, no merchant, no amount, amount only, spaces only (Check).
- C1 2, 12, 13 (fewer subscriptions), "3" as text, 0, 2.5 (Check, Type blank); C3 blank (Check) and 30; empty export; a row on row 2005; 130 merchants (Q16 30, CHECK Check); 2,000 rows; 200 random exports (monthly plans with stops and price changes, one-off merchants, refunds, extra-space names, text amounts, unsorted).
- Trap (LibreOffice only, scratch copy of the sample): unescaped `=COUNTIF(Charges!$B$5:$B$2004,"SQ *BAKERY")` 7 and `=SUMIF(Charges!$B$5:$B$2004,"SQ *BAKERY",Charges!$F$5:$F$2004)` 6 (escaped 3 and 2); `=SUMPRODUCT(--(Charges!$B$5:$B$2004="SQ *BAKERY"))` 3 (no wildcards with =). With the unescaped key in the mutation run, SQ *BAKERY also showed a total of $180.80 instead of $57.25. COUNTIF "Cloudbox Storage" counted all 12 charges (both spellings).

Pitfalls the post must state
- "Recurring" means charged in at least C1 distinct months. A plan billed once a year shows as One-time; check those lines. Yearly x 12 assumes a monthly plan and is left blank for a stopped subscription.
- COUNTIF, COUNTIFS and SUMIF read `*`, `?` and `~` as wildcards (link Microsoft's wildcard page); the key escapes them. Capital letters do not matter; spaces and store numbers do. Merchant names may need cleaning first (link categorize-expenses-by-keyword-search-lookup); PayPal and app-store charges can hide the merchant.
- No advice on what to cancel or deduct. Sample merchants are fictional.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A merchant with an asterisk shows too many charges and months: the COUNTIF criterion was not escaped (7 charges in 6 months instead of 3 in 2).
- A merchant appears twice in the list or Check: extra spaces: a leading, trailing or double space in the name.
- A monthly subscription shows Repeat vendor: two charges on the last date were added together, or the last amount differs from every earlier one.
- Q16 above 0 and CHECK red: more than 100 merchants in the export.

Hero: `render(compact_copy(t+"tidy-tabs-recurring-charges-finder.xlsx","Recurring",18,"rc",hide_rows=[1,2,3,4,5]),"Recurring",hide_cols="C",name="rc")` (the 8 merchants with months, totals, first and last amounts, green Subscription cells, orange Changed and Not seen in 45+ days, the yearly column and the summary; legible). Hiding column C also hides the settings values in C1:C3, so rows 1 to 5 are hidden too.

Help links (HEAD 200 = Microsoft; Google HEAD 404, GET 200): COUNTIFS https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842 and https://support.google.com/docs/answer/3256550; SUMIF https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b and https://support.google.com/docs/answer/3093583; SUBSTITUTE as in post 3; MATCH https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a and https://support.google.com/docs/answer/3093378; INDEX as in post 1; COUNTIF as in post 4; wildcards https://support.microsoft.com/en-us/office/using-wildcard-characters-in-searches-ef94362e-9999-4350-ad74-4d2371110adb.

---
Not built or tested (so not for the posts): anything in Excel or Google Sheets; typing by hand in the GUI (all inputs were written with openpyxl and recalculated headless); a non-US language setting (except the CHAR(160) locale probe); date-times in the card export; quarterly or annual subscription cycles; carrier transit times, marketplace processing-time settings or holiday delivery dates; when to invoice, which tax year income belongs to, deductions; budgeting or sales forecasts beyond last year's shape; how Excel or Sheets treat SUMIFS with "" or "=" or CHAR(160).
