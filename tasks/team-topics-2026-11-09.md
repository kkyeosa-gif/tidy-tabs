# Team topics - 2026-11-09 queue (5 new post topics)

Posts dated 2026-11-09. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc, a download that gives a reason to visit, no tax, wage-law or legal judgment).

Duplicate check done against posts/ready, posts/published and tasks/topics-seed.md (grep for restart/reset, fiscal, banded/alternate/zebra, duration/"1h 30m", price change/compare price, auto number returned nothing relevant). Closest existing posts and how they differ:
- Topic 1 vs time-to-decimal-hours (clock times like 9:30 AM, MOD and *24), add-up-hours-over-24 (real time values) and round-billable-time (numbers already): this one parses TEXT typed as "2h 30m" with FIND/MID.
- Topic 2 vs quarterly-sales-totals-sumifs (calendar quarters): this one has a start-month setting so the quarter label moves with a non-January year; labeling only.
- Topic 3 vs running-balance-column, year-to-date-sales-total, trailing-12-month: the running total here restarts on the 1st of each month.
- Topic 4 vs highlight-duplicate-values, simple-gantt, number-items-within-each-order: formatting by group change, not by value or count.
- Topic 5 vs compare-two-lists-find-missing (only missing items), price-increase-planner (plans a new price), price-in-effect-on-order-date (history lookup): this one compares an old and a new price list and labels Raised / Lowered / Same / New / Removed.

Dropped before choosing: auto-number rows that skip blanks (too close to number-items-within-each-order), sum of top 3 values (too close to pareto and three-biggest-unpaid), sum every nth row (low business fit), sinking-fund monthly savings (reads as financial advice), delivery date range text (too close to ship-by and return-by posts).

Function families: text parsing FIND/MID/VALUE (1), YEAR/MONTH/MOD/INT with SUMIFS (2), expanding-range SUMIFS with DATE (3), helper counter plus conditional formatting formula (4), INDEX/MATCH across two sheets with IF status (5). No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays.

How to verify in LibreOffice (all five): build the .xlsx with openpyxl, recalculate headless (soffice --headless --convert-to xlsx / csv), read the values back and compare with an independent Python calculation; also type the edge cases in the Test lines. Topic 4 also: convert to PDF and PNG and look at the shading. Excel and Google Sheets are not run by us; the posts say LibreOffice Calc only unless someone actually tests there.

Ranked best first:

1. How to Convert Text Like 2h 30m Into Minutes in Excel and Google Sheets
   - Slug: convert-text-hours-minutes-1h-30m-to-minutes-find-mid-excel-google-sheets
   - Search phrase: "convert 1h 30m to minutes excel", "text hours and minutes to decimal google sheets", "2h 30m to hours formula", "time entries typed as text spreadsheet".
   - Template: tidy-tabs-text-duration-to-minutes.xlsx (Entries, How to use). Entries: B1 hourly rate 65.00 (fictional), headers in row 3, data rows 4 to 15 (12 rows, extend note to row 104 for real use). A Date | B Time as typed | C Cleaned | D Hours | E Minutes | F Total minutes | G Decimal hours | H Amount | I Check.
     - Typed samples in B4:B15: "2h 30m", "45m", "1h", "1h15m", "3H 5M", " 90m ", "2h 0m", "1.5h", (blank), "30 min", "n/a", "0h 20m".
     - C4 `=LOWER(SUBSTITUTE(TRIM(B4)," ",""))`; D4 `=IF(ISNUMBER(FIND("h",C4)),VALUE(LEFT(C4,FIND("h",C4)-1)),0)`; E4 `=IF(ISNUMBER(FIND("m",C4)),VALUE(MID(C4,IF(ISNUMBER(FIND("h",C4)),FIND("h",C4)+1,1),FIND("m",C4)-IF(ISNUMBER(FIND("h",C4)),FIND("h",C4)+1,1))),0)`; F4 `=IF(C4="","",ROUND(D4*60+E4,0))`; G4 `=IF(C4="","",F4/60)`; H4 `=IF(C4="","",ROUND(G4*$B$1,2))`; I4 `=IF(C4="","",IF(F4=0,"Check",""))`. Totals row: SUM of F, G, H.
     - Expected: 150, 45, 60, 75, 185, 90, 120, 90, blank, 30, 0 (Check), 20; total 865 minutes.
   - Test: capital H/M, no spaces, extra spaces, nonbreaking space (add SUBSTITUTE UNICHAR(160) if it fails and say so), "1.5h", "90m" with no hours, "h30m" (no hours number, shows error or Check, document), "2:30" (not supported, shows Check), text "5" with no unit, blank, 100 rows, compare all to a Python regex parser.
   - Who ranks: Microsoft time help, Exceljet convert time, Stack Overflow "convert text to time". Angle: freelancers who jot "1h15m" in a notes column; result is minutes and decimal hours, no time-format cells.

2. How to Get the Fiscal Year and Quarter From a Date in Excel and Google Sheets
   - Slug: fiscal-year-quarter-from-date-custom-start-month-excel-google-sheets
   - Search phrase: "fiscal quarter from date excel", "fiscal year starting july formula", "quarter label from date google sheets", "fiscal year not calendar year spreadsheet".
   - Template: tidy-tabs-fiscal-year-quarter.xlsx (Sales, Summary, How to use). Sales: B1 first month of your year = 7 (July, fictional setting), headers row 3, rows 4 to 21 (18 sales): A Date | B Amount | C Fiscal year | D Fiscal quarter | E Label. Dates spread 07/2026 to 06/2027 plus one 06/2026 (FY2026) and one 07/2027.
     - C4 `=YEAR(A4)+IF(AND($B$1>1,MONTH(A4)>=$B$1),1,0)`; D4 `=INT(MOD(MONTH(A4)-$B$1,12)/3)+1`; E4 `="FY"&C4&" Q"&D4`. Convention stated in post: FY is named for the calendar year it ends in (a choice, not a rule).
     - Summary: labels FY2026 Q4, FY2027 Q1 to Q4, FY2028 Q1; Total `=SUMIFS(Sales!$B$4:$B$21,Sales!$E$4:$E$21,A4)`; check cell `=SUM(Sales!B4:B21)-SUM(B4:B9)` must be 0.
     - Expected: 11/09/2026 FY2027 Q2; 08/15/2026 FY2027 Q1; 01/05/2027 FY2027 Q3; 06/30/2027 FY2027 Q4; 06/15/2026 FY2026 Q4.
   - Test: start month 1 (FY = calendar year, Q1 in Jan), start 7, 10, 12, 2; Dec 31 and Jan 1 boundaries; leap day 02/29/2028; blank date (guard with IF(A4="","",...)); text date; start month entered as 13 or 0 (add data validation 1 to 12); every day of 3 years against Python.
   - Who ranks: Exceljet fiscal quarter, Microsoft date help, Stack Overflow answers using CHOOSE/MONTH. Angle: one setting cell, no lookup table; says plainly it labels dates and does not decide anyone's tax year.

3. How to Make a Running Total That Restarts Each Month in Excel and Google Sheets
   - Slug: running-total-that-restarts-each-month-sumifs-excel-google-sheets
   - Search phrase: "running total reset each month excel", "month to date total column google sheets", "cumulative sum by month spreadsheet", "running total restart new month".
   - Template: tidy-tabs-monthly-running-total.xlsx (Expenses, How to use). B1 monthly spending limit 450.00 (fictional), headers row 3, data from row 4 sorted by date: A Date | B What | C Amount | D Month to date | E Left this month | F Month total.
     - Sample rows (8): 10/03/2026 Supplies 120.00; 10/12 Booth fee 85.00; 10/27 Shipping 64.50; 11/02 Supplies 210.00; 11/09 Packaging 47.25; 11/18 Ads 150.00; 11/30 Shipping 38.75; 12/04 Supplies 95.00.
     - D4 `=SUMIFS($C$4:C4,$A$4:A4,">="&DATE(YEAR(A4),MONTH(A4),1))`; E4 `=$B$1-D4`; F4 `=SUMIFS($C$4:$C$63,$A$4:$A$63,">="&DATE(YEAR(A4),MONTH(A4),1),$A$4:$A$63,"<"&EDATE(DATE(YEAR(A4),MONTH(A4),1),1))`.
     - Expected D: 120.00, 205.00, 269.50, 210.00, 257.25, 407.25, 446.00, 95.00.
   - Test: rows not sorted by date (expanding version miscounts earlier-dated rows below, show it and the sort step), first day of month, last day of month, year change Dec to Jan, blank rows in the middle, text amounts, refund as negative, 300 rows, compare with Python groupby cumulative sum.
   - Who ranks: Exceljet running total, Ablebits cumulative sum, Stack Overflow. Angle: month-to-date column for a spending limit, no pivot table, no TODAY.

4. How to Shade Alternating Groups of Rows by Order Number in Excel and Google Sheets
   - Slug: alternate-row-colors-by-group-order-number-conditional-formatting-excel-google-sheets
   - Search phrase: "alternate row color by group excel", "banded rows based on value change", "shade rows by order number google sheets", "zebra stripes by group conditional formatting".
   - Template: tidy-tabs-shade-groups.xlsx (Lines, How to use). Lines: headers row 3, data rows 4 to 27 (24 line items, orders 5001 to 5009 with 1 to 4 lines each, sorted by order): A Order | B Item | C Qty | D Group no.
     - D4 `1`; D5 `=IF(A5=A4,D4,D4+1)` filled to D27. Conditional formatting on A4:D27, custom formula `=MOD($D4,2)=0`, light gray fill (also give the helper-free rule `=MOD(SUMPRODUCT(--($A$4:$A4<>$A$3:$A3)),2)=0`, test it before recommending).
     - Plain row banding for contrast: `=MOD(ROW(),2)=0` (shows why it cuts through an order).
   - Test: LibreOffice recalc for D values (1,1,2,3,3,3,...), convert to PDF then PNG and view the shading, rule survives xlsx reopen (inspect conditionalFormatting XML), blank row in the data, unsorted orders (same order in two places gets two shades, say so), insert a row inside the range, 200 rows. Google Sheets rule steps from Google help, not claimed as tested.
   - Who ranks: Microsoft banded rows help, Exceljet conditional formatting formulas, Ablebits. Angle: shading follows the order, so a 3-line order reads as one block on a printout or packing list.

5. How to Compare Two Price Lists and See What Changed in Excel and Google Sheets
   - Slug: compare-old-and-new-price-lists-change-percent-index-match-excel-google-sheets
   - Search phrase: "compare two price lists excel", "find price changes between two lists", "old price vs new price spreadsheet", "vlookup compare prices google sheets".
   - Template: tidy-tabs-compare-price-lists.xlsx (Old, New, Compare, How to use). Old A1:B9 (SKU | Price), New A1:B9, 8 rows each.
     - Old: CND-8 18.00; LIP-1 3.50; TOT-1 24.00; PIN-1 8.00; STK-4 4.50; BOX-1 32.00; CRD-1 5.00; MUG-1 16.00. New: CND-8 19.50; LIP-1 3.50; TOT-1 24.00; PIN-1 7.50; STK-4 4.50; BOX-1 34.00; SOAP-1 9.00; MUG-1 16.00.
     - Compare (rows 4 to 11 follow New): A SKU `=New!A2`; B New price `=New!B2`; C Old price `=IFERROR(INDEX(Old!$B$2:$B$9,MATCH(A4,Old!$A$2:$A$9,0)),"")`; D Change `=IF(C4="","",ROUND(B4-C4,2))`; E Percent `=IF(OR(C4="",C4=0),"",D4/C4)`; F Status `=IF(C4="","New",IF(D4>0,"Raised",IF(D4<0,"Lowered","Same")))`. Removed block (rows 14 to 21 follow Old): `=IF(COUNTIF(New!$A$2:$A$9,Old!A2)=0,"Removed","")`. Counts with COUNTIF.
     - Expected: Raised 2 (CND-8 +1.50, 8.3%; BOX-1 +2.00, 6.3%), Lowered 1 (PIN-1 -0.50, -6.3%), Same 4, New 1 (SOAP-1), Removed 1 (CRD-1).
   - Test: SKU case differs (MATCH ignores case), trailing space in a SKU (TRIM both sides, document), SKU stored as number vs text, duplicate SKU in Old (first match wins, add a check), old price 0 (percent blank), price typed as text, 200 rows, compare to a Python dict diff.
   - Who ranks: Microsoft VLOOKUP help, Exceljet compare two lists, Ablebits. Angle: five-way status instead of a plain match; includes the items that disappeared; made-up prices, no advice on what to charge.
