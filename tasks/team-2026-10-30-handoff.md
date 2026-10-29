# Handoff 2026-10-30: five templates built and checked (LibreOffice Calc 24.2.7.2 only)

Status: topics (tasks/team-topics-2026-10-30.md) done; the posts are the next step. The templates were built by five new functions in `scripts/build-templates.py` (`billable_time_rounding`, `all_orders_for_customer`, `overnight_shift_hours`, `lapsed_customer_flag`, `year_over_year_sales`, added after `customer_tiers` and registered at the end of the main block). Only these five were run (a small runner that imports the script and calls just them, with `PYTHONDONTWRITEBYTECODE=1`, using `/usr/bin/python3` because the default `python3` on this machine has no openpyxl); older .xlsx were not rewritten. Nothing was committed. New in `git status`: the five .xlsx, the edited script, tasks/briefs.md, this file.

Exact version string used for `tested_in`: **LibreOffice 24.2.7.2 420(Build:2)** (Calc, Linux, headless). Write "LibreOffice Calc 24.2.7.2". **Nothing was opened in Excel or Google Sheets and nothing may be claimed about them** (including whether MAXIFS, SUMIFS with EDATE, IFERROR inside SUMPRODUCT, TEXT(...,"mmmm"), [h]:mm negative display, or the dropdown warning behave the same there).

Method: edited copies made with openpyxl (inputs only), recalculated with `soffice --headless --convert-to xlsx` into a scratch dir, read back with `data_only=True`, compared with independent Python models that read only the typed inputs (Fraction and Decimal ROUND_HALF_UP for minutes and money, datetime arithmetic with day rollover, plain filters and grouped sums). 111 workbooks (5 committed samples recalculated + 106 edited copies: 20 + 27 + 14 + 27 + 23 per template including each sample), 6,810 cell comparisons. Final run: 0 mismatches and no error value (#NAME?, #VALUE!, #REF!, #N/A, #DIV/0!, #NUM!, Err:) in any workbook, in any of the five templates, including the invalid-input copies. Along the way the tests found 5 bugs, all fixed in the script and re-run (listed per template below). No intentional error values are shown in any sample (messages are text: Check times, Review ..., CHECK). Formulas use only: IF, IFERROR, AND, OR, NOT, ISNUMBER, COUNT, COUNTA, COUNTIF, COUNTIFS, SUM, SUMIF, SUMIFS, SUMPRODUCT, MAXIFS (written `_xlfn.MAXIFS`), MAX, MIN, ABS, ROUND, ROUNDUP, MOD, INT, N, INDEX, MATCH, ROWS, DATE, EDATE, TEXT (no TODAY, NOW, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, dynamic arrays, WEBSERVICE; one IFERROR inside a SUMPRODUCT in the billable CHECK, no array entry). Sample data fictional, dollars US. Yellow (FFF2CC) = inputs, gray (EDEDED) = formulas. Each file has a "How to use" sheet and one CHECK cell. PDF export (soffice headless, one PDF per file): every page in all five files is US Letter (792 x 612 pt landscape or 612 x 792 portrait, as built per tab; no other size appears).

Titles and slugs are in tasks/team-topics-2026-10-30.md. Posts must say: LibreOffice Calc 24.2.7.2 only, sample data fictional, no advice (no contract or billing-practice advice, no overtime or pay law, no marketing or consent advice).

Deviations from the topics file sketches (each is explained in its section): billable log has 60 slots and an Extra minutes column; orders list uses rows 5 to 104 (100 slots, not 40); overnight rate sits on the Shifts tab (B1), not on a Settings tab; the lapsed sheet uses COUNTIFS with ">0" instead of COUNTIF in the last-order formula; year-over-year gets a latest-order cell and a "same months only" row.

---

## 1. Round billable time up to blocks (`templates/tidy-tabs-billable-time-rounding.xlsx`, function `billable_time_rounding`)
- Sheets: Time log (landscape, print area A1:M64, header row 4 repeats), Settings, How to use.
- Inputs (yellow): Settings B1 block size in minutes (15; whole number 1 to 480, dropdown-free, data validation), B2 hourly rate ($85.00). Time log A5:E64: Date, Client, Task, Start, End (60 slots, 15 filled with fictional bookkeeper entries). Settings B3 = settings check (OK or a message).
- Columns: F minutes worked, G billed minutes, H billed hours, I amount, J extra minutes billed. Summary (labels L, values M): M5 entries, M6 actual minutes, M7 billed minutes, M8 extra minutes, M9 amount, M10 rows showing Check times, M11 CHECK.
- Formulas to quote (row 5):
  - F5 `=IF(AND(D5="",E5=""),"",IF(AND(ISNUMBER(D5),ISNUMBER(E5)),IF(E5>=D5,ROUND((E5-D5)*1440,0),"Check times"),"Check times"))`; core `=ROUND((E5-D5)*1440,0)`
  - G5 `=IF(ISNUMBER(F5),IF(Settings!$B$3<>"OK","Check settings",ROUNDUP(F5/Settings!$B$1,0)*Settings!$B$1),"")`; core `=ROUNDUP(F5/Settings!$B$1,0)*Settings!$B$1`
  - H5 `=IF(ISNUMBER(G5),G5/60,"")`; I5 `=IF(ISNUMBER(H5),ROUND(H5*Settings!$B$2,2),"")`; J5 `=IF(ISNUMBER(G5),G5-F5,"")`
  - Settings B3 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),IF(AND(B1>=1,B1<=480,B1=INT(B1),B2>=0),"OK","Block: whole minutes from 1 to 480. Rate: 0 or more"),"Type a number in each yellow cell")`
  - M11 CHECK `=IF(AND(Settings!B3="OK",COUNT(G5:G64)=COUNT(F5:F64),SUMPRODUCT(--ISNUMBER(G5:G64),--(IFERROR(MOD(G5:G64,Settings!B1),0)<>0))=0,COUNTIF(J5:J64,"<0")=0,COUNTIF(J5:J64,">="&Settings!B1)=0,ROUND(M8-SUM(J5:J64),6)=0,COUNTA(A65:E500)=0),"OK","CHECK")`
- Errors shown on purpose: none (a start after the end shows the text "Check times", not an error).
- Example table (sample, 15 min block, $85.00; every cell equals the model):
  | Start | End | Minutes | Billed | Hours | Amount | Extra |
  | 9:02 AM | 9:49 AM | 47 | 60 | 1.00 | $85.00 | 13 |
  | 10:00 AM | 10:45 AM | 45 | 45 | 0.75 | $63.75 | 0 |
  | 10:00 AM | 10:46 AM | 46 | 60 | 1.00 | $85.00 | 14 |
  | 1:15 PM | 1:15 PM | 0 | 0 | 0.00 | $0.00 | 0 |
  | 8:30 AM | 9:31 AM | 61 | 75 | 1.25 | $106.25 | 14 |
  | 2:00 PM | 2:01 PM | 1 | 15 | 0.25 | $21.25 | 14 |
  Summary: 15 entries, 799 actual minutes, 885 billed, 86 extra minutes, $1,253.75, M11 OK.
- Edge findings (20 workbooks, 907 cell comparisons, 0 mismatches):
  - 15 stays 15, 30 stays 30, 60 stays 60, 0 stays 0, 1 becomes 15, 14 becomes 15, 16 becomes 30, 0:00 to 23:59 (1439 minutes) becomes 1440. Block 6: 6 stays 6, 7 becomes 12, 5 becomes 6, 59 becomes 60. Block 30: 29 and 30 become 30, 31 becomes 60. Block 1: billed equals worked. Block 60: 1 becomes 60, 61 becomes 120. Rate 0: amounts $0.00, rate $77.77: matches Decimal ROUND_HALF_UP.
  - Seconds: the minutes are rounded to the nearest whole minute first (29 s is 0 minutes, 30 s is 1, 14 min 29 s is 14, 14 min 31 s is 15, 15 min 30 s is 16), then rounded up to the block. Say so; a client agreement may count it another way and the post must not advise.
  - Start after end, a Start without an End, an End without a Start, and text times ("9:00" typed as text, "ten"): "Check times", no billed minutes or amount, row left out of every total; M10 counts them; M11 stays OK because the row is visible, not hidden.
  - Settings block 0, 2.5, "fifteen", 481 or rate -5: Settings B3 shows the message, G shows "Check settings", H, I and J blank, M11 CHECK. No error values.
  - 60 filled rows (15-minute and 7-minute blocks, random times): all match. A date typed in A65 (below the table): M11 CHECK.
- Limits: a shift that crosses midnight is not handled (shows Check times). Not run: a block of exactly 480, more than 60 rows, time zones or daylight-saving changes, a time with a date part. Says nothing about what a client contract allows. Excel and Sheets not tested; the topics file asks to link the Microsoft ROUNDUP and CEILING help pages (not opened).

---

## 2. List every order for one customer (`templates/tidy-tabs-all-orders-for-customer.xlsx`, function `all_orders_for_customer`)
- Sheets: Orders (landscape, print area A1:F44, header row 4 repeats), Customer view (portrait, print area A1:D16), How to use.
- Inputs (yellow): Orders A5:E104 (Order #, Customer, Order date, Item, Amount; 100 slots, 40 filled, 9 fictional customers: counts 12, 7, 6, 5, 4, 3, 1, 1, 1). Customer view B1 = customer (sample "Alma Reyes"; dropdown from the yellow pick-list H5:H24, 10 names; the dropdown is a warning type so a typed name is allowed). The topics file said 40 rows (A5:A44); the build has 100 slots and ranges to row 104.
- Orders column F is a gray helper (match number). Rows of the picked customer turn green.
- Customer view: B2 orders found, B3 total, D1 rows listed, D2 CHECK, D3 note; results table header row 4, rows 5 to 16 (12 rows) with Order #, Order date, Item, Amount.
- Formulas to quote:
  - Orders F5 `=IF(AND('Customer view'!$B$1<>"",B5='Customer view'!$B$1),COUNTIF($B$5:B5,'Customer view'!$B$1),"")` (the topics file's version without the blank guard would match every blank row when B1 is empty)
  - Customer view A5 `=IFERROR(INDEX(Orders!$A$5:$A$104,MATCH(ROWS($A$4:A4),Orders!$F$5:$F$104,0)),"")`; B5, C5, D5 are the same with Orders!$C, $D, $E
  - B2 `=IF(B1="",0,COUNTIF(Orders!$B$5:$B$104,B1))`; B3 `=IF(B1="",0,SUMIF(Orders!$B$5:$B$104,B1,Orders!$E$5:$E$104))`; D1 `=SUMPRODUCT(--(A5:A16<>""))`
  - D2 CHECK `=IF(B1="","Pick a customer",IF(AND(D1=B2,ROUND(SUM(D5:D16)-B3,2)=0),"OK","CHECK"))`
  - D3 `=IF(B1="","",IF(B2=0,"No orders found for this name",IF(B2>12,"Only the first 12 are listed","")))`
- Errors shown on purpose: none. MATCH gives #N/A below the last match inside the formula and IFERROR turns it into a blank; no #N/A is visible.
- Example (sample pick Alma Reyes): B2 12, B3 $258.75, D1 12, D2 OK. First rows: 5010 07/16/2026 Cedar candle $36.00; 5011 07/19/2026 Brass hoops $24.00; 5012 07/20/2026 Cedar candle $54.00; 5013 07/23/2026 Birthday card $6.00. Other picks: Ben Tran 7 orders $178.00, Cora Lind 6 $68.25, Dev Kapoor 5 $143.25, Elena Marsh 4 $75.75, Felix Ode 3 $122.00, Grace Hale 1 $6.50, Hugo Park 1 $36.00, Iris Wong 1 $54.00, Jonah Bell (in the pick-list, no orders) 0.
- Edge findings (27 workbooks, 1,326 cell comparisons, 0 mismatches against a Python filter):
  - Customer with 1 order, 0 orders (D3 "No orders found for this name", D2 OK), name not in the data, blank B1 (D2 "Pick a customer", nothing listed, no error values).
  - Exactly 12 orders (the sample): all listed, D2 OK. 11 orders: 11 listed. 13 and 14 orders: 12 listed, B2 13 or 14, D2 CHECK, D3 "Only the first 12 are listed". 60 orders for one name: 12 listed, CHECK.
  - Capitalization does not matter ("alma reyes", "ALMA REYES" find the same 12). A trailing space in B1 ("Alma Reyes ") finds 0 orders (not an error, no CHECK, D3 says none found). A trailing space inside one Orders cell drops that order (11 found).
  - A blank Amount or Item cell in a matching row is shown as 0 (INDEX of an empty cell); the total is still right. A text Order # (A-77) is listed as text. Negative amount: listed and totaled. A blank row in the middle of the data: skipped.
  - Wildcards: typing `Al*` or `Cora Lin?` in B1 makes COUNTIF find the orders (12, 6) but the exact-match helper does not, so nothing is listed and D2 shows CHECK. The CHECK catches it.
- Limits: 12 result rows (extend by copying the last row of results down and, if needed, the ROWS reference). Cell ranges end at row 104. FILTER does this in one cell in current Excel 365 and Google Sheets; FILTER was not tested here and the template does not use it. Not run: more than 100 orders, a name with a leading space in the data, very long names.

---

## 3. Hours worked past midnight with an unpaid break (`templates/tidy-tabs-overnight-shift-hours.xlsx`, function `overnight_shift_hours`)
- Sheets: Shifts (landscape, print area A1:M64, header row 4 repeats), How to use. The topics file said "Settings!B1" for the rate; the build puts the yellow rate on Shifts B1 ($18.00) with a rate check in B2, so there is no Settings sheet.
- Inputs (yellow): B1 hourly rate; Shifts A5:E64 (Date, Name, Start, End, Break minutes; 60 slots, 14 filled with fictional night-market and event crew). A blank break counts as 0.
- Columns: F hours worked ([h]:mm), G decimal hours, H pay, I plain end minus start (the problem, [h]:mm), J Review. Summary (labels L, values M): M5 shifts with hours, M6 total decimal hours, M7 total pay, M8 rows to review, M9 rows where the plain subtraction is negative, M10 CHECK.
- Formulas to quote (row 5):
  - Core (as in the topics file): `=MOD(D5-C5,1)-E5/1440`
  - F5 `=IF(J5="OK",MAX(0,MOD(D5-C5,1)-N(E5)/1440),"")` (N() makes a blank break 0; MAX(0,...) removes a tiny negative float tail when break equals the shift)
  - G5 `=IF(ISNUMBER(F5),ROUND(F5*24,2),"")`; H5 `=IF(ISNUMBER(G5),IF($B$2="OK",ROUND(G5*$B$1,2),"Check rate"),"")`; I5 `=IF(AND(ISNUMBER(C5),ISNUMBER(D5)),D5-C5,"")`
  - J5 `=IF(AND(C5="",D5=""),"",IF(NOT(AND(ISNUMBER(C5),ISNUMBER(D5),OR(E5="",ISNUMBER(E5)))),"Review: start, end or break is not a number",IF(MOD(D5-C5,1)=0,"Review: start equals end (0 or 24 hours?)",IF(E5<0,"Review: break is negative",IF(ROUND(MOD(D5-C5,1)*1440-N(E5),4)<0,"Review: break is longer than the shift","OK")))))`
  - B2 `=IF(ISNUMBER(B1),IF(B1>=0,"OK","Rate must be 0 or more"),"Rate must be a number")`
  - M10 CHECK `=IF(AND(B2="OK",COUNT(F5:F64)=COUNTIF(J5:J64,"OK"),COUNT(F5:F64)+M8=SUMPRODUCT(--(J5:J64<>"")),ABS(SUM(F5:F64)*24-SUM(G5:G64))<=0.005*COUNT(F5:F64)+0.000001,COUNTA(A65:E500)=0),"OK","CHECK")`
- Errors shown on purpose: none. Column I shows the problem as negative times (LibreOffice shows -16:00 for 6:00 PM to 2:00 AM; the rows are not errors). Rows with a start equal to the end or a break longer than the shift show a red "Review: ..." text and get no hours or pay.
- Example table (sample, $18.00):
  | Start | End | Break | Hours (h:mm) | Decimal | Pay | Plain end - start | Review |
  | 6:00 PM | 2:00 AM | 30 | 7:30 | 7.50 | $135.00 | -16:00 | OK |
  | 5:30 PM | 11:30 PM | 30 | 5:30 | 5.50 | $99.00 | 6:00 | OK |
  | 8:00 PM | 12:00 AM | 0 | 4:00 | 4.00 | $72.00 | -20:00 | OK |
  | 10:00 PM | 10:00 PM | 30 | blank | blank | blank | 0:00 | Review: start equals end (0 or 24 hours?) |
  | 5:00 PM | 5:20 PM | 30 | blank | blank | blank | 0:20 | Review: break is longer than the shift |
  | 4:00 PM | 11:59 PM | 0 | 7:59 | 7.98 | $143.64 | 7:59 | OK |
  | 6:00 PM | 12:30 AM | (blank) | 6:30 | 6.50 | $117.00 | -17:30 | OK |
  Summary: 12 shifts with hours, 77.48 decimal hours, $1,394.64, 2 rows to review, 8 rows where the plain subtraction is negative, M10 OK.
- Edge findings (14 workbooks, 869 cell comparisons, 0 mismatches; model is datetime arithmetic with a day rollover, Decimal half-up for hours and pay):
  - End exactly at midnight works (8:00 PM to 12:00 AM is 4:00; 12:00 PM to 12:00 AM with a 60 minute break is 11:00; 12:00 AM to 8:00 AM works; 11:59 PM to 12:01 AM is 2 minutes). 12:00 AM start to 11:59 PM end is 23:59.
  - Break equals the shift (480 minutes on 6 PM to 2 AM, 30 minutes on 9:00 to 9:30): OK with 0:00 hours and $0.00. One minute longer (481, 31): Review. One minute shorter: 0.02 decimal hours. A 29.5 minute break works. Blank break counts as 0.
  - Start equals end (10 PM to 10 PM, 12 AM to 12 AM, and a date-time exactly 24 hours later): Review, no guess of 0 or 24 hours.
  - Seconds in the times: hours and pay match the model to the cent (rounding to 2 decimals half up).
  - Date-times (a date part on start and end): work the same (6 PM 10/09 to 2 AM 10/10 with 30 minutes is 7.5). End earlier than the start on the same date is treated as the next day (11 PM to 8 PM is 21:00): the sheet cannot know that was a typo.
  - Text in start, end or break, a blank end, a blank start, a negative break: Review, no hours, no pay, no error values.
  - Rate 0 gives $0.00 pay; rate -1 or text: B2 message, H shows "Check rate", M10 CHECK. A date typed in A65: CHECK.
  - Two random 60-row sets (breaks 0 to 600 minutes): 53 OK and 7 Review in one, 60 OK in the other; all matched, M9 matched the count of rows with end before start.
- Limits: assumes every shift is under 24 hours and one row is one shift. How LibreOffice shows a negative time (a minus sign, "-16:00") is its behavior only; the topics file says the #### display is app dependent and it was not tested in Excel or Sheets, so do not claim it. Not run: shifts of exactly 24 hours (that is start equals end, a Review row), more than 60 rows, daylight-saving nights. No overtime rules, pay law or what must be paid; the pay column is only hours times a typed rate.

---

## 4. Flag customers with no order in 90 days (`templates/tidy-tabs-lapsed-customer-flag.xlsx`, function `lapsed_customer_flag`)
- Sheets: Orders (portrait, print area A1:C41, header row 1 repeats), Customers (landscape, print area A1:H54, header row 4 repeats), Settings, How to use.
- Inputs (yellow): Settings B1 as-of date (typed 10/30/2026, not TODAY) and B2 days without an order (90); Orders A2:C200 (Order date, Customer, Amount; 40 filled, 15 customers); Customers A5:A54 names (15 filled, 50 slots). Settings B3 = settings check.
- Customers columns: B last order, C days since, D status, E lifetime spend. Lapsed rows are red, No orders rows gray. Summary (labels G, values H): H5 customers, H6 Active, H7 Lapsed, H8 No orders, H9 lifetime spend of the lapsed, H10 orders with a name not in the list, H11 orders dated after the as-of date, H12 CHECK.
- Formulas to quote (row 5):
  - B5 `=IF(A5="","",IF(COUNTIFS(Orders!$B$2:$B$200,A5,Orders!$A$2:$A$200,">0")=0,"No orders",MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A5)))` (the file stores `_xlfn.MAXIFS`; the topics file had `COUNTIF(...)=0`; the build uses COUNTIFS with ">0" because with COUNTIF a customer whose only orders have text dates got MAXIFS = 0, a date of 12/30/1899, and status Lapsed; that bug was found by a test and fixed)
  - C5 `=IF(AND(ISNUMBER(B5),ISNUMBER(Settings!$B$1)),Settings!$B$1-B5,"")` (the topics file had `=IF(ISNUMBER(B5),Settings!$B$1-B5,"")`; a text as-of date gave #VALUE!, fixed with the second ISNUMBER)
  - D5 `=IF(A5="","",IF(Settings!$B$3<>"OK","Check settings",IF(NOT(ISNUMBER(C5)),"No orders",IF(C5>Settings!$B$2,"Lapsed","Active"))))`
  - E5 `=IF(A5="","",SUMIF(Orders!$B$2:$B$200,A5,Orders!$C$2:$C$200))`
  - H10 `=SUMPRODUCT(--(Orders!$B$2:$B$200<>""),--(COUNTIF(A5:A54,Orders!$B$2:$B$200)=0))`; H11 `=COUNTIF(Orders!$A$2:$A$200,">"&Settings!B1)`
  - H12 CHECK `=IF(AND(Settings!B3="OK",H6+H7+H8=H5,H10=0,SUMPRODUCT((A5:A54<>"")*(COUNTIF(A5:A54,A5:A54)>1))=0,COUNTA(A55:A500)=0),"OK","CHECK")`
  - Settings B3 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),IF(AND(B2>=0,B2=INT(B2)),"OK","Days must be a whole number, 0 or more"),"Type a date and a number in the yellow cells")`
- Errors shown on purpose: none. (With a typed text as-of date the Customers tab shows "Check settings", not errors.)
- Example table (sample, as-of 10/30/2026, 90 days):
  | Customer | Last order | Days since | Status | Lifetime spend |
  | Alma Reyes | 10/27/2026 | 3 | Active | $317.80 |
  | Hugo Park | 08/02/2026 | 89 | Active | $459.81 |
  | Iris Wong | 08/01/2026 | 90 | Active | $250.23 |
  | Jonah Bell | 07/31/2026 | 91 | Lapsed | $272.20 |
  | Kira Stone | 07/30/2026 | 92 | Lapsed | $263.52 |
  | Nate Dunn | 12/24/2025 | 310 | Lapsed | $226.43 |
  | Opal Frost | No orders | blank | No orders | $0.00 |
  Summary: 15 customers, 9 Active, 5 Lapsed, 1 No orders, lifetime spend of the lapsed $1,387.91, 0 orders for unlisted names, 0 after the as-of date, H12 OK. Exactly 90 days (Iris Wong) stays Active; 91 is Lapsed.
- Edge findings (27 workbooks, 1,932 cell comparisons, 0 mismatches after the two fixes; model is Python max date and day count):
  - As-of 10/31/2026: Iris Wong becomes 91 days and Lapsed (8 Active, 6 Lapsed). As-of 11/29/2026: 6 Active, 8 Lapsed. As-of 01/01/2027: 3 Active, 11 Lapsed. As-of 02/29/2028 with 90 days: all 14 with orders are Lapsed.
  - Days cutoff 0 and 1 on the sample: 0 Active, 14 Lapsed (nobody ordered on the as-of date); cutoff 30: 3 Active. Cutoff 365 and 1000: all 14 Active (the oldest last order is 310 days back). With a cutoff of 90, an extra order dated exactly the as-of date makes that customer 0 days since and Active (Opal Frost, 10 Active and 0 No orders).
  - As-of 08/01/2026 (earlier than some orders): 10 orders are after the as-of date, H11 counts them, their customers get a negative days since and stay Active. One order dated 11/15/2026: H11 1, Dev Kapoor days since -16, Active.
  - Customer with no orders: "No orders", blank days, spend 0. No orders at all: 15 No orders. No customers listed: H5 0 and H10 40 (all orders unmatched), H12 CHECK.
  - Capitalization: "BEN TRAN" in the list matches "Ben Tran" orders, H12 OK. A trailing space in a list name ("Alma Reyes "): that customer shows No orders, H10 counts the 2 orders that no longer match a listed name, H12 CHECK. A name listed twice: both rows computed, H5 counts 16, H12 CHECK. An order for a name not listed: H10 1, H12 CHECK.
  - An order whose date is text ("March 3", "2026-09-01"): ignored for the last-order date, but its amount is still in the spend; a customer whose only orders have text dates shows No orders with a spend above 0 (say so). Blank amount: skipped; text amount "x": skipped by SUMIF.
  - Invalid settings (text as-of date, 2.5, -5 or "ninety" days): B3 message, all statuses "Check settings", no counts, H12 CHECK, no error values (after the fix).
  - 50 customers and 190 random orders: all match.
- Limits: MAXIFS needs Excel 2019 or newer (stated in the sheet); not tested in Excel. The topics file offered an older-Excel alternative only if tested: none was built or tested. Replacing the as-of date with `=TODAY()` was not run (the sheet explains the effect; do not claim results). Not run: more than 50 customers or 199 orders, trailing spaces in Orders names, a time part on the order date. Labels only; nothing about contacting anyone or consent.

---

## 5. Year-over-year sales by month (`templates/tidy-tabs-year-over-year-sales.xlsx`, function `year_over_year_sales`)
- Sheets: Orders (portrait, print area A1:B121, header row 1 repeats), Year over year (landscape, print area A1:F26, header row 3 repeats), How to use.
- Inputs (yellow): Orders A2:B500 (Order date, Amount; 120 fictional online-shop rows from 01/01/2025 to 09/29/2026, March 2025 has none, one refund of -$24.00 on 05/31/2026, first and last day of most months present); Year over year B1 this year (2026; whole number 1901 to 9998 with data validation). C1 = year check, B2 = last year, E1 = latest order date, E2 = CHECK.
- Table (header row 3, rows 4 to 15): A month name, B month start, C this year, D last year, E change dollars, F change percent. Row 16 totals, row 17 "last year, same months only". C19 best month, C20 weakest month, C21 orders without a real date, C22 orders in other years.
- Formulas to quote (row 4; the sheet reads whole columns A and B of Orders as the topics file did):
  - B4 `=IF($C$1<>"OK","",DATE($B$1,ROWS($A$4:A4),1))`; A4 `=IF(B4="","",TEXT(B4,"mmmm"))`
  - C4 `=IF(B4="","",IF(NOT(ISNUMBER($E$1)),"",IF(B4>$E$1,"",SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&B4,Orders!$A:$A,"<"&EDATE(B4,1)))))` (core: `=SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&B4,Orders!$A:$A,"<"&EDATE(B4,1))`)
  - D4 `=IF(B4="","",SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&EDATE(B4,-12),Orders!$A:$A,"<"&EDATE(B4,-11)))` (the topics file's last-year version used EDATE(B4,-12) for the start; the end is EDATE(B4,-11), the first of the next month a year earlier)
  - E4 `=IF(OR(C4="",D4=""),"",ROUND(C4-D4,2))`; F4 `=IF(OR(C4="",D4=""),"",IF(D4=0,"n/a",C4/D4-1))`
  - E1 `=IF(COUNT(Orders!$A:$A)=0,"",MAX(Orders!$A:$A))`; C1 `=IF(ISNUMBER(B1),IF(AND(B1=INT(B1),B1>=1901,B1<=9998),"OK","Type a whole year, like 2026"),"Type a whole year, like 2026")`
  - C16 `=IF(OR($C$1<>"OK",COUNT(C4:C15)=0),"",SUM(C4:C15))`; D16 `=IF($C$1<>"OK","",SUM(D4:D15))`; F16 `=IF(OR(C16="",D16=""),"",IF(D16=0,"n/a",C16/D16-1))`; D17 `=IF(C16="","",SUMPRODUCT(--ISNUMBER(C4:C15),D4:D15))`
  - C19 `=IF(COUNT(E4:E15)=0,"n/a",IF(MAX(E4:E15)<=0,"No month is up",INDEX(A4:A15,MATCH(MAX(E4:E15),E4:E15,0))))` (the topics file's plain `INDEX(A4:A15,MATCH(MAX(E4:E15),E4:E15,0))` returned the first month when all changes were 0 or all blank; the guard is the fix)
  - E2 CHECK: the two yearly totals equal a SUMIFS over the full calendar year (this year `">="&DATE(B1,1,1)` to `"<"&DATE(B1+1,1,1)`, last year the same with B1-1) to within 0.005, and C21 is 0. Full text is in the file; when no month of this year is shown it checks last year's total and that there are no orders in this year.
- Errors shown on purpose: none. A month with no sales last year shows n/a (never #DIV/0!). Months that start after the latest order date are blank in this year (not zero) and in the change columns.
- Example table (sample, year 2026, latest order 09/29/2026; every cell equals the model):
  | Month | This year | Last year | Change $ | Change % |
  | January | $353.02 | $171.32 | $181.70 | 106.1% |
  | February | $222.48 | $76.39 | $146.09 | 191.2% |
  | March | $480.83 | $0.00 | $480.83 | n/a |
  | April | $294.09 | $312.40 | -$18.31 | -5.9% |
  | May | $391.20 | $276.42 | $114.78 | 41.5% |
  | June | $275.80 | $289.21 | -$13.41 | -4.6% |
  | July | $376.10 | $140.21 | $235.89 | 168.2% |
  | August | $421.72 | $234.86 | $186.86 | 79.6% |
  | September | $354.26 | $286.95 | $67.31 | 23.5% |
  | October to December | blank | $441.61, $334.83, $221.67 | blank | blank |
  Total $3,169.50 vs $2,785.87, +$383.63, +13.8% (a full last year against a partial this year, so read row 17 too: last year same months $1,787.76, +$1,381.74, +77.3%). Best month March, weakest April, E2 OK, C21 0, C22 0.
- Edge findings (23 workbooks, 1,776 cell comparisons, 0 mismatches after the fixes; model is Python grouped sums by month):
  - Year 2025 (no 2024 data): last year all $0.00, every percent n/a, F16 n/a, best month October (all dollars up), weakest "No month is down"; C22 63 (the 2026 orders are outside the two years). Year 2024 (before the data): all zeros, "No month is up" and "No month is down". Year 2027 or 2030: no month is shown, so this year's totals and percent stay blank (D16 still shows last year's $3,169.50 for 2027), E2 OK.
  - A latest order of 10/01/2026 12:00 makes October show (partial month); latest 09/26/2026 shows September. A full year of orders (last 12/31/2026) shows all 12 months, $3,321.00 vs $2,785.87.
  - Boundaries: 01/31/2026 23:59:59 stays in January, 02/01/2026 00:00 is February, 02/28 23:59:59 stays in February, 03/01 is March; 12/31 and 01/01 orders land in the correct years (10, 20, 30, 40 test: total 70 vs 30 with 2024 and 2027 orders excluded and counted in C22 as 2).
  - Leap day: 02/29/2024 counts in February 2024. Year 2025 against 2024: February $50.00 vs $101.00 (02/28 plus 02/29), March $7.00 vs $2.00; 2024 against 2023: February $101.00 vs $4.00. EDATE(first of month,-12) never lands on a missing 02/29 because it is always the first.
  - Refund as a negative amount: reduces the month; a month whose net is negative (-$30 + $10 = -$20 in May 2026 vs $20) shows -200% (the percent on a negative base can look odd; the sheet shows it without judging). Base 0 and this year 0: n/a. Fractions of a cent (0.1, 0.2, 0.3, 0.15, 0.15): the ROUND(...,2) in E removes float tails.
  - Text date row ("03/15/2026" typed as text): left out of every sum, C21 1, E2 CHECK. Blank or text amount ("n/a"): counted as 0, no error values.
  - Invalid year (text "twenty", 2026.5, 1899): C1 message, month rows and totals blank, E2 CHECK, no error values (the text year gave #VALUE! everywhere before the fix: AND does not short-circuit, so INT of text failed; rewritten as nested IF). 400 random orders with negatives over 2025 and 2026: all matched for 2026 and 2025.
- Limits: a blank this-year month means "after the latest order", not "zero sales"; the latest month shown can be partial. Whole-column SUMIFS ran fine in LibreOffice; Excel and Sheets not tested. Not run: 10,000 or more rows, Orders dates with a time zone, years before 1901. Dollars only; no advice on what a change means.
