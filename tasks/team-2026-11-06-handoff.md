# Handoff 2026-11-06: five topics researched, hooks written, templates built and tested (LibreOffice Calc 24.2.7.2 only), for the script-writers

Stage 1 of the content team (researcher, hook-writer, builder) was done by one agent. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: tasks/team-topics-2026-11-06.md, tasks/hooks-2026-11-06.md, scripts/build-templates.py (`text_between_parentheses`, `ship_date_with_cutoff`, `change_due_bills_coins`, `shipping_zone_from_zip`, `new_vs_returning_customers`), five new files in templates/, results in tasks/briefs.md.

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 109 workbooks, 74,477 comparisons, 0 mismatches against independent Python models, no error value in any recalculated workbook. All five files exported to PDF: every page US Letter, no error values in the PDF text. Four bugs and one test-case removal are listed in tasks/briefs.md ("2026-11-06 builder results"); all fixed and re-run.

Post file names: `posts/ready/2026-11-06-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only (workbook and comparison counts below).

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-06-team-01-extract-text-between-parentheses-mid-find-excel-google-sheets | tidy-tabs-text-between-parentheses.xlsx | 12 workbooks, 6,072 comparisons |
| 2 | 2026-11-06-team-02-ship-date-daily-cutoff-time-workday-excel-google-sheets | tidy-tabs-ship-date-cutoff-time.xlsx | 32 workbooks, 12,992 comparisons |
| 3 | 2026-11-06-team-03-change-due-count-bills-coins-int-sumproduct-excel-google-sheets | tidy-tabs-change-due-bills-coins.xlsx | 18 workbooks, 12,150 comparisons |
| 4 | 2026-11-06-team-04-shipping-zone-from-zip-code-index-match-excel-google-sheets | tidy-tabs-shipping-zone-from-zip.xlsx | 21 workbooks, 17,367 comparisons |
| 5 | 2026-11-06-team-05-new-vs-returning-customers-by-month-minifs-countifs-excel-google-sheets | tidy-tabs-new-vs-returning-customers.xlsx | 26 workbooks, 25,896 comparisons |

Every workbook count includes the committed sample. "Comparisons" are cell-by-cell checks of the recalculated workbook against the Python model.

---

## 1. How to Extract Text Between Parentheses in Excel and Google Sheets
Hook: Pull the text inside the first pair of parentheses with `=TRIM(MID(A5,FIND("(",A5)+1,FIND(")",A5,FIND("(",A5))-FIND("(",A5)-1))`. For "Soy candle (Lavender, 8 oz)" it returns Lavender, 8 oz.

Template: templates/tidy-tabs-text-between-parentheses.xlsx. Sheets: **Listings**, **How to use**.

Layout (Listings): A1 title, A2 note. Header row 4. Input rows 5 to 104 (100 titles, column A yellow). B "Inside the parentheses", C "Title without them", D "Status", E "Before the first comma", F "After the first comma". Summary block H4:I9: I4 titles listed, I5 with a pair, I6 with no parentheses, I7 missing the closing ), I8 empty parentheses, I9 CHECK.

Key formulas (row 5):
- B5 `=IF(A5="","",IF(ISNUMBER(FIND("(",A5)),IF(ISNUMBER(FIND(")",A5,FIND("(",A5))),TRIM(MID(A5,FIND("(",A5)+1,FIND(")",A5,FIND("(",A5))-FIND("(",A5)-1)),""),""))`
- D5 `=IF(A5="","",IF(ISNUMBER(FIND("(",A5)),IF(ISNUMBER(FIND(")",A5,FIND("(",A5))),IF(B5="","Empty ()","OK"),"No closing )"),"No ("))`
- C5 `=IF(A5="","",IF(OR(D5="OK",D5="Empty ()"),TRIM(LEFT(A5,FIND("(",A5)-1)&" "&MID(A5,FIND(")",A5,FIND("(",A5))+1,500)),TRIM(A5)))`
- E5 `=IF(B5="","",TRIM(LEFT(B5,FIND(",",B5&",")-1)))`
- F5 `=IF(ISNUMBER(FIND(",",B5)),TRIM(MID(B5,FIND(",",B5)+1,500)),"")`
- I9 `=IF(AND(I7=0,I8=0),"OK","Check")`

Sample (12 titles, rows 5 to 16): I4 12, I5 11, I6 1, I7 0, I8 0, I9 OK.
- A5 Soy candle (Lavender, 8 oz): B5 Lavender, 8 oz; C5 Soy candle; D5 OK; E5 Lavender; F5 8 oz.
- A9 Enamel mug (Blue): B9 Blue; C9 Enamel mug; E9 Blue; F9 blank (no comma).
- A11 Linen napkin set (Natural, Set of 4): B11 Natural, Set of 4; C11 Linen napkin set; E11 Natural; F11 Set of 4.
- A13 Letterpress birthday card: B13 blank; C13 Letterpress birthday card; D13 No (.
- A14 Tote bag (Canvas, 12 oz) Medium: B14 Canvas, 12 oz; C14 Tote bag Medium.
- A15 Wool dryer balls (3 pack) (Gray): B15 3 pack; C15 Wool dryer balls (Gray) (only the first pair is used); E15 3 pack.
- Other titles: rows 6 to 8, 10, 12, 16 (Soy candle (Cedar, 8 oz); Brass hoop earrings (Gold, Small); Brass hoop earrings (Silver, Large); Enamel mug (Speckled White); Linen napkin set (Sage, Set of 6); Oatmeal goat milk soap bar (Unscented)).

Limitations / tested behavior: only the first "(" and the first ")" after it; a nested pair is cut at the first ) (tested: "mug (a (b) c) d" gives a (b); no claim beyond that); a title with ( but no ) gives blank B, status "No closing )" and C shows the trimmed title; "()" and "( )" give Empty (); a number in column A is treated as text; TRIM only removes ordinary spaces (a tab inside stays). Not tested: Excel, Sheets, titles over 300 characters, non-breaking spaces.

---

## 2. How to Calculate a Ship Date With a Daily Cutoff Time in Excel and Google Sheets
Hook: Find the day an order counts with `=IF(ROUND(MOD(B7,1)*86400,0)<=ROUND($B$2*86400,0),WORKDAY(INT(B7)-1,1,Holidays!$A$5:$A$24),WORKDAY(INT(B7),1,Holidays!$A$5:$A$24))`, then add the shipping days with `=WORKDAY(C7,$B$3,Holidays!$A$5:$A$24)`. With a 2:00 PM cutoff, an order at 5:05 PM on Wednesday 11/25/2026 ships Monday 11/30/2026.

Template: templates/tidy-tabs-ship-date-cutoff-time.xlsx. Sheets: **Orders**, **Holidays**, **How to use**.

Layout (Orders): A2 "Daily cutoff time", **B2** 2:00 PM (yellow, h:mm AM/PM). A3 "Business days to ship after the order counts", **B3** 1 (yellow). A4 "Settings check", **B4**. Header row 6, input rows 7 to 106 (100 orders). A order #, B order date and time (yellow, mm/dd/yyyy h:mm AM/PM), C "Order counts as", D "Ship date", E "Ship day" (Mon, Tue ...), F "Days after order date". Summary H6:I10: I6 orders listed, I7 latest ship date, I8 longest wait (calendar days), I9 rows showing Check, I10 CHECK.
Layout (Holidays): header row 4 (A "Closed date", B "Day"), input A5:A24 (20 dates, yellow), B day name. Sample closed days A5 11/26/2026, A6 12/25/2026, A7 01/01/2027.

Key formulas (row 7):
- B4 `=IF(AND(ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B2>=0,B2<1,B3>=0,B3=INT(B3),COUNT(Holidays!A5:A24)=COUNTA(Holidays!A5:A24)),"OK","Check"),"Check")`
- C7 `=IF(B7="","",IF(AND(ISNUMBER(B7),$B$4="OK"),IF(B7>=1,IF(ROUND(MOD(B7,1)*86400,0)<=ROUND($B$2*86400,0),WORKDAY(INT(B7)-1,1,Holidays!$A$5:$A$24),WORKDAY(INT(B7),1,Holidays!$A$5:$A$24)),"Check"),"Check"))`
- D7 `=IF(B7="","",IF(ISNUMBER(C7),WORKDAY(C7,$B$3,Holidays!$A$5:$A$24),"Check"))`
- E7 `=IF(ISNUMBER(D7),TEXT(D7,"ddd"),D7)`; F7 `=IF(ISNUMBER(D7),D7-INT(B7),D7)`
- I10 `=IF(AND(B4="OK",I9=0,COUNT(B7:B106)=COUNTA(B7:B106)),"OK","Check")`

Sample (cutoff 2:00 PM, 1 day, rows 7 to 18; counts as / ship date / day / days after):
- 5101 11/02/2026 9:15 AM: 11/02/2026, 11/03/2026 Tue, 1
- 5102 11/02/2026 1:59:59 PM: 11/02/2026, 11/03/2026 Tue, 1
- 5103 11/02/2026 2:00:00 PM (exactly the cutoff, counts that day): 11/02/2026, 11/03/2026 Tue, 1
- 5104 11/02/2026 2:00:01 PM: 11/03/2026, 11/04/2026 Wed, 2
- 5105 11/05/2026 3:30 PM (Thursday): 11/06/2026, 11/09/2026 Mon, 4
- 5106 11/06/2026 4:45 PM (Friday): 11/09/2026, 11/10/2026 Tue, 4
- 5107 11/07/2026 10:00 AM (Saturday): 11/09/2026, 11/10/2026 Tue, 3
- 5108 11/08/2026 10:10 PM (Sunday): 11/09/2026, 11/10/2026 Tue, 2
- 5109 11/25/2026 8:30 AM (Wednesday): 11/25/2026, 11/27/2026 Fri (Thursday is closed), 2
- 5110 11/25/2026 5:05 PM: 11/27/2026, 11/30/2026 Mon, 5
- 5111 11/26/2026 11:20 AM (closed Thursday): 11/27/2026, 11/30/2026 Mon, 4
- 5112 12/24/2026 3:00 PM (Thursday): 12/28/2026 (Friday 12/25 is closed), 12/29/2026 Tue, 5
Summary: I6 12, I7 12/29/2026, I8 5, I9 0, I10 OK, B4 OK.

Limitations / tested behavior: the cutoff test is "at or before", compared in whole seconds; days 0 means ship on the day the order counts; a cutoff of 12:00 AM only counts orders at exactly midnight; B2 must be a time (0 up to but not including 1), B3 a whole number of 0 or more, every cell in Holidays!A5:A24 a date or empty, otherwise B4 and all rows say Check; a time-only or text order cell says Check; a date typed in the Holidays list that falls on a weekend has no effect. The template does not know carrier pickup times, weekend shipping or delivery time. TEXT(D7,"ddd") shows English day names only in an English setting (not tested otherwise). Not tested: Excel or Sheets WORKDAY behavior, TODAY.

---

## 3. How to Calculate Change Due and Count Bills and Coins in Excel and Google Sheets
Hook: Get the change with `=ROUND(C8-B8,2)`, turn it into whole cents with `=ROUND(D8*100,0)`, then count each bill or coin with `=INT($E8/F$4)` and the matching formulas in the columns to its right. A $31.35 sale paid with $50.00 gives $18.65 change: one $10, one $5, three $1 bills, two quarters, one dime and one nickel.

Template: templates/tidy-tabs-change-due-bills-coins.xlsx. Sheets: **Change** (landscape), **How to use**.

Layout (Change): A1 title, A2 note. **Row 4**: A4 "Value in cents", F4 to M4 = 2000, 1000, 500, 100, 25, 10, 5, 1 (typed constants, bold). **Row 5**: A5 "Total to hand out, all sales", D5 total change, F5 to M5 = sum of each column. Header row 7, input rows 8 to 67 (60 sales). A sale #, B sale total (yellow), C cash given (yellow), D change due, E change in cents, F $20 bills, G $10 bills, H $5 bills, I $1 bills, J quarters, K dimes, L nickels, M pennies, N Check. Summary P7:Q12: Q7 sales listed, Q8 sales with exact change, Q9 sales where cash is short, Q10 largest change due, Q11 rows showing Check, Q12 CHECK.

Key formulas (row 8):
- D8 `=IF(AND(B8="",C8=""),"",IF(AND(ISNUMBER(B8),ISNUMBER(C8)),IF(AND(B8>=0,C8>=0),IF(ROUND(C8-B8,2)>=0,ROUND(C8-B8,2),"Short"),"Check"),"Check"))`
- E8 `=IF(ISNUMBER(D8),ROUND(D8*100,0),"")`
- F8 `=IF(ISNUMBER($E8),INT($E8/F$4),"")`
- G8 `=IF(ISNUMBER($E8),INT(($E8-SUMPRODUCT($F8:F8,$F$4:F$4))/G$4),"")` (each later column widens the range: M8 uses `$F8:L8` and `$F$4:L$4`)
- N8 `=IF(ISNUMBER(E8),IF(SUMPRODUCT(F8:M8,$F$4:$M$4)=E8,"OK","Check"),D8)` (shows Short or Check from D when there is no change in cents)
- F5 `=SUM(F8:F67)`; D5 `=SUM(D8:D67)`; Q12 `=IF(AND(Q11=0,COUNT(B8:B67)=COUNT(C8:C67),Q9=0),"OK","Check")`

Sample (rows 8 to 17: sale total, cash given, change, then $20 / $10 / $5 / $1 / quarters / dimes / nickels / pennies):
- 1: $7.50, $20.00, $12.50: 0 / 1 / 0 / 2 / 2 / 0 / 0 / 0
- 2: $12.25, $20.00, $7.75: 0 / 0 / 1 / 2 / 3 / 0 / 0 / 0
- 3: $18.00, $20.00, $2.00: 0 / 0 / 0 / 2 / 0 / 0 / 0 / 0
- 4: $4.75, $5.00, $0.25: one quarter
- 5: $23.40, $40.00, $16.60: 0 / 1 / 1 / 1 / 2 / 1 / 0 / 0
- 6: $9.99, $10.00, $0.01: one penny
- 7: $15.00, $15.00, $0.00: nothing
- 8: $31.35, $50.00, $18.65: 0 / 1 / 1 / 3 / 2 / 1 / 1 / 0
- 9: $6.85, $10.00, $3.15: 0 / 0 / 0 / 3 / 0 / 1 / 1 / 0
- 10: $8.20, $20.00, $11.80: 0 / 1 / 0 / 1 / 3 / 0 / 1 / 0
Summary: D5 $72.71; row 5 pieces F5 to M5 = 0, 4, 3, 14, 13, 3, 3, 1; Q7 10, Q8 1, Q9 0, Q10 $18.65, Q11 0, Q12 OK. (Check: 4 x 10 + 3 x 5 + 14 + 13 x 0.25 + 3 x 0.10 + 3 x 0.05 + 0.01 = $72.71.)

Limitations / tested behavior: fewest pieces using only those eight values (no $2 bills, half dollars, $50 or $100 bills, and it does not know what is in the cash box). Type amounts in whole cents: a difference of exactly half a cent (2.505 - 2.50, 5.00 - 4.995) is stored as 0.004999... and gives change $0.00 instead of $0.01 or Short (found in a standalone run; the post should mention it as the reason for typing whole cents). 4.355 from 5.00 gives $0.65 and 4.345 gives $0.66 (checked). Cash less than the total gives "Short" in D and N and CHECK says Check; text, blank on one side or negative amounts give Check. Cash up to $99,999.99 tested. Not tested: Excel or Sheets, more than 60 sales.

---

## 4. How to Look Up a Shipping Zone From a ZIP Code in Excel and Google Sheets
Hook: Restore a ZIP that lost its zero with `=TEXT(C5,"00000")`, take the first three digits with `=VALUE(LEFT(D5,3))`, then look up the zone with `=INDEX(Zones!$B$5:$B$14,MATCH(E5,Zones!$A$5:$A$14,1))`. In the sample, 2108 becomes 02108 and lands in Zone 1 at $5.90 (a made-up zone table).

Template: templates/tidy-tabs-shipping-zone-from-zip.xlsx. Sheets: **Orders**, **Zones**, **How to use**.

Layout (Orders): A1 title, A2 note. Header row 4, input rows 5 to 204 (200 orders). A order #, B customer, C "ZIP as typed" (yellow; text cells hold ZIPs with a leading zero, number cells hold ZIPs that lost it), D "ZIP (5 digits)" (text), E "First 3 digits" (number), F zone, G shipping price.
Layout (Zones): A1 title, A2 note. Header row 4. Table rows 5 to 14 (10 slots, yellow): A "Zone starts at first 3 digits" (numbers, smallest first, first is 0), B zone name, C shipping price; D orders, E shipping total (formulas). Row 15 "All zones": D15, E15. Summary G4:H8: **H4** zone table check, H5 orders listed, H6 orders with a ZIP to check, H7 orders counted in a zone (= D15), **H8** CHECK.

Key formulas (Orders row 5):
- D5 `=IF(C5="","",IF(ISNUMBER(C5),IF(AND(C5>=0,C5<=99999,C5=INT(C5)),TEXT(C5,"00000"),"Check"),IF(OR(LEN(TRIM(C5))=5,AND(LEN(TRIM(C5))=10,MID(TRIM(C5),6,1)="-")),IF(SUMPRODUCT(--ISNUMBER(FIND(MID(TRIM(C5),{1,2,3,4,5},1),"0123456789")))=5,LEFT(TRIM(C5),5),"Check"),"Check")))`
- E5 `=IF(OR(D5="",D5="Check"),"",VALUE(LEFT(D5,3)))`
- F5 `=IF(D5="","",IF(OR(D5="Check",Zones!$H$4<>"OK"),"Check",INDEX(Zones!$B$5:$B$14,MATCH(E5,Zones!$A$5:$A$14,1))))`
- G5 `=IF(D5="","",IF(F5="Check","Check",INDEX(Zones!$C$5:$C$14,MATCH(E5,Zones!$A$5:$A$14,1))))`
- Zones D5 `=IF(B5="","",COUNTIF(Orders!$F$5:$F$204,B5))`; E5 `=IF(B5="","",SUMIF(Orders!$F$5:$F$204,B5,Orders!$G$5:$G$204))`
- H4 `=IFERROR(IF(AND(COUNT(A5:A14)>=1,A5=0,COUNT(A5:A14)=COUNTA(A5:A14),COUNT(A5:A14)=MATCH(9.99E+307,A5:A14),SUMPRODUCT(--(A6:A14>A5:A13),--ISNUMBER(A6:A14))=COUNT(A5:A14)-1,COUNT(C5:C14)=COUNT(A5:A14),COUNTA(B5:B14)=COUNT(A5:A14)),"OK","Check"),"Check")`
- H8 `=IF(AND(H4="OK",H6=0,H7=H5),"OK","Check")`

Zone table (fictional, "for a shop shipping from Boston, MA"; start, zone, price): 0 Zone 1 $5.90; 100 Zone 2 $6.40; 200 Zone 3 $7.10; 300 Zone 4 $7.80; 500 Zone 5 $8.60; 700 Zone 6 $9.30; 850 Zone 7 $10.20; 960 Zone 8 $11.40. The 1 in MATCH finds the last start not above the number, so prefix 164 uses the 100 row.

Sample (orders rows 5 to 20; typed -> D5-style ZIP, first 3 digits, zone, price):
- 7001 Maria Alvarez "02108" (text) -> 02108, 21, Zone 1, $5.90
- 7002 Ben Carter 2108 (number) -> 02108, 21, Zone 1, $5.90
- 7003 Priya Shah "07030" -> 07030, 70, Zone 1, $5.90; 7004 Tom Nguyen 7030 (number) -> same
- 7005 Lena Fischer 10001 -> 100, Zone 2, $6.40
- 7006 Omar Haddad 33101 -> 331, Zone 4, $7.80
- 7007 Chloe Park "60601" -> 606, Zone 5, $8.60
- 7008 Jamal Reed 94105 -> 941, Zone 7, $10.20
- 7009 Sofia Rossi 98101 -> 981, Zone 8, $11.40
- 7010 Ana Torres 73301 -> 733, Zone 6, $9.30
- 7011 Dan Whitaker "02134-1234" (ZIP+4) -> 02134, 21, Zone 1, $5.90
- 7012 Kira Olsen 90210 -> 902, Zone 7, $10.20
- 7013 Ravi Patel "30301" -> 303, Zone 4, $7.80
- 7014 Grace Kim "85001" -> 850, Zone 7, $10.20 (exactly on a boundary)
- 7015 Noah Brooks "05401" -> 54, Zone 1, $5.90
- 7016 Ella Davis 80202 -> 802, Zone 6, $9.30
Zones tab results: Zone 1 6 orders $35.40; Zone 2 1 / $6.40; Zone 3 0 / $0.00; Zone 4 2 / $15.60; Zone 5 1 / $8.60; Zone 6 2 / $18.60; Zone 7 3 / $30.60; Zone 8 1 / $11.40; D15 16, E15 $126.60. H4 OK, H5 16, H6 0, H7 16, H8 OK.

Limitations / tested behavior: a numeric ZIP is padded to 5 digits, so 123 becomes 00123 (it cannot know the original); numbers with decimals, negatives or above 99999 give Check; text must be 5 characters or ZIP+4 (10 characters with a dash in position 6) after TRIM, and the first 5 must be digits (so "1e345" and "12E45" give Check; letters give Check); only the first 5 digits of a ZIP+4 are checked. Zone names must be unique (a repeated name makes CHECK say Check). A table that does not start at 0, is not rising, has a gap, has a text price or a blank name makes H4 and every row say Check; an empty table gives Check. The zones and prices are made up; the post must say to use the carrier's chart for real prices. Not tested: Excel or Sheets; a TRUE/FALSE value in the ZIP column (LibreOffice treats it as a number); more than 200 orders.

---

## 5. How to Count New and Returning Customers Each Month in Excel and Google Sheets
Hook: Label an order New when its date equals the customer's earliest date, with `=IF(B6=E6,"New","Returning")`, where E6 is `=MINIFS($B$6:$B$305,$C$6:$C$305,C6)`. In the sample, October has 3 new customers and 8 of its 11 orders come from returning customers.

Template: templates/tidy-tabs-new-vs-returning-customers.xlsx. Sheets: **Orders**, **Months**, **How to use**.

Layout (Orders): A1 title, A2 and A3 notes. Header row 5, input rows 6 to 305 (300 orders). A order #, B order date (yellow, dates only, no times), C customer (yellow), D amount (yellow, $), E "First order date", F "Type", G "New customer" (1 on the first listed row of a customer's first date).
Layout (Months): A1 title, **B2** year (yellow, 2026). Header row 4, month rows 5 to 16 (A5 = January ... A16 = December, shown as the month name), **row 17 "Year total"**. Columns: A month, B orders, C new customers, D orders on first day, E orders from returning, F sales on first day, G sales from returning, H returning share of orders. Summary J4:K8: K4 order rows listed, K5 orders in other years, K6 rows showing Check, K7 customers seen in the list, **K8** CHECK.

Key formulas:
- Orders E6 `=IF(OR(B6="",C6=""),"",IF(AND(ISNUMBER(B6),B6=INT(B6)),_xlfn.MINIFS($B$6:$B$305,$C$6:$C$305,C6),"Check"))` (shown to readers as `MINIFS`; the file stores the `_xlfn.` prefix)
- F6 `=IF(E6="","",IF(E6="Check","Check",IF(B6=E6,"New","Returning")))`
- G6 `=IF(F6="New",IF(COUNTIFS($C$6:C6,C6,$B$6:B6,B6)=1,1,0),IF(F6="Returning",0,""))`
- Months A5 `=DATE($B$2,1,1)`
- B5 `=COUNTIFS(Orders!$B$6:$B$305,">="&$A5,Orders!$B$6:$B$305,"<"&EDATE($A5,1))`
- C5 `=SUMIFS(Orders!$G$6:$G$305,Orders!$B$6:$B$305,">="&$A5,Orders!$B$6:$B$305,"<"&EDATE($A5,1))`
- D5 same range pair plus `,Orders!$F$6:$F$305,"New")` with COUNTIFS; E5 the same with "Returning"
- F5 `=SUMIFS(Orders!$D$6:$D$305,Orders!$B$6:$B$305,">="&$A5,Orders!$B$6:$B$305,"<"&EDATE($A5,1),Orders!$F$6:$F$305,"New")`; G5 the same with "Returning"
- H5 `=IF(B5=0,"",E5/B5)` (format 0.0%)
- K8 `=IF(AND(K6=0,COUNT(Orders!B6:B305)=COUNTA(Orders!B6:B305),COUNT(Orders!B6:B305)=COUNTA(Orders!C6:C305),B17=D17+E17,K5=0),"OK","Check")`

Sample (30 fictional orders 1001 to 1030, 08/03/2026 to 10/31/2026, 12 customers; rows 6 to 35). Examples: 1001 Maria Alvarez 08/03/2026 $42.00 New, flag 1; 1004 Maria Alvarez 08/12/2026 Returning; 1005 and 1006 Tom Nguyen both 08/14/2026 are both New and only 1005 has flag 1.
Months tab results (month: orders / new customers / orders on first day / orders from returning / sales on first day / sales from returning / returning share):
- August (row 12): 9 / 5 / 6 / 3 / $220.25 / $80.75 / 33.3%
- September (row 13): 10 / 4 / 4 / 6 / $225.75 / $174.74 / 60.0%
- October (row 14): 11 / 3 / 3 / 8 / $158.00 / $243.75 / 72.7%
- Year total (row 17): 30 / 12 / 13 / 17 / $604.00 / $499.24 / 56.7%
- Other months: all 0, share blank. K4 30, K5 0, K6 0, K7 12, K8 OK.

Limitations / tested behavior: "New" only means the customer's first date in the list; anyone who ordered before the list starts looks new in the first month. Orders on a customer's first date are all "Orders on first day" (D), but only one row counts in "New customers" (C). Names match without regard to capitals; a space before or after a name makes a different customer. A date with a time or a text date gives Check, and a time-of-day date can still pull a customer's earliest date, so fix every Check row first. Orders outside the year in B2 are not in the table (K5 counts them and CHECK says Check). MINIFS is in Excel 2019 and Microsoft 365 and in Google Sheets, not in Excel 2016 or earlier (the file's How to use tab says so). Not tested: Excel or Sheets, wildcard characters (* ? ~) in customer names, more than 300 orders.

---

## Notes for all five script-writers
- Every formula above is the exact text in the committed file; copy cell references from this file, not from memory.
- Sample facts are exact (read back from the recalculated sample). Sample data is fictional; say so once.
- Official help links to cite for functions (Microsoft/Google): MID, FIND, TRIM, WORKDAY, MOD, INT, ROUND, SUMPRODUCT, INDEX, MATCH, TEXT, VALUE, LEFT, MINIFS, COUNTIFS, SUMIFS, EDATE. These links were NOT checked in this stage; check them before listing.
- No hero image, screenshot, threads line or image_prompts exist yet; the PDF exports used for page-count checks were scratch files and are not in the repo.
- Excel and Google Sheets behavior is not claimed anywhere. Say "checked in LibreOffice Calc 24.2.7.2 only".
