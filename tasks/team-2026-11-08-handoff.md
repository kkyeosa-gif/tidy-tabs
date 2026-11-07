# Team handoff - 2026-11-08 (hook-writer + template-builder to script-writer)

Post files: posts/ready/2026-11-08-team-0N-<slug>.md. Hooks (titles and openings) are in tasks/hooks-2026-11-08.md; test log in tasks/briefs.md (2026-11-08 builder results). Build: scripts/build-templates.py (functions `profit_by_product`, `state_abbreviations`, `items_per_order_one_cell`, `last_business_day`, `date_hit_sales_goal`). All values below were recalculated in LibreOffice Calc 24.2.7.2 (headless) and matched independent Python models. Excel and Google Sheets were NOT tested; `tested_in` must say LibreOffice only. Every workbook has a "How to use" tab (notes), yellow = input cells, gray = formula cells. No topic was changed from tasks/team-topics-2026-11-08.md. Two deviations from the topic file are listed under posts 4 and 5 and in the pitfalls.

Rules for all five: no TODAY, FILTER, LET, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays. Fictional data. Do not state a good margin, a legal holiday or a legal return rule.

---
## 1. profit-by-product-from-sales-log-sumifs-excel-google-sheets
Title: How to Calculate Profit by Product From a Sales Log in Excel and Google Sheets
Template: templates/tidy-tabs-profit-by-product.xlsx (tabs: Sales, Products, How to use)

Key cells
- Products tab: A1:C21 price list (A product, B price, C unit cost; data in A2:C7, slots to row 21). D Units, E Revenue after discounts, F Cost of goods, G Profit, H Margin in rows 2:21; Total row 22. Side block J1:K7: K1 products listed, K2 products with no sales, K3 units (summary), K4 units on Sales tab, K5 unmatched units (must be 0), K6 Sales rows showing Check, K7 CHECK.
- Sales tab: A1:F61 (A Date, B Order, C Product, D Qty, E Discount $, F Row check); data in rows 2:13, slots to row 61.

Sample data
Products: Soy Candle 8 oz 18.00 / 6.40; Lip Balm 3.50 / 0.95; Tote Bag 24.00 / 9.10; Enamel Pin 8.00 / 2.30; Sticker Pack 4.50 / 0.70; Gift Box 32.00 / 14.25.
Sales (date, order, product, qty, discount): 10/01 1001 Soy Candle 8 oz 3 0; 10/02 1002 Tote Bag 1 2.00; 10/04 1003 Lip Balm 4 0; 10/06 1004 Enamel Pin 2 0; 10/09 1005 Soy Candle 8 oz 2 1.50; 10/11 1006 Sticker Pack 6 3.00; 10/14 1007 Tote Bag 2 0; 10/16 1008 Lip Balm 3 0; 10/19 1009 Enamel Pin 5 5.00; 10/21 1010 Sticker Pack 4 0; 10/25 1011 Soy Candle 8 oz 1 0; 10/28 1012 Tote Bag 1 0. (all 2026)

Formulas (row 2)
- D2 `=IF(A2="","",SUMIFS(Sales!$D$2:$D$61,Sales!$C$2:$C$61,$A2))`
- E2 `=IF(A2="","",IF(AND(ISNUMBER(B2),ISNUMBER(C2)),ROUND(D2*B2-SUMIFS(Sales!$E$2:$E$61,Sales!$C$2:$C$61,$A2),2),"Check"))`
- F2 `=IF(ISNUMBER(E2),ROUND(D2*C2,2),E2)`; G2 `=IF(ISNUMBER(E2),E2-F2,E2)`; H2 `=IF(ISNUMBER(G2),IF(E2=0,"",G2/E2),G2)`
- Row 22: D22:G22 `=SUM(D2:D21)` etc., H22 `=IF(E22=0,"",G22/E22)`
- K5 `=K4-K3` (K4 `=SUM(Sales!D2:D61)`), K7 `=IF(AND(ROUND(K5,6)=0,K6=0,COUNTIF(E2:E21,"Check")=0),"OK","Check")`
- Sales F2: Check product / Check qty / Check discount / OK per row (name not on the price list, text or negative qty, discount text or larger than price x qty).

Verified results (sample)
| Product | Units | Revenue | Cost | Profit | Margin |
|---|---|---|---|---|---|
| Soy Candle 8 oz | 6 | $106.50 | $38.40 | $68.10 | 63.9% |
| Lip Balm | 7 | $24.50 | $6.65 | $17.85 | 72.9% |
| Tote Bag | 4 | $94.00 | $36.40 | $57.60 | 61.3% |
| Enamel Pin | 7 | $51.00 | $16.10 | $34.90 | 68.4% |
| Sticker Pack | 10 | $42.00 | $7.00 | $35.00 | 83.3% |
| Gift Box | 0 | $0.00 | $0.00 | $0.00 | blank |
| Total | 34 | $318.00 | $104.55 | $213.45 | 67.1% |
K5 = 0, K7 = OK. Misspelled name ("Tote  Bag" with two spaces, 3 units): K5 = 3, Sales row says Check product, K7 Check.

Pitfalls
- Margin on a product with revenue 0 is blank on purpose (no #DIV/0!).
- Overall margin H22 is total profit / total revenue, not the average of the product margins (67.1% vs average of the five margins).
- SUMIFS is not case sensitive ("tote bag" matches) and treats * and ? as wildcards; trailing/extra spaces make a name unmatched.
- Text qty ("3") is ignored by SUMIFS, so K5 does NOT catch it; the Sales F column does (Check qty).
- A discount larger than the line gives negative revenue for that product (flagged Check discount); we only do arithmetic, no advice on target margin.
- Rounding: revenue and cost are ROUNDed to cents per product (half away from zero, Decimal half-up in the test).
- Duplicate product rows in the price list double-count units; K5 turns negative.
- Slots: 20 products, 60 sales rows. Sample Products total row sits at row 22.

---
## 2. state-name-to-abbreviation-lookup-index-match-excel-google-sheets
Title: How to Look Up a State Abbreviation From the Full State Name in Excel and Google Sheets
Template: templates/tidy-tabs-state-abbreviations.xlsx (tabs: Customers, States, How to use)

Key cells
- States tab: A1:B52 (A State name, B 2-letter code; 50 states plus District of Columbia, alphabetical). D1 note: source is USPS Publication 28 (link the publication as source in the post, no judgment).
- Customers tab: A1:F201 (A Name, B City, C State as typed, D State cleaned, E Code, F Status); data rows 2:13, slots to row 201. Summary H1:I5: I1 customers listed, I2 from name, I3 already a code, I4 Check rows, I5 "Rows add up" (OK/Check).

Sample data (Name, City, State as typed): Ava Martin, Austin, Texas; Ben Ortiz, Columbus, " ohio "; Cara Lee, Albany, NEW YORK; Dan Wu, Dallas, tx; Eli Shah, San Diego, Calif.; Fay Gray, Boise, (empty); Gus Hall, Reno, nevada; Hana Kim, Salem, OR; Ian Cole, Miami, Florida; Jo Park, Denver, Colorado; Kai Stone, Portland, ZZ; Lia Moss, Macon, Georgia.

Formulas (row 2)
- D2 `=IF(COUNTA(A2:C2)=0,"",TRIM(SUBSTITUTE(C2,_xlfn.UNICHAR(160)," ")))` (shown to users as `UNICHAR(160)`)
- E2 `=IF(COUNTA(A2:C2)=0,"",IFERROR(INDEX(States!$B$2:$B$52,MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(D2,"~","~~"),"*","~*"),"?","~?"),States!$A$2:$A$52,0)),IF(AND(LEN(D2)=2,SUMPRODUCT(--(States!$B$2:$B$52=D2))>0),UPPER(D2),"Check")))`
  Simplified teaching version for the post body: `=IFERROR(INDEX(States!$B$2:$B$52,MATCH(D2,States!$A$2:$A$52,0)),"Check")`, then the code fallback. Say the built file adds the wildcard escaping; do not claim the simplified version was the one tested.
- F2 `=IF(E2="","",IF(E2="Check","Check",IF(UPPER(D2)=E2,"Already a code","From name")))`

Verified results (sample)
| Typed | Code | Status |
|---|---|---|
| Texas | TX | From name |
| " ohio " | OH | From name |
| NEW YORK | NY | From name |
| tx | TX | Already a code |
| Calif. | Check | Check |
| (empty) | Check | Check |
| nevada | NV | From name |
| OR | OR | Already a code |
| Florida | FL | From name |
| Colorado | CO | From name |
| ZZ | Check | Check |
| Georgia | GA | From name |
Summary: 12 listed, 7 from name, 2 already a code, 3 Check, I5 OK. Also verified: all 51 names in 4 forms, "District of Columbia" gives DC, non-breaking space "New York" gives NY, number 12 gives Check, Puerto Rico and PR give Check, "T*", "??", "?" give Check (not wildcards).

Pitfalls
- Use UNICHAR(160), not CHAR(160): LibreOffice maps CHAR(160) through the system locale. UNICHAR needs Excel 2013 or later (Google Sheets has it).
- MATCH with 0 is case-insensitive and would read * and ? as wildcards; the built formula escapes them. A fully empty row (A:C empty) shows blank; a row with a name but no state shows Check.
- The code list is the USPS list; territories and military codes are intentionally absent and show Check.
- TRIM also collapses double spaces inside a name ("New  York" works).
- A state typed as a number is Check; periods ("Calif.", "CA.") are Check, the file does not guess.

---
## 3. combine-order-items-into-one-cell-running-join-excel-google-sheets
Title: How to Combine All Items in an Order Into One Cell in Excel and Google Sheets
Template: templates/tidy-tabs-items-per-order-one-cell.xlsx (tabs: Lines, Orders, How to use)

Key cells
- Lines tab: A1:F101 (A Order, B Item, C Qty, D Running text, E Last line?, F Sort check); data rows 2:13, slots to row 101.
- Orders tab: A1:D41 (A Order typed by you, B Items in the order, C Lines, D Lookup check); data rows 2:6, slots to row 41. Summary F1:G6: G1 orders listed, G2 not found, G3 lines found, G4 lines on Lines tab, G5 order numbers split apart, G6 CHECK.

Sample data (Order, Item, Qty): 2001 Soy candle Lavender 2; 2001 Lip balm 1; 2001 Gift box 1; 2002 Tote bag 1; 2003 Enamel pin 3; 2003 Sticker pack 2; 2003 Soy candle Cedar 1; 2003 Lip balm 4; 2004 Tote bag 2; 2004 Gift box 1; 2005 Sticker pack 5; 2005 Enamel pin 1. Orders tab A2:A6 = 2001 to 2005.

Formulas
- D2 `=IF(A2="","",IF(A2=A1,D1&", "&C2&" x "&B2,C2&" x "&B2))` (A1 is the header, text, so row 2 always starts new text)
- E2 `=IF(A2="","",IF(A2<>A3,1,0))`
- F2 `=IF(A2="","",IF(A2=A1,"",IF(COUNTIF($A$1:A1,A2)>0,"Order split","")))`
- Orders!B2 `=IF(A2="","",IFERROR(IF(INDEX(Lines!$A$2:$A$101,MATCH(A2,Lines!$A$2:$A$101,1))=A2,INDEX(Lines!$D$2:$D$101,MATCH(A2,Lines!$A$2:$A$101,1)),"Not found"),"Not found"))`
  Teaching version: `=INDEX(Lines!$D$2:$D$101,MATCH(A2,Lines!$A$2:$A$101,1))`; the built version adds the equality check so a missing order does not borrow the previous order's text.
- Orders!C2 `=IF(A2="","",COUNTIF(Lines!$A$2:$A$101,A2))`, D2 `=IF(A2="","",IF(B2="Not found","Check","OK"))`

Verified results (sample)
| Order | Items | Lines |
|---|---|---|
| 2001 | 2 x Soy candle Lavender, 1 x Lip balm, 1 x Gift box | 3 |
| 2002 | 1 x Tote bag | 1 |
| 2003 | 3 x Enamel pin, 2 x Sticker pack, 1 x Soy candle Cedar, 4 x Lip balm | 4 |
| 2004 | 2 x Tote bag, 1 x Gift box | 2 |
| 2005 | 5 x Sticker pack, 1 x Enamel pin | 2 |
G1..G6 = 5, 0, 12, 12, 0, OK. Also verified: single-line order, one order only, text order numbers (all text works), 528 character result, 100 lines (32 orders).

Pitfalls
- The Lines must be sorted by order and each order's rows together. In the unsorted test (lines for orders 2, 1, 2, 3) order 2 returned only "1 x A" instead of both of its lines "1 x A, 3 x C", F says Order split, G5 = 1 and G6 says Check. Show this as the "what breaks" example and the sort step as the fix.
- Mixing number 2001 and text "2001" splits the order (equal-looking, different values); F flags it.
- The MATCH type 1 on sorted data returns the LAST row of the order. Empty slots at the end of the range did not break it in LibreOffice; not tested in Excel/Sheets.
- The 255-character limit applies to text typed inside a formula, not to a result; a 528-character result worked in LibreOffice.
- Do not mention TEXTJOIN as used. Orders tab numbers must be ascending too.
- Lines F column has a header-row dependency: row 2 compares to the header row (A1), which is why a header is needed.

---
## 4. last-business-day-of-month-eomonth-workday-excel-google-sheets
Title: How to Find the Last Business Day of Each Month in Excel and Google Sheets
Template: templates/tidy-tabs-last-business-day.xlsx (tabs: Months, Holidays, How to use)

Key cells (built layout differs from the topic file: title and settings above the table)
- Months tab: B2 first month (input, 11/01/2026), B3 settings check; header row 5; rows 6:17 = 12 months: A Month start, B Month end, C Last business day, D Weekday, E Days earlier than month end. Summary G5:H9: H5 months listed, H6 month end is a business day, H7 moved back, H8 most days earlier, H9 CHECK.
- Holidays tab: A1:C21 (A name optional, B Date, C Date check); 10 sample dates in rows 2:11, slots to row 21.

Sample holidays (fictional closed days): 11/26/2026, 12/24/2026, 12/25/2026, 12/31/2026, 01/01/2027, 01/18/2027, 05/31/2027, 07/05/2027, 09/06/2027, 11/25/2027.

Formulas
- A6 `=IF($B$3<>"OK","Check",DATE(YEAR($B$2),MONTH($B$2),1))`; A7 `=IF(ISNUMBER(A6),EDATE(A6,1),"Check")`
- B6 `=IF(ISNUMBER(A6),EOMONTH(A6,0),"Check")`
- C6 `=IF(ISNUMBER(B6),WORKDAY(B6+1,-1,Holidays!$B$2:$B$21),"Check")`
- D6 `=IF(ISNUMBER(C6),TEXT(C6,"dddd"),"Check")`; E6 `=IF(ISNUMBER(C6),B6-C6,"Check")`
- B3 `=IF(AND(ISNUMBER(B2),COUNTIF(Holidays!C2:C21,"Check")=0),IF(AND(B2>=1,B2<=2900000),"OK","Check"),"Check")`
- Holidays C2 `=IF(B2="","",IF(ISNUMBER(B2),IF(AND(B2>=1,B2=INT(B2)),"OK","Check"),"Check"))`

Verified results (sample, start 11/01/2026)
| Month end | Last business day | Weekday | Days earlier |
|---|---|---|---|
| 11/30/2026 | 11/30/2026 | Monday | 0 |
| 12/31/2026 | 12/30/2026 | Wednesday | 1 |
| 01/31/2027 | 01/29/2027 | Friday | 2 |
| 02/28/2027 | 02/26/2027 | Friday | 2 |
| 03/31/2027 | 03/31/2027 | Wednesday | 0 |
| 04/30/2027 | 04/30/2027 | Friday | 0 |
| 05/31/2027 | 05/28/2027 | Friday | 3 |
| 06/30/2027 | 06/30/2027 | Wednesday | 0 |
| 07/31/2027 | 07/30/2027 | Friday | 1 |
| 08/31/2027 | 08/31/2027 | Tuesday | 0 |
| 09/30/2027 | 09/30/2027 | Thursday | 0 |
| 10/31/2027 | 10/29/2027 | Friday | 2 |
H5..H9 = 12, 6, 6, 3, OK. Also verified: Nov 30 2024 (Saturday) gives Friday 11/29/2024 with no holidays, leap Feb 2028 (02/29/2028 Tuesday), month end on a listed holiday, two listed holidays in a row, empty holiday list, 120 months, mid-month start.

Pitfalls
- Deviation from the topic: 12 rows and the Holidays range is column B (names in column A), so the range is Holidays!$B$2:$B$21 not $A$2:$A$12.
- WORKDAY(B+1,-1,...) is the trick: start the day after month end, step back one working day, so a month end that is already a business day stays put.
- Weekend is always Saturday and Sunday (WORKDAY, not WORKDAY.INTL). The holiday list is yours, not legal.
- A holiday date typed as text or with a time gives Check in Holidays C, B3 and every month row; it is not silently ignored.
- A closed day that falls on a weekend changes nothing. Holidays on 05/31/2027 (Monday) moves that month to Friday 05/28.
- TEXT(C,"dddd") shows English day names only in an English locale; untested otherwise.
- The formula copies for more months: drag row 17 down (the test used 120 rows this way).

---
## 5. date-you-hit-sales-goal-cumulative-total-match-excel-google-sheets
Title: How to Find the Date You Hit a Sales Goal in Excel and Google Sheets
Template: templates/tidy-tabs-date-hit-sales-goal.xlsx (tabs: Daily sales, How to use)

Key cells
- Daily sales tab: B1 goal (input, 10,000.00), B2 goal check; header row 3; A date, B sales (inputs), C running total, D goal reached?, E row check in rows 4:369 (366 slots; 14 sample rows 4:17). Summary G3:H10: H3 date you hit the goal, H4 days it took, H5 running total that day, H6 amount over, H7 total of all rows, H8 still to go, H9 rows needing attention, H10 CHECK.

Sample data (dates 09/01/2026 to 09/14/2026, sales): 1,250.00; 980.00; 0; 1,430.50; 760.00; 1,105.25; 0; 890.00; 1,320.00; 640.75; 1,510.00; 725.50; 980.00; 1,175.00.

Formulas
- C4 `=IF(COUNTA(A4:B4)=0,"",SUM($B$4:B4))`
- D4 `=IF(C4="","",IF($B$2<>"OK","Check",IF(C4>=$B$1,"Yes","")))`
- E4 `=IF(COUNTA(A4:B4)=0,"",IF(AND(ISNUMBER(A4),OR(B4="",ISNUMBER(B4))),IF(AND(ISNUMBER(A3),A4<A3),"Out of order","OK"),"Check"))`
- H3 `=IF(B2<>"OK","Check",IFERROR(INDEX($A$4:$A$369,MATCH("Yes",$D$4:$D$369,0)),"Not yet"))`
- H4 `=IF(ISNUMBER(H3),H3-$A$4+1,"")`; H5 `=IF(ISNUMBER(H3),INDEX($C$4:$C$369,MATCH("Yes",$D$4:$D$369,0)),"")`; H6 `=IF(ISNUMBER(H3),H5-$B$1,"")`
- B2 `=IF(ISNUMBER(B1),IF(B1>=0,"OK","Check"),"Check")`
- Alternative shown in the post (NOT in a cell): `=INDEX(A4:A369,COUNTIF(C4:C369,"<"&B1)+1)`.

Verified results (sample)
| Date | Sales | Running total |
|---|---|---|
| 09/10/2026 | 640.75 | 8,376.50 |
| 09/11/2026 | 1,510.00 | 9,886.50 |
| 09/12/2026 | 725.50 | 10,612.00 (first Yes) |
| 09/14/2026 | 1,175.00 | 12,767.00 |
H3 = 09/12/2026, H4 = 12, H5 = $10,612.00, H6 = $612.00, H7 = $12,767.00, H8 blank, H10 OK. Goal exactly equal to the 09/08 running total (6,415.75) gives 09/08/2026 (>= counts). Goal 0 gives 09/01/2026. Goal 99,999 gives Not yet with Still to go $87,232.00.

Pitfalls
- Deviation from the topic: the topic's `MATCH(TRUE,INDEX(C>=B1,0),0)` is not used. Empty rows hold "" and in Excel text compares greater than any number, so empty rows would count as TRUE; the built file matches the first "Yes" in column D instead (also easier to read and explain). Say so in the post; do not present the TRUE version as tested.
- COUNTIF alternative breaks with refunds (negative amounts): sales 1000, 700, -600, 300, 900 with goal 1500 gives the real first date day 2, but COUNTIF counts the day after the dip: day 4 in the test. Second test 2500, -1000, 400, 2000 with goal 2000: correct day 1, COUNTIF gives day 3.
- Text amounts are ignored by SUM, which is why column E flags them (Check). Blank sales days are fine (treated as 0). A date out of order flags Out of order, and the Days it took then turns unreliable.
- Goal blank, text or negative: B2 Check, H3 Check.
- Days it took counts the first row as day 1 and assumes the first row is the start of the period.
- Slots: 366 rows; the PDF print is 14 pages. 
- Fictional sales only; no TODAY.
