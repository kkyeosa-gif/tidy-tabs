# Handoff for posts dated 2026-10-15

Status: topics (researcher), titles and openings (hook-writer) and templates (template-builder) done. Templates were built by the new functions in `scripts/build-templates.py` and recalculated in LibreOffice Calc (Linux, headless: `soffice --headless --convert-to xlsx` into a scratch dir, read back with openpyxl `data_only`) on 10/14/2026. Every result below was compared with Python math. Edge cases listed were run by editing a copy of the template with openpyxl, recalculating it in LibreOffice and reading the values. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. No tasks/search-stats.md exists: no search data. Sample data is fictional (say so once per post). Each template tab is set to US Letter, fit to one page wide (landscape for the first three); a PDF export from LibreOffice gave one Letter page per tab (landscape 792 x 612 pt, portrait 612 x 792 pt).

Titles carry "Excel and Google Sheets", name the task, and have no colon, dash or hyphen. Openings give the exact formula first. Cell addresses below are the real ones in the built files. Only the five new functions were run (older templates untouched; git status shows only the five new .xlsx files plus the edited script and task notes).

---

## 1. 2026-10-15-team-01-average-days-to-get-paid-excel-google-sheets
- Title: How to Calculate Average Days to Get Paid in Excel and Google Sheets
- Opening: Subtract the invoice date from the paid date with `=D2-C2`, then average the results with `=AVERAGE(E2:E11)`. For the eight paid sample invoices, clients took 25.5 days on average, from 7 days at the fastest to 45 at the slowest.
- Template: `templates/tidy-tabs-days-to-get-paid.xlsx` (function `days_to_get_paid`). Sheets: Invoices (landscape), How to use.
- Inputs (yellow): Invoices A2:D30 (invoice, client, invoice date, paid date; 10 sample rows, 2 unpaid); J2 payment terms in days = 30.
- Formulas (filled to row 30):
  - E2 `=IF(OR(C2="",D2=""),"",D2-C2)`; F2 `=IF(C2="","",IF(D2="","Unpaid","Paid"))`; E turns red when above J2.
  - J4 `=IFERROR(AVERAGE(E2:E30),"")`; J5 `=IF(COUNT(E2:E30)=0,"",MIN(E2:E30))`; J6 same with MAX; J7 `=COUNT(E2:E30)`; J8 `=COUNTIF(F2:F30,"Unpaid")`; J9 `=COUNTIF(E2:E30,">"&J2)`; J10 `=IF(J7+J8=COUNTA(A2:A30),"OK","CHECK")`.
- Verified values (LibreOffice Calc only), days to pay: INV-2001 20, 2002 35, 2003 14, 2004 35, 2005 7, 2006 45, 2007 20, 2008 28; 2009 and 2010 blank and Unpaid. J4 25.5, J5 7, J6 45, J7 8, J8 2, J9 3, J10 OK. Python: 204 days over 8 invoices = 25.5, min 7, max 45, 3 over 30. Match.
- Edge: INV-2009 given a paid date 09/20/2026: E 4, Paid, J4 23.1111 (208/9 matches Python), J7 9, J8 1, J10 OK.
- Edge: all paid dates cleared: J4, J5, J6 blank, J7 0, J8 10, J10 OK.
- Edge: paid date before invoice date (07/01 vs 07/08): E -7, J4 20.25, J5 -7 (a negative result is not flagged; the notes say to fix the typo).
- Edge: paid date typed as text "07/21/2026": LibreOffice converted it and returned 20, same as the real date; this depends on the app and locale, so tell readers to enter real dates.
- Edge: paid date with no invoice date (row 12): E and F blank, J10 CHECK.
- Limits: looks back at paid invoices only (not an aging report); calendar days, not business days; one average for all clients (no per client split).

---

## 2. 2026-10-15-team-02-early-payment-discount-2-10-net-30-excel-google-sheets
- Title: How to Calculate a 2/10 Net 30 Early Payment Discount in Excel and Google Sheets
- Opening: Find the discount with `=ROUND(D2*$M$2,2)` and the last day to take it with `=C2+$M$3`, where M2 is 2% and M3 is 10 days. On a $1,250.00 invoice dated 09/01/2026, the client saves $25.00 by paying $1,225.00 on or before 09/11/2026.
- Template: `templates/tidy-tabs-early-payment-discount.xlsx` (function `early_payment_discount`). Sheets: Discounts (landscape), How to use.
- Inputs (yellow): Discounts A2:E25 (invoice, client, invoice date, amount, paid date; 8 sample rows, 3 unpaid); M2 rate 2%, M3 discount days 10, M4 net days 30.
- Formulas (filled to row 25):
  - F2 `=IF(C2="","",C2+$M$3)`; G2 `=IF(OR(C2="",NOT(ISNUMBER(D2))),"",ROUND(D2*$M$2,2))`; H2 `=IF(G2="","",D2-G2)`; I2 `=IF(C2="","",C2+$M$4)`.
  - J2 `=IF(OR(C2="",G2=""),"",IF(E2="","Open",IF(E2<=F2,"Discount earned","Full amount")))`; Full amount is red, Discount earned is green.
  - M6 `=SUMIF(J2:J25,"Discount earned",G2:G25)`; M7 same for "Full amount"; M8 same for "Open"; M9 `=COUNTIF(J2:J25,"Open")`.
- Verified values (LibreOffice Calc only), discount, pay by deadline, status: INV-3001 $25.00, $1,225.00, earned (deadline 09/11/2026). 3002 $9.60, $470.40, Full amount (paid 09/20, deadline 09/13). 3003 $44.00, $2,156.00, earned (paid 09/18, the deadline day). 3004 $16.71, $818.79, Full amount. 3005 $19.20, $940.80, earned. 3006 $6.25, $306.25, Open. 3007 $29.60, $1,450.40, Open. 3008 $14.51, $710.89, Open (14.508 rounds to 14.51). M6 $88.20, M7 $26.31, M8 $50.36, M9 3. Full due dates 10/01, 10/03, 10/08, 10/15, 10/22, 10/31, 11/04, 11/08/2026. All match Python (Decimal round half up).
- Edge: INV-3001 paid on the deadline day 09/11/2026: Discount earned. Paid 09/12/2026: Full amount, M6 $63.20, M7 $51.31 (match Python).
- Edge: amount typed as text "1,250.00": G, H, J blank (row ignored).
- Edge: terms changed to 1%, 15 days, 45 net days: INV-3001 deadline 09/16/2026, discount $12.50, pay $1,237.50, due 10/16/2026, earned; INV-3002 (paid 09/20, deadline 09/18) Full amount.
- Edge: all rows cleared: row 2 blank, M6 to M9 all 0.
- Limits: calendar days only; terms of the form discount%/days net days, one set for the whole sheet; no end of month terms; open invoices are not marked as expired once the deadline passes. Arithmetic from typed terms; the contract decides what a client owes.

---

## 3. 2026-10-15-team-03-sell-through-rate-inventory-excel-google-sheets
- Title: How to Calculate Sell Through Rate in Excel and Google Sheets
- Opening: Divide units sold by units received with `=C2/B2` and format the cell as a percent. Selling 41 of 48 lavender candles is an 85.4% sell through rate.
- Template: `templates/tidy-tabs-sell-through-rate.xlsx` (function `sell_through_rate`). Sheets: Products (landscape), How to use.
- Inputs (yellow): Products A2:C30 (product, units received, units sold; 8 sample rows); I2 Strong at or above 70%, I3 Slow below 40% (examples, not industry rules).
- Formulas (filled to row 30):
  - D2 `=IF(OR(B2="",C2=""),"",B2-C2)`; E2 `=IF(OR(NOT(ISNUMBER(B2)),NOT(ISNUMBER(C2))),"",IF(B2<=0,"",C2/B2))`; F2 `=IF(E2="","",IF(E2>1,"Check counts",IF(E2>=$I$2,"Strong",IF(E2<$I$3,"Slow","OK"))))`.
  - I5 `=SUM(B2:B30)`; I6 `=SUM(C2:C30)`; I7 `=IF(I5=0,"",I6/I5)`; I8 `=COUNTIF(F2:F30,"Slow")`; I9 `=SUMIF(F2:F30,"Slow",D2:D30)`.
- Verified values (LibreOffice Calc only), rate and status: Lavender candle 41/48 85.4% Strong; Cedar candle 30/48 62.5% OK; Oatmeal soap 96/120 80.0% Strong; Charcoal soap 21/60 35.0% Slow; Brass hoop earrings 12/30 40.0% OK (exactly at the Slow cutoff, not below it); Birthday card 78/100 78.0% Strong; Linen tote 25/25 100.0% Strong; Gift tag set 18/80 22.5% Slow. I5 511, I6 321, I7 62.818% (62.8%), I8 2, I9 101 (39 + 62). All match Python.
- Edge: units received set to 0 (Gift tag set): E and F blank, D shows -18, overall rate becomes 321/431 = 74.5%.
- Edge: sold 60 of 48 (Cedar candle): rate 125.0%, F "Check counts", D -12.
- Edge: sold typed as text "96": E and F blank; D still 24 (LibreOffice coerces text in subtraction) but SUM skips it, so I6 225 and I7 44.0%. The row is quietly dropped from the totals.
- Edge: sold cleared: D, E, F blank. All cleared: I5, I6, I8, I9 0 and I7 blank.
- Limits: units only, not dollars; same time period for every product; cutoffs are the reader's own; a received count of 0 or text sold shows no warning flag, only blanks.

---

## 4. 2026-10-15-team-04-how-many-boxes-needed-roundup-excel-google-sheets
- Title: How to Calculate How Many Boxes You Need in Excel and Google Sheets
- Opening: Get the box count with `=ROUNDUP(B2/C2,0)`, where B2 is units to ship and C2 is units per box. Shipping 50 candles at 12 per box needs 5 boxes, not 4, because 50 divided by 12 is 4.17.
- Template: `templates/tidy-tabs-boxes-needed.xlsx` (function `boxes_needed`). Sheets: Orders, How to use (both portrait).
- Inputs (yellow): Orders A2:C30 (item, units to ship, units per box; 7 sample rows); J2 cost per box $1.45.
- Formulas (filled to row 30):
  - D2 `=IF(AND(ISNUMBER(B2),ISNUMBER(C2),C2>0),ROUNDUP(B2/C2,0),"")`; E2 `=IF(OR(D2="",B2=0),"",B2-(D2-1)*C2)`; F2 `=IF(D2="","",D2*C2-B2)`; G2 `=IF(D2="","",D2*$J$2)`.
  - J4 `=SUM(D2:D30)`; J5 `=SUM(G2:G30)`; J6 `=SUM(F2:F30)`.
- Verified values (LibreOffice Calc only), boxes, units in last box, empty slots, box cost: Soy candle 50 at 12 5, 2, 10, $7.25. Soap 30 at 8 4, 6, 2, $5.80. Earring cards 17 at 20 1, 17, 3, $1.45. Mugs 26 at 6 5, 2, 4, $7.25. Tote bags 48 at 12 4, 12, 0, $5.80. Birthday cards 130 at 25 6, 5, 20, $8.70. Gift tag sets 0 at 40: 0 boxes, last box blank, 0 slots, $0.00. J4 25 boxes, J5 $36.25, J6 39 slots. All match Python math.ceil.
- Edge: units per box 0 or blank: D to G blank, totals drop that row (J4 20). 
- Edge: units typed as text "30": D to G blank, J4 21.
- Edge: 48 at 12 (exact multiple): 4 boxes, last box 12, 0 empty.
- Edge: 10.5 units at 5 per box: 3 boxes, last box 0.5, 4.5 empty (decimals accepted; real units should be whole).
- Edge: all rows cleared: D blank, J4, J5, J6 0.
- Limits: one box size per row; no weight, dimension or carrier checks; text input drops the row from totals with no warning.

---

## 5. 2026-10-15-team-05-median-order-value-vs-average-excel-google-sheets
- Title: How to Find the Median Order Value in Excel and Google Sheets
- Opening: Use `=MEDIAN(C2:C100)` for the middle order and `=AVERAGE(C2:C100)` for the mean. In the sample, one $640.00 wholesale order lifts the average to $83.35 while the median stays at $35.25.
- Template: `templates/tidy-tabs-median-order-value.xlsx` (function `median_order_value`). Sheets: Orders, How to use (both portrait).
- Inputs (yellow): Orders A2:C100 (order number, date 09/01/2026 to 09/25/2026 every other day, total; 13 sample rows); F2 cutoff $500.00.
- Formulas: F4 `=COUNT(C2:C100)`; F5 `=IFERROR(AVERAGE(C2:C100),"")`; F6 `=IFERROR(MEDIAN(C2:C100),"")`; F7 `=IF(F5="","",F5-F6)`; F8 `=IF(F4=0,"",MAX(C2:C100))`; F9 `=IFERROR(AVERAGEIF(C2:C100,"<="&F2),"")`; F10 `=COUNTIF(C2:C100,">"&F2)`.
- Verified values (LibreOffice Calc only): F4 13, F5 $83.35 (83.3461538), F6 $35.25, F7 $48.10 (48.0961538), F8 $640.00, F9 $36.96 (36.9583333), F10 1. Python mean, statistics.median and filtered mean give the same.
- Edge: 14th order of $50.00 added (even count): F4 14, F5 80.9643, F6 36.625 (average of the two middle orders $35.25 and $38.00); Python matches.
- Edge: one order cleared (blank row): F4 12, F5 88.7917, F6 36.625; matches Python.
- Edge: the $640.00 stored as text "640.00": skipped, F4 12, F5 36.9583, F6 34.50, F8 $62.00; matches Python and shows the silent drop.
- Edge: one order of $40: F5 and F6 $40.00, F7 $0.00. All cleared: F4 0, F5 to F9 blank, F10 0.
- Limits: totals only; no refund, tax or shipping handling; text totals are skipped with no warning (compare F4 with the number of rows); the $500 cutoff is the reader's choice.
