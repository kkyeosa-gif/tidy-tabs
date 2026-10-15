# Handoff for posts dated 2026-10-16

Status: topics (researcher, tasks/team-topics-2026-10-16.md) done; titles and openings (hook-writer) and the template builds (template-builder) are the next steps for the posts. Templates were built by the new functions in `scripts/build-templates.py` and recalculated in LibreOffice Calc (Linux, headless: `soffice --headless --convert-to xlsx` into a scratch dir, read back with openpyxl `data_only`) on 10/15/2026. Every result below was compared with an independent Python calculation. Edge cases were run by editing a copy of the template with openpyxl, recalculating it in LibreOffice and reading the values. No error values (#NAME?, Err:, #VALUE!, #N/A) appeared in any sheet except the one deliberate mismatched-range test noted under post 3. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. The functions used (WEEKDAY, MOD, SUMIFS, MIN, MAX, IF, AND, OR, ISNUMBER, COUNTIF, COUNTIFS, FORECAST, SLOPE, EDATE, DATEDIF, COUNTBLANK, SUMPRODUCT, CHOOSE) exist in both Excel and Google Sheets by name, but that was not run there. No tasks/search-stats.md exists: no search data. Sample data is fictional (say so once per post). PDF export from LibreOffice gave US Letter pages (portrait 612 x 792 pt, incomplete-rows Customers tab landscape 792 x 612 pt); the week ending Log tab (60 formula rows) prints on more than one page, the others one page per tab.

Only the five new functions were run for the committed files. Running the whole script rewrites every older .xlsx (new zip timestamps only), so those were restored with git checkout; git status shows only the five new .xlsx plus the edited script and task notes.

Titles carry "Excel and Google Sheets" and have no colon, dash or hyphen. Cell addresses below are the real ones in the built files.

---

## 1. 2026-10-16-team-01-week-ending-friday-date-weekday-excel-google-sheets
- Title: (hook-writer) example: How to Find the Week Ending Date for Any Date in Excel and Google Sheets
- Template: `templates/tidy-tabs-week-ending-date.xlsx` (function `week_ending_date`). Sheets: Log, Weeks, How to use (portrait).
- Inputs (yellow): Log A2:C60 (date, client, hours; 12 sample rows 09/28/2026 to 10/14/2026); Log G2 week ending weekday number = 6 (1 Sunday ... 7 Saturday; Friday), whole number 1 to 7 validation.
- Formulas (Log filled to row 60):
  - Log D2 `=IF(ISNUMBER(A2),A2+MOD($G$2-WEEKDAY(A2,1),7),"")`; G3 shows the weekday name with CHOOSE; G4 `=SUM(C2:C60)`.
  - Weeks A2 `=IF(COUNT(Log!D2:D60)=0,"",MIN(Log!D2:D60))`; A3 `=IF(A2="","",IF(A2+7>MAX(Log!D2:D60),"",A2+7))` filled to row 27; B2 `=IF(A2="","",SUMIFS(Log!C2:C60,Log!D2:D60,A2))`.
  - Weeks E2 `=SUM(B2:B27)`; E3 `=SUM(Log!C2:C60)`; E4 `=IF(ROUND(E2-E3,6)=0,"OK","CHECK")`.
- Verified values (LibreOffice Calc only): 09/28 to 10/02 give 10/02/2026; 10/03 to 10/09 give 10/09/2026 (10/09 is a Friday and stays); 10/12 and 10/14 give 10/16/2026. Weeks: 10/02/2026 21.5 h, 10/09/2026 17.75 h, 10/16/2026 7.75 h, total 47; E4 OK. The 12 week ending dates and the three weekly sums match Python weekday math.
- Edge: date already on the week ending day (10/02, 10/09) stays; the day after (Saturday 10/03) jumps 6 days to 10/09.
- Edge: Sunday 10/18/2026 gives 10/23/2026; 12/30/2026 (Wednesday) gives 01/01/2027 (year boundary), Weeks list shows 01/01/2027 with 2 h and 0 h rows for the empty weeks in between; matches Python.
- Edge: G2 changed to 1 (week ends Sunday): 10/04 25 h, 10/11 14.25 h, 10/18 7.75 h; G3 shows Sunday; matches Python.
- Edge: date typed as text '10/20/2026': D blank, hours left out of Weeks, E2 47 vs E3 50, E4 CHECK.
- Edge: hours typed as text '2.5': SUM and SUMIFS both skip it (44.5 on both tabs), E4 still OK, so the drop is silent.
- Edge: blank date row: D blank.
- Limits: date math only, no pay period, overtime or time zone rules; up to 26 weeks listed on Weeks; weeks with no hours inside the range show 0; text hours are dropped quietly.

---

## 2. 2026-10-16-team-02-free-shipping-threshold-amount-to-add-if-and-excel-google-sheets
- Title: (hook-writer) example: How to Check a Free Shipping Threshold and Show the Amount to Add in Excel and Google Sheets
- Template: `templates/tidy-tabs-free-shipping-threshold.xlsx` (function `free_shipping_threshold`). Sheets: Orders, How to use (portrait).
- Inputs (yellow): Orders A2:C31 (order, subtotal, delivery Ship or Pickup with a dropdown; 10 sample rows); I2 threshold $50.00; I3 close gap $10.00.
- Formulas (filled to row 31):
  - D2 `=IF(OR(NOT(ISNUMBER(B2)),C2=""),"",IF(C2="Pickup","Pickup",IF(AND(C2="Ship",B2>=$I$2),"Free","Pay")))`; E2 `=IF(OR(D2="Pay",D2="Free"),MAX(0,ROUND($I$2-B2,2)),"")`.
  - I4 `=COUNTIF(D2:D31,"Free")`; I5 `=COUNTIF(D2:D31,"Pay")`; I6 `=COUNTIFS(D2:D31,"Pay",E2:E31,"<="&I3)`; I7 `=COUNTIF(D2:D31,"Pickup")`; I8 `=IF(I4+I5+I7=COUNTA(A2:A31),"OK","CHECK")`.
- Verified values (LibreOffice Calc only), status and amount to add: ORD-5001 $32.50 Pay $17.50; 5002 $50.00 Free $0.00 (exactly at the threshold); 5003 $49.99 Pay $0.01; 5004 $78.25 Free; 5005 $24.00 Pickup; 5006 $45.00 Pay $5.00; 5007 $112.40 Free; 5008 $18.75 Pay $31.25; 5009 $61.00 Pickup; 5010 $41.50 Pay $8.50. I4 3, I5 5, I6 3 (5003, 5006, 5010), I7 2, I8 OK. All match Python Decimal comparison.
- Edge: threshold 45 and gap 5: Free 5, Pay 3, close 1, Pickup 2 (Python agrees).
- Edge: an 11th order with subtotal typed as text '55', or a blank subtotal, or a blank delivery cell: that row gets no status, no amount, and I8 shows CHECK.
- Edge: Pickup rows show Pickup and no amount even when above the threshold.
- Limits: own threshold only, no carrier rates or marketplace rules; ROUND(...,2) used so 50 minus 49.99 shows exactly 0.01; the delivery word must be exactly Ship or Pickup (dropdown provided).

---

## 3. 2026-10-16-team-03-forecast-next-month-sales-forecast-linear-excel-google-sheets
- Title: (hook-writer) example: How to Forecast Next Month's Sales With FORECAST in Excel and Google Sheets
- Template: `templates/tidy-tabs-forecast-next-month.xlsx` (function `forecast_next_month`). Sheets: Sales, How to use (portrait).
- Inputs (yellow): Sales A2:C25 (month number, month 01/01 of each month shown as mmm yyyy, sales; 8 sample rows Jan to Aug 2026: 1200, 1350, 1500, 1480, 1620, 1710, 1690, 1850).
- Formulas:
  - E2 `=IF(COUNT($A$2:$A$25)=0,"",MAX($A$2:$A$25)+1)` (E3 +2, E4 +3); F2 `=IF(OR(E2="",COUNT($B$2:$B$25)=0),"",EDATE(MAX($B$2:$B$25),E2-MAX($A$2:$A$25)))`; G2 `=IF(OR(E2="",COUNT($C$2:$C$25)<2),"",FORECAST(E2,$C$2:$C$25,$A$2:$A$25))` (G3, G4 same with E3, E4).
  - G6 `=IF(COUNT($C$2:$C$25)<2,"",SLOPE($C$2:$C$25,$A$2:$A$25))`; G7 `=COUNT($C$2:$C$25)`; G8 hand calculation of the same first projection from SUMPRODUCT, SUMSQ and AVERAGE; G9 `=IF(G2="","",IF(ROUND(G2-G8,6)=0,"OK","CHECK"))`.
- Verified values (LibreOffice Calc only): G2 $1,926.07 (1926.0714), G3 $2,009.64, G4 $2,093.21 for Sep, Oct, Nov 2026 (months 9, 10, 11); slope $83.57 per month (83.5714); G7 8; G8 equals G2; G9 OK. A Python least squares fit gives the same to 10 digits.
- Edge: flat sales (all 1000): slope 0, forecasts 1000. Falling sales (2000 down to 1300): slope -98.8095, forecasts 1217.86, 1119.05, 1020.24; match Python.
- Edge: only 2 months (1200, 1350): forecasts 1500, 1650, 1800, slope 150.
- Edge: a month number typed as text, a blank month number, or a blank sales cell in the middle: LibreOffice skips that pair, and the forecast equals the Python fit on the remaining pairs (for example 1928.39 for month 9); G7 counts the sales cells, so G9 shows CHECK because the hand calculation assumes complete pairs.
- Edge: mismatched range lengths (a test formula `=FORECAST(9,C2:C9,A2:A8)` added to a copy): LibreOffice returned #VALUE!, not #N/A. Excel was not run; if the post shows #N/A it must come from a Microsoft help page, not from us.
- Edge: all inputs cleared: forecast cells, slope and check blank, G7 0.
- Limits: a straight-line estimate from the past, not a prediction, no seasons or promotions; needs at least 2 months; uses the older FORECAST name (FORECAST.LINEAR is the newer name, not run); month numbers should be in order with no gaps.

---

## 4. 2026-10-16-team-04-client-tenure-years-months-datedif-excel-google-sheets
- Title: (hook-writer) example: How to Calculate How Long a Client Has Been With You in Years and Months in Excel and Google Sheets
- Template: `templates/tidy-tabs-client-tenure.xlsx` (function `client_tenure`). Sheets: Clients, How to use (portrait).
- Inputs (yellow): Clients A2:B30 (client, first project date; 10 sample rows); I2 as of date 10/16/2026 (typed); I3 years threshold 3.
- Formulas (filled to row 30):
  - C2 `=IF(B2="","",IF(NOT(ISNUMBER(B2)),"Check date",IF(B2>$I$2,"Check date",DATEDIF(B2,$I$2,"y"))))`; D2 `=IF(ISNUMBER(C2),DATEDIF(B2,$I$2,"ym"),"")`; E2 `=IF(ISNUMBER(C2),C2&" yr "&D2&" mo","")`; F2 `=IF(ISNUMBER(C2),C2*12+D2,"")`.
  - I4 `=COUNT(C2:C30)`; I5 `=COUNTIF(C2:C30,">="&I3)`; I6 `=IF(COUNT(F2:F30)=0,"",MAX(F2:F30))`; I7 `=COUNTIF(C2:C30,"Check date")`; I8 `=IF(I4+I7=COUNTA(A2:A30),"OK","CHECK")`.
- Verified values (LibreOffice Calc only), as of 10/16/2026: Maple Street Bakery (03/15/2021) 5 yr 7 mo, 67 months; Harbor Yoga (11/02/2022) 3 yr 11 mo, 47; Birch Design (02/29/2024) 2 yr 7 mo, 31; Lantern Print (08/20/2025) 1 yr 1 mo, 13; Cedar Row (10/16/2023) 3 yr 0 mo, 36; Willow Pet (06/30/2019) 7 yr 3 mo, 87; Granite Fitness (01/05/2026) 0 yr 9 mo, 9; Juniper Books (12/31/2020) 5 yr 9 mo, 69; Copper Kettle (10/01/2026) 0 yr 0 mo, 0; Pine Hill Dental (04/09/2018) 8 yr 6 mo, 102. I4 10, I5 6, I6 102, I7 0, I8 OK. All match the Python month difference (months between, minus 1 if the start day is later in the month than the as of day).
- Edge: start day later in the month than the as of day (08/20/2025 on 10/16/2026) gives 1 yr 1 mo, not 1 yr 2 mo; same day exactly (start 10/15/2023, as of 10/15/2026) gives 3 yr 0 mo.
- Edge: Feb 29 start (02/29/2024) with as of 02/28/2027: 2 yr 11 mo (35 months), matches Python; as of 10/16/2026 gives 2 yr 7 mo.
- Edge: start after the as of date (01/01/2027): C shows Check date, no #NUM!, I7 1, I8 OK. A date typed as text '03/15/2020': Check date.
- Edge: blank row removed: I4 9, I5 5. I3 changed to 5: I5 4 (Python 4).
- Not run: the raw `=DATEDIF(B2,$I$2,"y")` with a start after the as of date was not evaluated on its own, so the post must not claim #NUM! came from our run (the Troubleshooting case can link Microsoft's DATEDIF page). The template guards it before DATEDIF is called.
- Limits: DATEDIF counts complete months only; it is not in Excel's function picker and not run in Excel; the as of date is typed (swap in `=TODAY()` for a live count); no business days.

---

## 5. 2026-10-16-team-05-find-incomplete-customer-rows-countblank-excel-google-sheets
- Title: (hook-writer) example: How to Find Rows With Missing Information in a Customer List in Excel and Google Sheets
- Template: `templates/tidy-tabs-incomplete-rows.xlsx` (function `incomplete_rows`). Sheets: Customers (landscape), How to use (portrait).
- Inputs (yellow): Customers A2:E30 (name, email, phone, street, ZIP; ZIP column formatted as text; 12 sample rows).
- Formulas (filled to row 30):
  - F2 `=IF(COUNTA(A2:E2)=0,"",COUNTBLANK(A2:E2))`; G2 `=IF(F2="","",IF(F2=0,"Complete","Missing"))`; H2 `=IF(OR(F2="",F2=0),"",IF(A2="","Name",IF(B2="","Email",IF(C2="","Phone",IF(D2="","Street","ZIP")))))`.
  - K2 `=COUNT(F2:F30)`; K3 `=COUNTIF(G2:G30,"Complete")`; K4 `=COUNTIF(G2:G30,"Missing")`; K5 `=SUMPRODUCT((F2:F30<>"")*(E2:E30=""))`; K6 `=IF(SUM(F2:F30)=COUNT(F2:F30)*5-COUNTA(A2:E30),"OK","CHECK")`.
- Verified values (LibreOffice Calc only), blanks and first missing: Ava Thompson 0 Complete; Ben Ortiz 1 Email; Carla Nguyen 1 Phone; Dev Patel 1 Street; Erin Walsh 1 ZIP; Frank Lee 2 Email; Gina Rossi 0; (no name) hal.k 1 Name; Iris Chen 0; Jon Baker 4 Email; Kim Alvarez 0; Liam Burke 1 Phone. K2 12, K3 4, K4 8, K5 2 (Erin Walsh, Jon Baker), K6 OK. ZIPs 02108, 07030, 02210 kept their leading zeros as text. All match the Python blank count.
- Edge: a cell holding only a space (B3): counted as filled, row 3 shows 0 Complete, K3 5, K4 7 (Python agrees); this is the Troubleshooting case.
- Edge: a cell with `=""` (B3): LibreOffice counted it as blank (F3 1, Missing); K6 shows CHECK because COUNTA treats the formula cell as filled.
- Edge: an all-blank row in the middle (row 14) stays blank (no 5), a fully complete row after it shows 0 Complete; K2 13, K3 5, K4 8.
- Edge: all rows cleared: K2 to K5 0, K6 OK.
- Limits: counts empty cells only, no check that an email or ZIP is valid; a row with no data at all is treated as unused and not flagged; first missing field follows the fixed order Name, Email, Phone, Street, ZIP; a ZIP typed as a number (7030) loses its leading zero (the column is formatted as text, so type or paste as text); Excel and Google Sheets behavior of COUNTBLANK with `""` was not run.
