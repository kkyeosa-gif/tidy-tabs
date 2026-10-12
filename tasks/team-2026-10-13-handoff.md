# Handoff for posts dated 2026-10-13

Status: topics (researcher), titles and openings (hook-writer) and templates (template-builder) done. Templates are built by `scripts/build-templates.py` and were recalculated in LibreOffice Calc 24.2 (Linux, headless: `soffice --headless --convert-to xlsx` into a scratch dir, read back with openpyxl `data_only`) on 10/12/2026. Every result below was compared with Python math. Edge cases listed were run by editing a copy of the template with openpyxl, recalculating it in LibreOffice and reading the values. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. No tasks/search-stats.md exists: no search data. Sample data is fictional (say so once per post). Each template tab is set to US Letter, fit to one page wide (landscape for invoice due dates and supplier price compare); the PDF export from LibreOffice confirmed Letter size, with the daily sales tab running two pages tall.

Titles carry "Excel and Google Sheets", name the task, and have no colon, dash or hyphen. Openings give the exact formula first. Cell addresses below are the real ones in the built files. Only the five new functions were run (older templates untouched). Note: the 10-12 template functions are not in scripts/build-templates.py (see tasks/briefs.md), so a full run of the script would not recreate those five files.

---

## 1. 2026-10-13-team-01-invoice-due-date-payment-terms-excel-google-sheets
- Title: How to Calculate an Invoice Due Date From Payment Terms in Excel and Google Sheets
- Opening: Add the days in the terms to the invoice date with `=C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE)`, where C2 is the invoice date and J2:K6 lists each terms name next to its number of days. An invoice dated 09/01/2026 on Net 30 terms is due 10/01/2026.
- Template: `templates/tidy-tabs-invoice-due-dates.xlsx`. Sheets: Due dates (landscape), How to use.
- Inputs (yellow): terms table J2:K6 (Due on receipt 0, Net 15, Net 30, Net 45, Net 60); as-of date K8 = 10/13/2026 (typed so the sample does not change; the notes say `=TODAY()` works). Invoices A2:D9 (8 rows); D has a terms dropdown for rows 2 to 20.
- Formulas (filled to row 20):
  - E2 `=IF(D2="","",IFERROR(VLOOKUP(D2,$J$2:$K$6,2,FALSE),"Check terms"))`
  - F2 `=IF(OR(C2="",NOT(ISNUMBER(E2))),"",C2+E2)`
  - G2 `=IF(F2="","",F2-$K$8)`
  - H2 `=IF(G2="","",IF(G2<0,"Past due",IF(G2=0,"Due today","Open")))`; Past due cells are red.
  - K10 `=COUNTIF(H2:H20,"Open")+COUNTIF(H2:H20,"Due today")`; K11 `=COUNTIF(H2:H20,"Past due")`.
- Verified values (LibreOffice Calc 24.2 only), due date, days until due, status:
  - INV-1001 09/01/2026 Net 30: 10/01/2026, -12, Past due. INV-1002 09/15 Net 15: 09/30/2026, -13, Past due. INV-1003 09/28 Net 45: 11/12/2026, 30, Open. INV-1004 10/05 Net 30: 11/04/2026, 22, Open. INV-1005 10/13 Due on receipt: 10/13/2026, 0, Due today. INV-1006 10/01 Net 60: 11/30/2026, 48, Open. INV-1007 08/20 Net 60: 10/19/2026, 6, Open. INV-1008 10/10 Net 15: 10/25/2026, 12, Open. K10 6, K11 2. All match Python date math. Rows 10 to 20 stay blank.
  - Edge: terms "Net 20" (not in table): E "Check terms", F, G, H blank. Terms in lowercase ("net 30"): still found (VLOOKUP ignores case), 30 days.
  - Edge: invoice date 01/31/2026 plus Net 30: 03/02/2026 (days are added, not months).
  - Edge: invoice date missing: F, G, H blank. Invoice date stored as text "10/01/2026": LibreOffice converted it and returned 10/31/2026; this depends on the app and locale, so tell readers to enter real dates.
  - Edge: as-of date K8 moved to 11/01/2026: INV-1001, 1002, 1005, 1007 and 1008 become Past due (5 of the 8 original rows).
- Limits: due date is not moved for weekends or holidays (WORKDAY would); no early-pay discount terms such as 2/10 net 30 and no end-of-month terms. Date arithmetic only: the real terms are in the reader's contract.

---

## 2. 2026-10-13-team-02-sales-by-day-of-week-sumif-excel-google-sheets
- Title: How to Total Sales by Day of the Week in Excel and Google Sheets
- Opening: Name each date's weekday with `=TEXT(A2,"ddd")`, then average one day with `=AVERAGEIF($C$2:$C$29,"Sat",$B$2:$B$29)`. In the sample month, Saturdays average $159.19 and Sundays $229.25.
- Template: `templates/tidy-tabs-sales-by-weekday.xlsx`. Sheets: Daily sales, By weekday, How to use.
- Data: Daily sales A2:B29 holds 28 days from 09/01/2026 (a Tuesday) to 09/28/2026, so each weekday appears 4 times. Helper formulas fill C2:C60.
- Formulas:
  - Daily sales!C2 `=IF(A2="","",TEXT(A2,"ddd"))`
  - By weekday (A2:A8 Mon to Sun): B2 `=COUNTIF('Daily sales'!$C$2:$C$200,A2)`; C2 `=SUMIF('Daily sales'!$C$2:$C$200,A2,'Daily sales'!$B$2:$B$200)`; D2 `=IF(B2=0,"",C2/B2)`; E2 `=IF($C$9=0,"",C2/$C$9)`; F2 `=IF(D2="","",IF(D2=MAX($D$2:$D$8),"Best day",""))`.
  - B9 `=SUM(B2:B8)`, C9 `=SUM(C2:C8)`; F11 `=IF(B9=COUNT('Daily sales'!A2:A200),"OK","CHECK")`; F12 `=IF(ROUND(C9-SUM('Daily sales'!B2:B200),2)=0,"OK","CHECK")`.
- Verified values (LibreOffice Calc 24.2 only), total and average per day: Mon $378.50, $94.63; Tue $165.50, $41.38; Wed $150.75, $37.69; Thu $202.75, $50.69; Fri $265.25, $66.31; Sat $636.75, $159.19; Sun $917.00, $229.25 (Best day). Days counted 4 each, B9 28, C9 $2,716.50, F11 OK, F12 OK. Unrounded averages 94.625, 41.375, 37.6875, 50.6875, 66.3125, 159.1875, 229.25 match Python.
- Edge: a date stored as text ("09/29/2026" in A31, $100.00): LibreOffice still named it Tue, so Tue shows 5 days and $265.50; B9 30 does not equal COUNT of real dates 29, so F11 shows CHECK (F12 OK).
- Edge: a 09/30/2026 row with $0.00 sales counts as a day (Wed row gets a 5th day with no added sales, average falls).
- Edge: all rows cleared: counts and totals 0, averages and shares blank, no Best day, checks OK.
- Limits: weekday names are English (Mon to Sun) from a US-English file; a locale with other day names would not match the names in A2:A8. One row per date is assumed. Ranges stop at row 200 (helper formulas at row 60).

---

## 3. 2026-10-13-team-03-repeat-customer-rate-countif-excel-google-sheets
- Title: How to Calculate Repeat Customer Rate in Excel and Google Sheets
- Opening: Count orders per customer with `=COUNTIF(Orders!$B$2:$B$200,A2)`, then divide repeat customers by customers who ordered with `=COUNTIF(B2:B9,">=2")/COUNTIF(B2:B9,">0")`. In the sample, 3 of 7 customers ordered more than once, a 42.9% repeat rate.
- Template: `templates/tidy-tabs-repeat-customers.xlsx`. Sheets: Orders, Customers, How to use.
- Data: Orders A2:D14 (13 orders, 08/04/2026 to 10/09/2026, customer in column B). Customers A2:A9 lists 8 names once each; formulas fill rows 2 to 31.
- Formulas (Customers tab):
  - B2 `=IF(A2="","",COUNTIF(Orders!$B$2:$B$200,A2))`; C2 `=IF(A2="","",IF(B2=0,"No orders",IF(B2=1,"One time","Repeat")))`; D2 `=IF(A2="","",SUMIF(Orders!$B$2:$B$200,A2,Orders!$D$2:$D$200))`.
  - G2 `=COUNTIF(B2:B31,">0")`; G3 `=COUNTIF(B2:B31,">=2")`; G4 `=IF(G2=0,"",G3/G2)`; G5 `=SUM(B2:B31)`; G6 `=COUNTA(Orders!B2:B200)`; G7 `=IF(G5=G6,"OK","CHECK")`; G8 `=SUMIF(C2:C31,"Repeat",D2:D31)`; G9 `=IF(SUM(D2:D31)=0,"",G8/SUM(D2:D31))`.
- Verified values (LibreOffice Calc 24.2 only): orders per customer Ana Ruiz 3 ($94.50), Ben Ortiz 1 ($65.50), Cara Lee 2 ($72.00), Dev Patel 1 ($54.25), Eli Brooks 4 ($113.75), Fay Nolan 1 ($12.00), Gus Hall 0 ($0.00, No orders), Hana Cole 1 ($90.00). G2 7, G3 3, G4 42.857% (shown as 42.9%), G5 13, G6 13, G7 OK, G8 $280.25, G9 55.8%. All match Python.
- Edge: typos in Orders (B2 "Ana Ruis", B3 "eli brooks", B4 "Ben Ortiz " with a trailing space): Ana 2, Ben 0, Eli 4 (capitals ignored), G2 6, G3 3, G4 50.0%, G5 11 against G6 13, G7 CHECK.
- Edge: Orders emptied: G2 0, G3 0, G4 and G9 blank, G7 OK.
- Edge: new customer added in A10 with one order: B10 1, C10 One time, G2 8, G3 3, G7 OK.
- Limits: counts only the orders pasted in, so compare rates over the same date range; says what share ordered again, not why. Each customer must be spelled the same in both tabs. Customers list is typed by hand (no unique-list formula).

---

## 4. 2026-10-13-team-04-cheapest-supplier-cost-per-ounce-index-match-excel-google-sheets
- Title: How to Find the Cheapest Supplier by Cost Per Ounce in Excel and Google Sheets
- Opening: Divide price plus shipping by package size with `=(B2+D2)/C2`, then return the cheapest name with `=INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0))`. Of five fictional wax quotes, Harbor Craft Wholesale wins at $0.2667 per ounce.
- Template: `templates/tidy-tabs-supplier-price-compare.xlsx`. Sheets: Compare (landscape), How to use.
- Inputs (yellow): quotes A2:D6 (supplier, package price, package size in oz, shipping per package); I2 amount needed 200 oz.
- Formulas:
  - E2 `=IF(OR(B2="",C2<=0),"",(B2+D2)/C2)` (E2:E6, shown to 4 decimals); F2 `=IF(E2="","",ROUND(E2*$I$2,2))`; G2 `=IF(E2="","",RANK(E2,$E$2:$E$6,1))`.
  - I4 `=IFERROR(INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)),"Enter prices")`; I5 `=IF(COUNT(E2:E6)=0,"",MIN(E2:E6))`; I6 same with MAX; I7 `=IF(COUNT(F2:F6)=0,"",MIN(F2:F6))`; I8 `=IF(COUNT(F2:F6)=0,"",MAX(F2:F6)-MIN(F2:F6))`.
- Verified values (LibreOffice Calc 24.2 only), cost per oz, cost for 200 oz, rank:
  - Pine Ridge Wax Co ($42.50, 160 oz, $12.00): $0.3406 (0.340625), $68.13, 2. Lakeside Supply ($27.99, 80 oz, $9.50): $0.4686, $93.73, 4. Harbor Craft Wholesale ($64.00, 240 oz, $0.00): $0.2667, $53.33, 1. Birchwood Candle Goods ($18.50, 48 oz, $7.25): $0.5365, $107.29, 5. Maple Lane Supply ($33.00, 100 oz, $8.00): $0.4100, $82.00, 3.
  - I4 Harbor Craft Wholesale, I5 $0.2667, I6 $0.5365, I7 $53.33, I8 $53.96. All match Python (cents rounded half up).
- Edge: package size 0 (C2) or blank price: that row's E, F, G stay blank and it drops out of MIN, MATCH and RANK. A price typed as text ("$64.00" in B4): LibreOffice converted it and still got $0.2667 (app dependent; tell readers to keep prices as numbers).
- Edge: all prices blank: I4 "Enter prices", I5 to I8 blank.
- Edge: tie for lowest (Pine Ridge and Harbor both $0.20): both rank 1 and I4 returns the first listed, Pine Ridge Wax Co.
- Limits: price and shipping only, per package; no minimum order, free shipping thresholds, quality, delivery time or multi-package shipping. Use one unit in every row. Five suppliers; extend ranges for more.

---

## 5. 2026-10-13-team-05-split-sku-code-into-parts-left-mid-find-excel-google-sheets
- Title: How to Split a Product Code Into Parts With LEFT, MID and FIND in Excel and Google Sheets
- Opening: Take everything before the first dash with `=LEFT(A2,FIND("-",A2)-1)`, so CND-LAV-8OZ gives CND. MID with a second FIND returns the middle part, and the rest of the code comes after the second dash.
- Template: `templates/tidy-tabs-split-sku-code.xlsx`. Sheets: SKUs, How to use.
- Data: SKUs A2:B11 (10 codes with units sold); formulas fill C2:E30.
- Formulas:
  - C2 `=IF(A2="","",IFERROR(LEFT(A2,FIND("-",A2)-1),"Check SKU"))`
  - D2 `=IF(A2="","",IFERROR(MID(A2,FIND("-",A2)+1,FIND("-",A2,FIND("-",A2)+1)-FIND("-",A2)-1),"Check SKU"))`
  - E2 `=IF(A2="","",IFERROR(MID(A2,FIND("-",A2,FIND("-",A2)+1)+1,LEN(A2)),"Check SKU"))`
  - Category codes G2:G5 (CND, SOP, EAR, CRD, yellow); H2 `=SUMIF($C$2:$C$30,G2,$B$2:$B$30)`; H6 `=SUM(H2:H5)`; H7 `=SUM(B2:B30)`; H8 `=IF(H6=H7,"OK","CHECK")`.
- Verified values (LibreOffice Calc 24.2 only): C, D, E match Python `str.split("-")` for all 10 codes (CND-LAV-8OZ gives CND, LAV, 8OZ; EAR-STUD-SLV gives EAR, STUD, SLV). H2 51 (CND), H3 60 (SOP), H4 27 (EAR), H5 68 (CRD), H6 206, H7 206, H8 OK.
- Edge: "CND-LAV" (one dash): D and E show Check SKU, C still CND. "NOHYPHEN": C, D, E all Check SKU. "CND-LAV-8OZ-GIFT": E returns 8OZ-GIFT (the extra part stays in the last column).
- Edge: " SOP-OAT-4OZ" with a leading space: C is " SOP" and misses the SOP total, so H8 shows CHECK (use TRIM first). Lowercase "cnd-ced-8oz": C "cnd", still counted under CND (SUMIF ignores case). With all four edits H2 63, H3 60, H6 218, H7 221, H8 CHECK.
- Limits: assumes exactly two dashes (three parts). The category list in G2:G5 is typed, so a new category needs a new row (H8 flags it). Formulas filled to row 30.
