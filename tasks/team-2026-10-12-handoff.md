# Handoff for posts dated 2026-10-12

Status: topics (researcher), titles and openings (hook-writer) and templates (template-builder) done. Templates are built by `scripts/build-templates.py` and were recalculated in LibreOffice Calc 24.2 (Linux, headless: `soffice --headless --convert-to xlsx` into a temp dir, read back with openpyxl `data_only`) on 10/11/2026. Every result below was compared with hand or Python math. Edge cases listed were run by editing a copy of the template with openpyxl, recalculating it in LibreOffice and reading the values. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. No tasks/search-stats.md exists: no search data. Sample data is fictional (say so once per post). Each template tab prints on one US Letter page wide; the PDF export from LibreOffice gave one page per sheet (landscape for tiered commission, batch cost and percent of total).

Titles carry "Excel and Google Sheets", name the task, and have no colon, dash or hyphen. Openings give the exact formula first. Cell addresses below are the real ones in the built files. Note that running scripts/build-templates.py re-saves every older template too; only the five new files were kept.

---

## 1. 2026-10-12-team-01-compare-two-lists-find-missing-countif-excel-google-sheets
- Title: How to Compare Two Lists and Find Missing Items in Excel and Google Sheets
- Opening: Flag every item in list A that is missing from list B with `=IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found")`, then fill the formula down. COUNTIF ignores capital letters, so only a different spelling or a stray space at the end of a cell counts as a mismatch.
- Template: `templates/tidy-tabs-compare-two-lists.xlsx`. Sheets: Lists, Summary, How to use.
- Formulas:
  - Lists!B2: `=IF(A2="","",IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found"))` (filled to B15); Lists!E2: `=IF(D2="","",IF(COUNTIF($A$2:$A$200,D2)=0,"Not found","Found"))` (filled to E15). The template adds the blank guard; the opening leaves it out. Not found cells are red.
  - Summary!B2 `=COUNTA(Lists!A2:A200)`, B3 `=COUNTIF(Lists!B2:B200,"Found")`, B4 Not found count, B5 to B7 the same for list B (column D and E), B8 and B9 `=IF(B3+B4=B2,"OK","CHECK")` and `=IF(B6+B7=B5,"OK","CHECK")`.
- Verified values (LibreOffice Calc 24.2 only):
  - Lists A2:A11 are 10 customer emails, D2:D9 are 8 subscriber emails. Not in list B: ana.gomez (B2), maya.cole (B7), omar.said (B8), jo.hart (B9). Not in list A: chris.doe (E8), pat.lane (E9). Tina.Brooks (A10) matches TINA.BROOKS (D4): Found, so case is ignored.
  - Summary B2 10, B3 6, B4 4, B5 8, B6 6, B7 2, B8 OK, B9 OK. Matches hand check.
  - Edge: trailing space added to D5 ("sam.patel@example.com "): B4 Not found and E5 Not found, so a stray space makes a false mismatch. Fix with TRIM.
  - Edge: a `*` typed as an item in A9 returns Found (it matches everything in list B). With both edits together Summary shows B3 6, B4 4, B6 5, B7 3.
- Limits: exact match of whole cells (apart from capitals); no fuzzy matching of near spellings. Wildcard characters `* ? ~` in an item behave as wildcards. Ranges stop at row 200; formulas are filled to row 15.

---

## 2. 2026-10-12-team-02-tiered-sales-commission-sumproduct-excel-google-sheets
- Title: How to Calculate Tiered Sales Commission in Excel and Google Sheets
- Opening: Pay each tier's rate only on the sales inside that tier with `=SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4)`, where column G holds where each tier starts and column I holds how much each tier raises the rate. With 3%, 5% and 8% tiers, $12,000.00 in sales earns $560.00.
- Template: `templates/tidy-tabs-tiered-commission.xlsx`. Sheets: Commission (landscape), How to use.
- Inputs (yellow, fictional): G2:G4 tier starts $0.00, $5,000.00, $10,000.00; H2:H4 rates 3.0%, 5.0%, 8.0%.
- Formulas:
  - I2 `=H2`, I3 `=H3-H2`, I4 `=H4-H3` (2.0% extra at I3, 3.0% at I4 and I2).
  - C2: `=ROUND(SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4),2)` (C2:C7). The template adds ROUND; the opening leaves it out.
  - D2 second method: `=ROUND(MIN(B2,$G$3)*$H$2+MAX(0,MIN(B2,$G$4)-$G$3)*$H$3+MAX(0,B2-$G$4)*$H$4,2)`; E2 `=IF(ROUND(C2-D2,2)=0,"Yes","CHECK")`. Totals row 8: B8, C8.
- Verified values (LibreOffice Calc 24.2 only):
  - Ana $4,200.00 gives $126.00; Ben $5,000.00 gives $150.00 (exactly on the edge); Cy $12,000.00 gives $560.00; Dee $0.00 gives $0.00; Eli $25,500.50 gives $1,640.04; Fay $10,000.00 gives $400.00. C2:C7 equals D2:D7 and E2:E7 all Yes. B8 $56,700.50, C8 $2,876.04. All match Python (per-slice math, round half up).
  - Edge: $10,000.01 gives $400.00 in both methods (the extra cent earns $0.0008).
  - Edge: negative sales (-100 in B5): C5 $0.00 but D5 -$3.00, so the two methods disagree and E5 would flag CHECK. Refunds should be handled before using this sheet.
  - Edge: sales typed as text ("5000" in B3): C3 $0.00 while D3 $400.00, so the methods disagree. Keep sales as real numbers.
- Limits: three tiers; adding a tier means extending the ranges. Tiers must be in ascending order. Arithmetic for a written plan only: no payroll, withholding or tax, and it does not decide what a rep is owed.

---

## 3. 2026-10-12-team-03-monthly-sales-goal-progress-excel-google-sheets
- Title: How to Track Progress to a Monthly Sales Goal in Excel and Google Sheets
- Opening: Divide what you have sold by the goal with `=B4/B1`, then find the daily amount you still need with `=MAX(0,B1-B4)/(EOMONTH(B2,0)-B2)`, where B1 is the goal, B4 the sales so far and B2 the as-of date. With an $8,000.00 goal and $2,699.25 sold by 10/12/2026, you need about $278.99 a day for the rest of the month.
- Template: `templates/tidy-tabs-sales-goal-progress.xlsx`. Sheets: Daily sales, Goal, How to use.
- Inputs (yellow, Goal tab): B1 goal $8,000.00, B2 as-of date 10/12/2026 (a typed date so the sample does not change; the notes say =TODAY() works).
- Formulas (Goal tab):
  - B3 `=EOMONTH(B2,-1)+1`; B4 `=SUMIFS('Daily sales'!$B$2:$B$500,'Daily sales'!$A$2:$A$500,">="&B3,'Daily sales'!$A$2:$A$500,"<="&B2)`; B5 `=B4/B1`; B6 `=MAX(0,B1-B4)`; B7 `=EOMONTH(B2,0)-B2`; B8 `=IF(B6=0,"",IF(B7=0,"Last day",B6/B7))`; B9 `=B4/DAY(B2)`; B10 `=B9*DAY(EOMONTH(B2,0))`; B11 `=IF(B10>=B1,"On pace","Behind")`.
- Verified values (LibreOffice Calc 24.2 only):
  - Daily sales A2:B12: 11 rows, 09/30/2026 $250.00 (prior month, excluded) plus 10 October days. B3 10/01/2026, B4 $2,699.25, B5 33.7%, B6 $5,300.75, B7 19, B8 $278.99, B9 $224.94, B10 $6,973.06, B11 Behind. All match Python.
  - Edge: as-of 10/31/2026: B7 0, B8 "Last day", B9 $87.07, B10 $2,699.25 (the average per day drops because all 31 days count).
  - Edge: goal $2,000.00: B5 134.96%, B6 $0.00, B8 blank, B11 On pace.
  - Edge: as-of 11/05/2026 with no November sales: B3 11/01/2026, B4 $0.00, B7 25, B8 $320.00, B10 $0.00, B11 Behind.
- Limits: the projection is a straight-line average, not a forecast. Sales are matched by date, so dates must be real dates (not text). Data range stops at row 500. One calendar month at a time.

---

## 4. 2026-10-12-team-04-batch-cost-per-unit-excel-google-sheets
- Title: How to Calculate Cost Per Unit for a Handmade Batch in Excel and Google Sheets
- Opening: Price each material with `=ROUND(B2/C2*D2,2)` (package price divided by units per package, times units used), add the rows, and divide the total by the items you can sell with `=ROUND(E8/H4,2)`. The sample candle batch costs $151.26 in materials, or $6.88 per sellable candle.
- Template: `templates/tidy-tabs-batch-cost-per-unit.xlsx`. Sheets: Batch cost (landscape), How to use.
- Inputs (yellow): H2 items made 24, H3 items lost or unsellable 2; materials in A2:D7 (package price, units per package, units used).
- Formulas:
  - E2 `=ROUND(B2/C2*D2,2)` (E2:E7); E8 `=SUM(E2:E7)`; H4 `=H2-H3`; H5 `=IF(H2<=0,"",ROUND(E8/H2,2))`; H6 `=IF(H4<=0,"",ROUND(E8/H4,2))`.
- Verified values (LibreOffice Calc 24.2 only):
  - E2:E7: Soy wax $18.90, Fragrance oil $4.20, Cotton wicks $2.88, Tin jars $120.00 ($60.00 per case of 12, 24 used), Labels $4.56, Dye chips $0.72. E8 $151.26, H4 22, H5 $6.30, H6 $6.88. Python (per-line rounding, half up) gives 151.26, 6.30, 6.88.
  - Edge: items lost equal items made (H3 24): H4 0, H6 blank (no divide by zero). Units per package 0 (C2): E2 and E8 and H5 show #DIV/0!.
  - Edge: text in a quantity cell ("2.4 oz" in D3): E3 and E8 show #VALUE!.
- Limits: materials only. Labor, packaging bought separately, fees, shipping, overhead and taxes are not included. Units must match inside a row (pounds with pounds). Per-line rounding to cents can differ by a cent or two from an unrounded total.

---

## 5. 2026-10-12-team-05-percent-of-total-product-sales-excel-google-sheets
- Title: How to Calculate Percent of Total Sales by Product in Excel and Google Sheets
- Opening: Divide each product's sales by the total with `=B2/SUM($B$2:$B$9)`; the dollar signs keep the total fixed when you fill down. Format the column as Percent and the shares add up to 100%.
- Template: `templates/tidy-tabs-percent-of-total-sales.xlsx`. Sheets: Sales by product (landscape), How to use.
- Inputs (yellow): H2 the share the top group covers, 80%. Products in A2:B9, sorted largest first.
- Formulas:
  - C2 `=B2/SUM($B$2:$B$9)`; D2 `=SUM($B$2:B2)/SUM($B$2:$B$9)`; E2 `=IF(ROUND(D2-C2,6)<$H$2,"Top sellers","Long tail")`; B10 `=SUM(B2:B9)`; C10 `=SUM(C2:C9)`; H4 `=COUNTIF(E2:E9,"Top sellers")`; H5 `=SUMIF(E2:E9,"Top sellers",C2:C9)`.
- Verified values (LibreOffice Calc 24.2 only):
  - Sales: Soy candles $4,200.00 (42.0%), Soap bars $2,600.00 (26.0%), Brass earrings $1,500.00 (15.0%), Gift boxes $900.00 (9.0%), Stickers $450.75 (4.5075%), Gift tags $199.50 (1.995%), Greeting cards $101.25 (1.0125%), Tote bags $48.50 (0.485%). Running %: 42.0, 68.0, 83.0, 92.0, 96.5075, 98.5025, 99.515, 100.0. Top sellers: first 3 (the third crosses 80% and is included), rest Long tail. B10 $10,000.00, C10 100.0%, H4 3, H5 83.0%. All match Python.
  - Edge: list not sorted (Tote bags $4,200.00 moved to row 9, Soy candles $48.50 to row 2): every row comes out Top sellers (H4 8, H5 100%), because running % is meaningless when unsorted.
  - Edge: all sales blank: C2 and C10 show #DIV/0!, B10 $0.00, H4 0.
- Limits: 8 products; extend ranges for more. The 80% line is a rule of thumb. Percent of sales says nothing about profit.
