# Handoff for posts dated 2026-10-10

Status: titles and openings (hook-writer step) done. Templates built by `scripts/build-templates.py` and recalculated in LibreOffice Calc 24.2.7.2 (Linux, headless) on 10/09/2026; every formula result below was compared with hand math. The formulas below come from the researcher's brief (tasks/briefs.md, "2026-10-10 신규 주제 5개") and were only pre-checked there in small LibreOffice Calc 24.2.7 (Linux) test files. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only, and only after the template recalculation step confirms the values. Cell addresses in the Verified values sections are the real ones in the built files; where the template differs from the opening formula (for example a sheet prefix), the section says so. Sample data is fictional (say so once per post).

Titles were weighed against the style guide: each carries "Excel" and "Google Sheets", names the task and the function people search for, and has no colon, dash or hyphen. Openings give the exact formula first and promise nothing the post body cannot show.

---

## 1. 2026-10-10-team-01-sum-visible-rows-only-subtotal-excel-google-sheets
- Title: How to Sum Only the Visible Rows After Filtering in Excel and Google Sheets
- Opening: Use `=SUBTOTAL(109,C2:C16)` instead of `SUM` to total only the rows a filter leaves visible. `SUM` keeps adding filtered-out rows, so its total does not change when you filter.
- Template: `templates/tidy-tabs-subtotal-filtered-sales.xlsx`. Sheets: Sales, Totals, How to use.
- Formulas:
  - `=SUBTOTAL(109,C2:C16)` (visible total)
  - `=SUBTOTAL(103,C2:C16)` (visible count)
  - `=SUBTOTAL(101,C2:C16)` (visible average)
  - `=SUM(C2:C16)` (comparison, ignores the filter)
- Verified values (LibreOffice Calc 24.2.7.2 only):
  - Template: Sales tab A1:C16 (header row 1, data rows 2 to 16, AutoFilter on A1:C16, 15 sales, fictional). Totals tab: B column shows each formula as text, C2:C6 hold the live formulas, which read `Sales!C2:C16` (so the formula typed on the Sales sheet itself is `=SUBTOTAL(109,C2:C16)`). Totals are on their own tab so a filter never hides them.
  - Unfiltered: C2 `SUBTOTAL(109)` = $200.50, C3 `SUBTOTAL(103)` = 15, C4 `SUBTOTAL(101)` = $13.37 (13.3667), C5 `SUM` = $200.50, C6 `COUNT` = 15. Hand math: sum of the 15 amounts is $200.50.
  - AutoFilter on Channel = Etsy (criteria saved in the file with 12 rows hidden, then recalculated): C2 = $75.50, C3 = 3, C4 = $25.17 (25.1667), C5 `SUM` = $200.50 unchanged. Hand math: 34.00 + 18.50 + 23.00 = 75.50.
  - AutoFilter on Channel = Farmers market (9 rows hidden): C2 = $68.50, C3 = 6, C4 = $11.42, SUM still $200.50. Hand math: 22.00 + 8.50 + 6.00 + 14.00 + 7.50 + 10.50 = 68.50.
  - Method limit: the filter criteria and hidden rows were written into a copy of the file with openpyxl and recalculated by LibreOffice; nobody clicked the filter arrow in the LibreOffice interface, and nothing was run in Excel or Google Sheets.
- Notes for next step: the brief's test file gave SUBTOTAL(109) = $75.50, SUBTOTAL(103) = 3, SUM = $160.50 with 2 rows hidden; re-check on the real template. Do not claim Excel or Sheets handle manually hidden rows the same way (109 vs 9); link Microsoft and Google SUBTOTAL help pages for that.

---

## 2. 2026-10-10-team-02-rolling-3-month-average-sales-excel-google-sheets
- Title: How to Calculate a Rolling Average of Monthly Sales in Excel and Google Sheets
- Opening: In the row for the third month, type `=IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4))` and fill it down. Each cell then averages that month and the two before it, and the first two months stay blank.
- Template: `templates/tidy-tabs-rolling-average-sales.xlsx`. Sheets: Monthly sales, How to use.
- Formulas:
  - `=AVERAGE(B2:B4)` (fill down)
  - `=IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4))` (blank until 3 months exist)
  - `=SUM(B2:B4)/3` (check, equals the average)
- Verified values (LibreOffice Calc 24.2.7.2 only):
  - Template: Monthly sales tab. A1:E13, months 01/2026 to 12/2026 in A2:A13, sales in B2:B13. Formulas start in row 4 (C4:C13 average with the IF/COUNT test, D4:D13 `=SUM(B2:B4)/3` check, E4:E13 Yes/CHECK). C2, C3, D2, D3 are intentionally empty (no formula), so the "first two months blank" behavior is shown by having no formula there. No chart.
  - Sales: 1200, 950, 1100, 2400, 3100, 1800, 1650, 1400, 2250, 2900, 3800, 4200.
  - C4:C13 = 1083.33, 1483.33, 2200.00, 2433.33, 2183.33, 1616.67, 1766.67, 2183.33, 2983.33, 3633.33. All 10 match an independent Python calculation, and column D equals column C in every row (E shows Yes).
  - The brief's check holds in the template: 1200, 950, 1100 gives 1083.33 in C4; 950, 1100, 2400 gives 1483.33 in C5; and so on.
- Notes for next step: the brief's check was 1200 / 950 / 1100 / 2400 / 3100 / 1800 giving 1083.33, 1483.33, 2200.00, 2433.33. Optional line chart is only promised if the template really has one.

---

## 3. 2026-10-10-team-03-prorate-first-month-subscription-excel-google-sheets
- Title: How to Prorate a Partial Month of Rent or a Retainer in Excel and Google Sheets
- Opening: Count the days left in the month with `=EOMONTH(A2,0)-A2+1`, then charge `=ROUND(B2*C2/D2,2)`, where B2 is the monthly fee, C2 the days charged and D2 the days in that month. A $150.00 fee starting 10/12/2026 comes to 20 of 31 days, or $96.77.
- Template: `templates/tidy-tabs-prorate-partial-month.xlsx`. Sheets: Proration, How to use.
- Formulas:
  - Days charged: `=EOMONTH(A2,0)-A2+1`
  - Days in month: `=DAY(EOMONTH(A2,0))`
  - Prorated amount: `=ROUND(B2*C2/D2,2)`
- Verified values (LibreOffice Calc 24.2.7.2 only):
  - Template: Proration tab. Header row 1; A = start date, B = monthly fee, C = days charged, D = days in month, E = prorated amount, F = label. 7 data rows (2 to 8). C and D are formatted as Number (0).
  - Row 2: 10/12/2026, $150.00 gives 20 of 31 days = $96.77. Row 3: 02/20/2026, $150.00 gives 9 of 28 = $48.21.
  - Extra rows: 11/01/2026 $150.00 gives 30 of 30 = $150.00; 12/31/2026 $150.00 gives 1 of 31 = $4.84; 09/15/2026 $85.00 gives 16 of 30 = $45.33; 03/10/2026 $1,200.00 gives 22 of 31 = $851.61; 02/20/2028 $150.00 gives 10 of 29 = $51.72 (leap year).
  - All 7 rows match an independent calculation (calendar days, Decimal rounding half up).
- Notes for next step: brief's checks were 10/12/2026 at $150.00 giving 20 of 31 days = $96.77, and 02/20/2026 at $150.00 giving 9 of 28 days = $48.21. Day-count cells must be formatted as numbers, not dates (Troubleshooting material). Say this is arithmetic only; leases and contracts may count days differently (for example a 30 day month), so the reader follows their own agreement. No legal or landlord-tenant advice.

---

## 4. 2026-10-10-team-04-two-way-rate-card-lookup-index-match-excel-google-sheets
- Title: How to Look Up a Price by Row and Column With INDEX and MATCH in Excel and Google Sheets
- Opening: Use `=INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0))` to return the price where a service row meets a speed column. The first `MATCH` finds the row for the service in F1 and the second finds the column for the turnaround in G1.
- Template: `templates/tidy-tabs-two-way-rate-card.xlsx`. Sheets: Rate card, Quote, How to use.
- Formulas:
  - `=INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0))`
  - `=IFERROR(INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0)),"Check spelling")`
  - `=SUMPRODUCT((A2:A4=F1)*(B1:D1=G1)*B2:D4)` (comparison)
- Layout: services Logo / Brochure / Website by Standard / Rush / Same day; two dropdowns on the Quote tab.
- Verified values (LibreOffice Calc 24.2.7.2 only):
  - Template: Rate card tab A1:D4 (services Logo, Brochure, Website in A2:A4; Standard, Rush, Same day in B1:D1; prices B2:D4). The post's `F1` and `G1` are not on the Rate card tab: the two dropdowns are Quote!B2 (service) and Quote!B3 (turnaround), so the template formulas use `'Rate card'!` ranges. Quote!B5 INDEX/MATCH, B6 IFERROR version, B7 SUMPRODUCT. Quote!B10:D13 repeat the INDEX/MATCH formula for all 9 cells, and Quote!B14 `=SUMPRODUCT(--(B11:D13='Rate card'!B2:D4))` counts matches.
  - Prices: Logo $300 / $420 / $600; Brochure $350 / $500 / $700; Website $1,200 / $1,600 / $2,100.
  - All 9 combinations were set in Quote!B2:B3 (one recalculated copy each): B5, B6 and B7 all returned the rate card price for every one. Brochure + Rush = $500.00. Quote!B14 = 9.
  - Typos: "Spa" + Rush gives B5 `#N/A`, B6 "Check spelling", B7 $0. "Brochure" + "Rus" gives the same. MATCH ignores letter case, so "brochure" + "rush" still returns $500 (do not say it is case sensitive).
  - Limit: the dropdown lists were created (data validation on Quote!B2 and B3) but clicking them was not tested; values were typed into the cells.
- Notes for next step: brief's check was Brochure + Rush = $500 and a typo such as "Spa" returning "Check spelling". Link the existing `vlookup-price-list` and `quantity-price-tiers` posts (one dimensional lookups). Dropdown clicking was not reproduced in any test; do not claim it was.

---

## 5. 2026-10-10-team-05-in-cell-bar-chart-rept-excel-google-sheets
- Title: How to Make a Bar Chart Inside Cells With REPT in Excel and Google Sheets
- Opening: Type `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))` next to each number to draw a bar made of block characters. The biggest value gets 20 blocks and every other row is scaled to match.
- Template: `templates/tidy-tabs-in-cell-bars.xlsx`. Sheets: Units sold, How to use.
- Formulas:
  - `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))`
  - `=LEN(C2)` (bar length check)
- Verified values (LibreOffice Calc 24.2.7.2 only):
  - Template: Units sold tab. A = product, B = units (rows 2 to 7), C = block bar, D = `=LEN(C2)`, E = the same bar built with `|`.
  - Units 3200 / 1600 / 0 / 450 / 2400 / 800 give 20 / 10 / 0 / 3 / 15 / 5 blocks in D2:D7. All match hand math (units / 3200 * 20 rounded). The `|` bars in column E have the same lengths. The sample avoids values that land on x.5 blocks, where rounding could differ.
  - PDF check: exported to PDF from LibreOffice Calc 24.2.7.2 (US Letter, landscape). The block character renders as solid bars with faint vertical seams between characters; the default Calibri substitute (Carlito) has no block glyph, so LibreOffice drew it from DejaVu Sans. So the glyph depends on font fallback: keep the `|` column as the documented fallback. Not checked in Excel or Google Sheets, where the look may differ.
- Notes for next step: brief's check was 3200 / 1600 / 0 / 450 units giving 20 / 10 / 0 / 3 blocks. The block character looks different by font; confirm with a PDF export and offer `|` as the fallback character. Say what the check was done in (LibreOffice only).
