# Handoff for posts dated 2026-10-10

Status: titles and openings (hook-writer step) done. Templates are not built yet. The formulas below come from the researcher's brief (tasks/briefs.md, "2026-10-10 신규 주제 5개") and were only pre-checked there in small LibreOffice Calc 24.2.7 (Linux) test files. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only, and only after the template recalculation step confirms the values. Cell addresses below follow the brief and may shift when the templates are built; confirm them then. Sample data is fictional (say so once per post).

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
- Verified values: (fill in at the template recalculation step, with the Etsy filter applied)
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
- Verified values: (fill in at the template recalculation step; 12 months of fictional craft or market sales)
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
- Verified values: (fill in at the template recalculation step)
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
- Verified values: (fill in at the template recalculation step; compare all 9 cells)
- Notes for next step: brief's check was Brochure + Rush = $500 and a typo such as "Spa" returning "Check spelling". Link the existing `vlookup-price-list` and `quantity-price-tiers` posts (one dimensional lookups). Dropdown clicking was not reproduced in any test; do not claim it was.

---

## 5. 2026-10-10-team-05-in-cell-bar-chart-rept-excel-google-sheets
- Title: How to Make a Bar Chart Inside Cells With REPT in Excel and Google Sheets
- Opening: Type `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))` next to each number to draw a bar made of block characters. The biggest value gets 20 blocks and every other row is scaled to match.
- Template: `templates/tidy-tabs-in-cell-bars.xlsx`. Sheets: Units sold, How to use.
- Formulas:
  - `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))`
  - `=LEN(C2)` (bar length check)
- Verified values: (fill in at the template recalculation step; 6 fictional products)
- Notes for next step: brief's check was 3200 / 1600 / 0 / 450 units giving 20 / 10 / 0 / 3 blocks. The block character looks different by font; confirm with a PDF export and offer `|` as the fallback character. Say what the check was done in (LibreOffice only).
