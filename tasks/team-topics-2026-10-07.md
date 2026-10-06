# Team topics - 2026-10-07 queue (5 new post topics)

Posts dated 2026-10-07 PT. No tasks/search-stats.md exists (Search Console not
connected): no search data, topics chosen on the 3-gate test only. Each was checked
against every title/slug in posts/ready and posts/published and against
topics-seed.md and team-topics-2026-09-29.md, 2026-10-02.md and 2026-10-06.md. No tax,
insurance, or legal judgment in any of them. Build each template in
scripts/build-templates.py and verify by recalculating in LibreOffice Calc
(`soffice --headless --convert-to xlsx`, read back with openpyxl data_only). Only claim
LibreOffice in `tested_in`. Titles and hooks go to hook-writer.

Formula mix is new to the blog: approximate-match lookup, VALUE/SUBSTITUTE, reorder
point math, discount math, hourly-rate math. None needs functions missing from
LibreOffice 24.2 (no FILTER, no LET).

Ranked best first.

1. Sale price and discount calculator
- Slug: `sale-price-discount-calculator-excel-google-sheets`
- Brief: Sale price from a percent off with `=B2*(1-C2)`, discount in dollars, and the reverse (original price from a sale price with `=D2/(1-C2)`), plus the "two discounts do not add up" case.
- Target keyword: discount calculator excel
- Template: tidy-tabs-sale-price-discount-calculator.xlsx (8 fictional Etsy shop items, 10%/15%/25% off, sale price, dollars saved, reverse-lookup block, stacked 20% + 10% example)
- Verify: recalculate; hand-check $48.00 at 25% off equals $36.00 and 20% then 10% off $100.00 equals $72.00, not $70.00.
- Competition: Microsoft percentage help page, ExcelJet; twist on shop sale pricing with a stacked-coupon check.

2. Shipping cost by weight tier (approximate match lookup)
- Slug: `shipping-cost-weight-tier-lookup-excel-google-sheets`
- Brief: Look up a shipping charge from a weight band table with `=VLOOKUP(B2,$F$2:$G$7,2,TRUE)` and `=INDEX/MATCH(...,1)`, and explain why the table must be sorted ascending.
- Target keyword: vlookup approximate match shipping rates
- Template: tidy-tabs-shipping-weight-tier-lookup.xlsx (fictional rate table in ounces/pounds bands, 12 sample orders, one unsorted-table demo, one below-minimum weight showing #N/A with IFERROR fix)
- Verify: recalculate; check boundary weights (exactly 4 oz, 4.01 oz) land in the expected band; confirm the unsorted table gives a wrong answer. Rates are made up, say so once. Links to the exact-match VLOOKUP post without repeating it.
- Competition: Microsoft VLOOKUP help, ExcelJet "approximate match" pages; twist on shipping bands for Etsy sellers.

3. Reorder point calculator
- Slug: `reorder-point-calculator-excel-google-sheets`
- Brief: Reorder point `=ROUNDUP(DailySales*LeadDays+SafetyStock,0)` with a flag `=IF(OnHand<=ReorderPoint,"REORDER","OK")`, and daily sales from 30 days of units sold.
- Target keyword: reorder point formula excel
- Template: tidy-tabs-reorder-point-calculator.xlsx (6 fictional craft supplies, units sold last 30 days, lead time days, safety stock, on hand, reorder flag)
- Verify: recalculate; hand-check 90 units in 30 days with 7 lead days and 6 safety stock gives 27; confirm flags at the boundary (on hand equals reorder point). Links to the craft inventory tracker post.
- Competition: inventory template sites, Investopedia style definitions; twist on a formula-only reorder flag for a handmade shop.

4. Freelance hourly rate calculator
- Slug: `freelance-hourly-rate-calculator-excel-google-sheets`
- Brief: Turn a yearly income goal, overhead, weeks off and billable-hour share into an hourly rate with `=(Goal+Overhead)/(BillableHours)`. Billable hours = `=(52-WeeksOff)*HoursPerWeek*BillablePct`.
- Target keyword: freelance hourly rate calculator spreadsheet
- Template: tidy-tabs-freelance-hourly-rate-calculator.xlsx (inputs: $60,000.00 goal, $6,000.00 overhead, 4 weeks off, 30 hours per week, 65% billable; resulting rate; 3 what-if columns)
- Verify: recalculate; hand-check (48 x 30 x 0.65 = 936 hours; $66,000.00 / 936 = $70.51). Illustration only: the post does not cover taxes or self-employment costs and says to see an accountant for those.
- Competition: freelancer blogs and online calculators; twist on a downloadable what-if sheet with the math shown cell by cell.

5. Convert text to numbers (pasted $ signs and commas)
- Slug: `convert-text-to-numbers-dollar-signs-excel-google-sheets`
- Brief: Fix numbers that will not add up because they are text, using `=VALUE(SUBSTITUTE(SUBSTITUTE(A2,"$",""),",",""))`, plus TRIM for stray spaces and a quick check with `=ISNUMBER(A2)`.
- Target keyword: excel numbers stored as text sum zero
- Template: tidy-tabs-convert-text-to-numbers.xlsx (10 fictional pasted bank/Etsy amounts such as "$1,250.00", " 18.00", "(45.50)", text column, ISNUMBER check, cleaned number column, SUM before and after)
- Verify: recalculate; confirm SUM of the text column is 0 and the cleaned column matches hand total; document what happens to the parenthesis negative (handle with a second SUBSTITUTE or leave and say so). Locale note: tested in LibreOffice with US locale only.
- Competition: Microsoft "numbers stored as text" help page, ExcelJet; twist on pasted dollar amounts from a bank or Etsy export. Links to the leading zeros post for the opposite problem.
