# Team topics — 2026-10-02 queue (5 new post topics)

Posts dated 2026-10-02 PT. No tasks/search-stats.md exists (Search Console not
connected): no search data, topics chosen on the 3-gate test only. Each was checked
against every title/slug in posts/ready and posts/published and against
topics-seed.md and team-topics-2026-09-29.md. No tax, insurance, or legal judgment in any
of them. Build each template in scripts/build-templates.py and verify by recalculating in
LibreOffice Calc (`soffice --headless --convert-to xlsx`, read back with openpyxl
data_only). Only claim LibreOffice in `tested_in`. Titles and hooks go to hook-writer.

Ranked best first.

1. Break-even calculator
- Slug: `break-even-calculator-excel-google-sheets`
- Brief: Units to cover fixed costs with `=ROUNDUP(FixedCosts/(Price-VariableCost),0)` for an Etsy or farmers market booth, plus break-even revenue.
- Target keyword: break even calculator excel template
- Template: tidy-tabs-break-even-calculator.xlsx (inputs + 4 sample products, units and revenue per product)
- Verify: hand-calc each row, compare to recalculated values.
- Competition: Investopedia/Corporate Finance Institute style pages; twist on small-booth examples with downloadable file.

2. Subscription renewal tracker
- Slug: `subscription-renewal-tracker-excel-google-sheets`
- Brief: Next renewal date with `=EDATE(Start,Months)` and days left with `=Renewal-TODAY()`, flag renewals due within 30 days.
- Target keyword: subscription tracker spreadsheet
- Template: tidy-tabs-subscription-renewal-tracker.xlsx (software/domain/Etsy fees sample, monthly cost, annual cost)
- Verify: fix the as-of date or state the run date as in the contact-list post; check EDATE on month-end dates (01/31 + 1 month).
- Competition: Microsoft EDATE help page plus template sites; twist on freelancer/small shop recurring costs.

3. Budget vs actual variance
- Slug: `budget-vs-actual-variance-excel-google-sheets`
- Brief: Variance `=C2-B2` and `=IF(B2=0,"",(C2-B2)/B2)` per category, red when over budget via conditional formatting.
- Target keyword: budget vs actual template excel
- Template: tidy-tabs-budget-vs-actual.xlsx (8 business expense categories, totals row)
- Verify: recalculate and compare to hand sums; include zero-budget row to confirm the blank case.
- Competition: Microsoft templates and big template sites; twist on a single-tab small business version with percent variance.

4. Farmers market sales log with SUMPRODUCT
- Slug: `farmers-market-sales-log-excel-google-sheets`
- Brief: Daily booth revenue from qty x price with `=SUMPRODUCT(B2:B20,C2:C20)`, and per-market-day totals with SUMIFS.
- Target keyword: farmers market sales tracker spreadsheet
- Template: tidy-tabs-farmers-market-sales-log.xlsx (Sales tab, Daily totals tab, fictional vendor data)
- Verify: recalculate; compare SUMPRODUCT to per-line sums by hand. Note overlap risk: Etsy SUMIFS post uses SUMIFS, so keep the angle on booth day totals and SUMPRODUCT.
- Competition: low; mostly generic POS and template sites, so the niche example works.

5. Clean up a customer list (TRIM, PROPER, CLEAN)
- Slug: `clean-customer-list-trim-proper-excel-google-sheets`
- Brief: Fix extra spaces, ALL CAPS and stray line breaks with `=PROPER(TRIM(CLEAN(A2)))`, and handle non-breaking spaces with SUBSTITUTE(A2,CHAR(160)," ").
- Target keyword: clean up names excel trim proper
- Template: tidy-tabs-clean-customer-list.xlsx (messy before column, cleaned after column, 10 fictional rows)
- Verify: recalculate and read back; check known edge cases ("McDonald" becomes "Mcdonald", "o'neil" becomes "O'Neil"), which become Troubleshooting. CHAR(160) behavior in LibreOffice only.
- Competition: Microsoft TRIM/PROPER help pages, ExcelJet; twist on customer-list cleanup before mailing or import. Overlaps lightly with the remove-duplicates and split-name posts, so link to them and avoid repeating their steps.
