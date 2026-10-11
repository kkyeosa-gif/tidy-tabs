# Team topics - 2026-10-12 queue (5 new post topics)

Posts dated 2026-10-12 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only
(buildable and checkable by us, a download that gives a reason to visit, no tax, insurance or legal judgment). Each was checked against every title and slug in posts/ready, posts/published,
tasks/topics-seed.md and tasks/team-topics-*.md. Closest existing posts and how these differ: highlight-duplicate-values and remove-duplicate-rows (duplicates inside one list, this one compares two lists),
quantity-price-tiers-lookup (one rate picked by a lookup, this one pays each slice at its own rate and adds them), year-to-date-sales-total and count-orders-by-month (calendar windows, this one is progress against a goal),
weighted-average-cost-sumproduct and break-even-calculator (cost of supplies bought at changing prices, and a break-even point; this one costs one batch from package prices), rank-top-customers (order by rank, this one is share of the total).

Formulas are all available in LibreOffice Calc 24.2 (no FILTER, LET, LAMBDA, XLOOKUP). Build in scripts/build-templates.py and check by recalculating in LibreOffice headless. Only claim LibreOffice in `tested_in`. The search results page is assumed to be led by Microsoft and Google help pages and big sites (not checked); each topic takes a small business angle.

Ranked best first. Each entry has the topic, then the hook-writer's final title and first paragraph.

1. Compare two lists: `compare-two-lists-find-missing-countif-excel-google-sheets`
- Title: How to Compare Two Lists and Find Missing Items in Excel and Google Sheets
- First paragraph: Flag every item in list A that is missing from list B with `=IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found")`, then fill the formula down. COUNTIF ignores capital letters, so only a different spelling or a stray space at the end of a cell counts as a mismatch.
- Other title options: How to Find Names on One List That Are Not on Another in Excel and Google Sheets; How to Match Two Lists With COUNTIF in Excel and Google Sheets
- Search wording people type: "compare two lists in excel", "find items in one list not in another", "customers not on mailing list google sheets"
- Template: tidy-tabs-compare-two-lists.xlsx (Lists, Summary, How to use). Customer emails vs newsletter subscribers, both directions, summary counts and check cells.
- Verify: hand-match 10 and 8 sample emails; capital letters differ; trailing space; a `*` typed as an item.

2. Tiered sales commission: `tiered-sales-commission-sumproduct-excel-google-sheets`
- Title: How to Calculate Tiered Sales Commission in Excel and Google Sheets
- First paragraph: Pay each tier's rate only on the sales inside that tier with `=SUMPRODUCT((B2>$G$2:$G$4)*(B2-$G$2:$G$4)*$I$2:$I$4)`, where column G holds where each tier starts and column I holds how much each tier raises the rate. With 3%, 5% and 8% tiers, $12,000.00 in sales earns $560.00.
- Other title options: How to Calculate Commission on Sales Tiers With SUMPRODUCT in Excel and Google Sheets
- Search wording: "tiered commission formula excel", "commission calculator with tiers", "sumproduct commission brackets"
- Template: tidy-tabs-tiered-commission.xlsx (Commission, How to use). Six reps, three fictional tiers in yellow cells, a second MIN/MAX method as a cross-check.
- Verify: Python per-slice math for six reps; exactly on a tier edge; $0; negative and text sales. Arithmetic only: the post tells readers to follow their written plan; no payroll or tax.

3. Progress to a monthly sales goal: `monthly-sales-goal-progress-excel-google-sheets`
- Title: How to Track Progress to a Monthly Sales Goal in Excel and Google Sheets
- First paragraph: Divide what you have sold by the goal with `=B4/B1`, then find the daily amount you still need with `=MAX(0,B1-B4)/(EOMONTH(B2,0)-B2)`, where B1 is the goal, B4 the sales so far and B2 the as-of date. With an $8,000.00 goal and $2,699.25 sold by 10/12/2026, you need about $278.99 a day for the rest of the month.
- Other title options: How to Show Percent of Sales Goal Reached in Excel and Google Sheets
- Search wording: "sales goal tracker excel", "how much do i need to sell per day", "monthly goal progress google sheets"
- Template: tidy-tabs-sales-goal-progress.xlsx (Daily sales, Goal, How to use). SUMIFS by month window, EOMONTH days left, straight-line projection.
- Verify: Python totals for October; last day of month; goal already met; a month with no sales; the projection is labeled a rough straight line.

4. Cost per unit for a handmade batch: `batch-cost-per-unit-excel-google-sheets`
- Title: How to Calculate Cost Per Unit for a Handmade Batch in Excel and Google Sheets
- First paragraph: Price each material with `=ROUND(B2/C2*D2,2)` (package price divided by units per package, times units used), add the rows, and divide the total by the items you can sell with `=ROUND(E8/H4,2)`. The sample candle batch costs $151.26 in materials, or $6.88 per sellable candle.
- Other title options: How to Work Out Cost Per Item for Candles, Soap or Jewelry in Excel and Google Sheets
- Search wording: "cost per unit calculator spreadsheet", "candle cost per candle spreadsheet", "handmade product cost breakdown template"
- Template: tidy-tabs-batch-cost-per-unit.xlsx (Batch cost, How to use). Six materials, items made and lost, cost per item made and per sellable item. Materials only: no labor, shipping, fees or overhead.
- Verify: Python per-line rounding and totals; all items lost; zero units per package; text typed in a quantity cell.

5. Percent of total by product: `percent-of-total-product-sales-excel-google-sheets`
- Title: How to Calculate Percent of Total Sales by Product in Excel and Google Sheets
- First paragraph: Divide each product's sales by the total with `=B2/SUM($B$2:$B$9)`; the dollar signs keep the total fixed when you fill down. Format the column as Percent and the shares add up to 100%.
- Other title options: How to Find Which Products Make 80 Percent of Your Sales in Excel and Google Sheets
- Search wording: "percent of total formula excel", "product share of sales", "80/20 analysis spreadsheet"
- Template: tidy-tabs-percent-of-total-sales.xlsx (Sales by product, How to use). Share, running share, Top sellers / Long tail label with a changeable 80% line.
- Verify: Python shares and running shares; the product that crosses 80% is included; unsorted list; blank sales (#DIV/0!). Says what sold, not what earned profit.
