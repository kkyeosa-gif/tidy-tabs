# Team topics - 2026-10-13 queue (5 new post topics)

Posts dated 2026-10-13 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only
(buildable and checkable by us, a download that gives a reason to visit, no tax, insurance or legal judgment). Each was checked against every title and slug in posts/ready, posts/published,
tasks/topics-seed.md and tasks/team-topics-*.md. Closest existing posts and how these differ: receivables-aging-report and highlight-overdue-invoices (how late an invoice is, this one works out the due date from the terms),
year-to-date-sales-total and count-orders-by-month (calendar windows, this one groups by weekday), count-unique-values and last-order-date-per-customer-maxifs (unique customers, latest date; this one is the share of customers who ordered twice or more),
weighted-average-cost-sumproduct and vlookup-price-list (cost of supplies bought at changing prices, a price lookup; this one compares suppliers on cost per ounce and returns the cheapest name), split-city-state-zip and split-a-full-name (addresses and names; this one splits product codes by their dashes).

Formulas are all available in LibreOffice Calc 24.2 (no FILTER, LET, LAMBDA, XLOOKUP). Build in scripts/build-templates.py and check by recalculating in LibreOffice headless. Only claim LibreOffice in `tested_in`. The search results page is assumed to be led by Microsoft and Google help pages and big sites (not checked); each topic takes a small business angle.

Ranked best first. Each entry has the topic, then the hook-writer's final title and first paragraph.

1. Invoice due date from payment terms: `invoice-due-date-payment-terms-excel-google-sheets`
- Title: How to Calculate an Invoice Due Date From Payment Terms in Excel and Google Sheets
- First paragraph: Add the days in the terms to the invoice date with `=C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE)`, where C2 is the invoice date and J2:K6 lists each terms name next to its number of days. An invoice dated 09/01/2026 on Net 30 terms is due 10/01/2026.
- Other title options: How to Add Net 30 Days to an Invoice Date in Excel and Google Sheets; How to Show Days Until an Invoice Is Due in Excel and Google Sheets
- Search wording people type: "net 30 due date formula excel", "calculate invoice due date from terms", "due date calculator google sheets"
- Template: tidy-tabs-invoice-due-dates.xlsx (Due dates, How to use). Terms dropdown, days lookup, due date, days until due and Past due / Due today / Open against a changeable as-of date.
- Verify: Python date math for 8 invoices; unknown terms; month end (01/31 plus 30 days); blank date; as-of date moved. Date arithmetic only: the terms come from the reader's own contract.

2. Sales by day of the week: `sales-by-day-of-week-sumif-excel-google-sheets`
- Title: How to Total Sales by Day of the Week in Excel and Google Sheets
- First paragraph: Name each date's weekday with `=TEXT(A2,"ddd")`, then average one day with `=AVERAGEIF($C$2:$C$29,"Sat",$B$2:$B$29)`. In the sample month, Saturdays average $159.19 and Sundays $229.25.
- Other title options: How to Find Your Best Sales Day of the Week in Excel and Google Sheets
- Search wording: "sales by day of week excel", "which day sells the most", "weekday sales summary google sheets"
- Template: tidy-tabs-sales-by-weekday.xlsx (Daily sales, By weekday, How to use). Count, total, average and share per weekday, a Best day flag, two check cells.
- Verify: Python weekday totals for 28 days; a date stored as text; an empty sheet; a $0.00 day.

3. Repeat customer rate: `repeat-customer-rate-countif-excel-google-sheets`
- Title: How to Calculate Repeat Customer Rate in Excel and Google Sheets
- First paragraph: Count orders per customer with `=COUNTIF(Orders!$B$2:$B$200,A2)`, then divide repeat customers by customers who ordered with `=COUNTIF(B2:B9,">=2")/COUNTIF(B2:B9,">0")`. In the sample, 3 of 7 customers ordered more than once, a 42.9% repeat rate.
- Other title options: How to Count Customers Who Ordered More Than Once in Excel and Google Sheets
- Search wording: "repeat customer rate formula", "count customers with more than one order excel", "returning customers spreadsheet"
- Template: tidy-tabs-repeat-customers.xlsx (Orders, Customers, How to use). Order count and type per customer, spend, repeat rate, revenue share from repeat customers, match check.
- Verify: Python counts and spend for 13 orders; misspelled name; capitals; trailing space; no orders at all; a new customer. Says what share ordered again, not why.

4. Cheapest supplier by cost per ounce: `cheapest-supplier-cost-per-ounce-index-match-excel-google-sheets`
- Title: How to Find the Cheapest Supplier by Cost Per Ounce in Excel and Google Sheets
- First paragraph: Divide price plus shipping by package size with `=(B2+D2)/C2`, then return the cheapest name with `=INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0))`. Of five fictional wax quotes, Harbor Craft Wholesale wins at $0.2667 per ounce.
- Other title options: How to Compare Supplier Prices Per Unit in Excel and Google Sheets
- Search wording: "compare supplier prices spreadsheet", "price per ounce calculator", "find lowest price index match"
- Template: tidy-tabs-supplier-price-compare.xlsx (Compare, How to use). Five quotes, cost per ounce with shipping, cost for the amount needed, rank, cheapest, spread.
- Verify: Python unit costs and rank; size 0; blank price; price typed as text; tie (first listed wins); all blank. Price only, no minimum order, quality or delivery.

5. Split a product code: `split-sku-code-into-parts-left-mid-find-excel-google-sheets`
- Title: How to Split a Product Code Into Parts With LEFT, MID and FIND in Excel and Google Sheets
- First paragraph: Take everything before the first dash with `=LEFT(A2,FIND("-",A2)-1)`, so CND-LAV-8OZ gives CND. MID with a second FIND returns the middle part, and the rest of the code comes after the second dash.
- Other title options: How to Pull the Prefix Out of a SKU in Excel and Google Sheets
- Search wording: "split sku into columns excel", "extract text before dash", "left find formula"
- Template: tidy-tabs-split-sku-code.xlsx (SKUs, How to use). Category, style and size columns, units sold by category, a CHECK cell.
- Verify: Python str.split on 10 codes; one dash only; a third dash; leading space; lowercase; no dash.
