# Team topics - 2026-10-11 queue (5 new post topics)

Posts dated 2026-10-11 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only
(buildable and checkable by us, a download that gives a reason to visit, no tax, insurance or legal judgment). Each was checked against every title and slug in posts/ready, posts/published,
tasks/topics-seed.md and tasks/team-topics-*.md. Closest existing posts and how these differ: sum-expenses-by-category (exact-match SUMIF, this one is wildcard text matching),
year-to-date-sales-total and count-orders-by-month (month and YTD windows, this one is quarters), highlight-duplicate-values (duplicates, this one is gaps in a sequence),
invoice-number-generator (makes numbers, this one audits them), sale-price-discount and markup-vs-margin (price math, this one is marketplace fees), loan-payment-calculator (interest, this one splits an invoice with no interest).

Formulas are all available in LibreOffice Calc 24.2 (no FILTER, LET, LAMBDA, XLOOKUP). Build in scripts/build-templates.py and check by recalculating in LibreOffice headless. Only claim LibreOffice in `tested_in`. The search results page is assumed to be led by Microsoft and Google help pages and big sites (not checked); each topic takes a small business angle.

Ranked best first.

1. Total sales by quarter: `quarterly-sales-totals-sumifs-excel-google-sheets`
- Title: How to Total Sales by Quarter in Excel and Google Sheets
- Search wording people type: "sum by quarter excel", "group dates by quarter", "quarterly sales total google sheets"
- Template: tidy-tabs-quarterly-sales-totals.xlsx (Sales, By quarter, How to use). Helper column `=YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0)`, then SUMIFS; second method by date range with EDATE.
- Verify: hand totals per quarter; boundary dates 03/31 and 04/01; dates stored as text.

2. Find missing invoice numbers: `find-missing-invoice-numbers-countif-excel-google-sheets`
- Title: How to Find Missing Invoice Numbers in a Sequence in Excel and Google Sheets
- Search wording: "find missing numbers in a sequence excel", "missing invoice number check", "gaps in invoice numbers"
- Template: tidy-tabs-missing-invoice-numbers.xlsx (Invoices, Check, How to use). Expected list `=MIN(...)+ROW()-2`, `=COUNTIF(...)`, Missing / Duplicate / OK. A missing number is not judged: it may be a voided invoice.
- Verify: sample with 3 gaps and 1 duplicate; text numbers, prefixes, numbers outside the list.

3. Sum cells that contain certain text: `sum-if-cell-contains-text-wildcard-sumif-excel-google-sheets`
- Title: How to Sum Cells That Contain Certain Text With SUMIF Wildcards in Excel and Google Sheets
- Search wording: "sumif contains text", "sum if cell contains word excel", "sumif wildcard"
- Template: tidy-tabs-sumif-contains-text.xlsx (Expenses, Totals, How to use). `=SUMIF(range,"*"&A2&"*",sum_range)`, starts with, ends with, `~*` escape, COUNTIF.
- Verify: hand totals; substring trap ("ink" inside "Pink"); blank search cell; number in text column.

4. Split an invoice into equal monthly payments: `split-invoice-into-installments-excel-google-sheets`
- Title: How to Split an Invoice Into Equal Monthly Payments in Excel and Google Sheets
- Search wording: "payment plan schedule excel", "split payment into equal installments", "installment schedule template"
- Template: tidy-tabs-invoice-installments.xlsx (Payment plan, How to use). `=ROUND(B1/B2,2)`, last payment takes the leftover cents, EDATE due dates, running balance. No interest, no fees.
- Verify: totals add to the invoice exactly; $1,000.00 in 3 payments; month-end start date; 1 and 12 payments.

5. Net payout after marketplace fees: `marketplace-fee-net-payout-calculator-excel-google-sheets`
- Title: How to Calculate Net Payout After Marketplace Fees in Excel and Google Sheets
- Search wording: "etsy fee calculator spreadsheet", "net payout after fees formula", "calculate price to cover fees"
- Template: tidy-tabs-marketplace-net-payout.xlsx (Net payout, How to use). Fee rates are fictional inputs; the reader types their own. Forward net and reverse price for a target payout. No shipping, sales tax or income tax.
- Verify: hand math per order; tiny orders where the fixed fee exceeds the price; target price round trip.
