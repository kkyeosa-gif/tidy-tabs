# Team topics - 2026-10-15 queue (5 new post topics)

Posts dated 2026-10-15 PT. tasks/search-stats.md does not exist: no search data (no volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us, a download that gives a reason to visit, no tax, insurance or legal judgment). Each was checked against every title and slug in posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md and the names in templates/.

Closest existing posts and how these differ: receivables-aging-report (ages open invoices today, this one averages the days it took paid invoices to get paid), invoice-due-date-payment-terms (its limits list says early pay discounts are not handled, this one is that discount), reorder-point-calculator (when to restock, this one measures how much of a batch sold), batch-cost-per-unit and shipping-cost-weight-tier-lookup (cost and rates, this one counts how many boxes a quantity fills), average-order-value-by-channel (AVERAGEIFS by channel, this one compares AVERAGE with MEDIAN on one list of orders).

Ranked best first. Each entry has the topic, then the hook-writer's title and first paragraph. The search results page is assumed to be led by Microsoft and Google help pages and big sites (not checked); each topic takes a small business angle. Only claim LibreOffice in `tested_in`.

1. Average days to get paid: `average-days-to-get-paid-excel-google-sheets`
- Title: How to Calculate Average Days to Get Paid in Excel and Google Sheets
- First paragraph: Subtract the invoice date from the paid date with `=D2-C2`, then average the results with `=AVERAGE(E2:E11)`. For the eight paid sample invoices, clients took 25.5 days on average, from 7 days at the fastest to 45 at the slowest.
- Other title options: How to Find How Long Clients Take to Pay Invoices in Excel and Google Sheets; How to Subtract Two Dates to Get Days Between in Excel and Google Sheets
- Search wording people type: "average days to pay invoice excel", "days between two dates excel", "how long do customers take to pay spreadsheet"
- Template: tidy-tabs-days-to-get-paid.xlsx (Invoices, How to use). Invoice and paid dates, days to pay, average, fastest, slowest, paid and unpaid counts, count paid later than the reader's terms, a CHECK cell.
- Verify: Python date subtraction for 8 paid invoices; unpaid rows must not count as 0; paid date typed as text; a paid date before the invoice date; no invoice paid; a paid date with no invoice date. Arithmetic only.

2. Early payment discount: `early-payment-discount-2-10-net-30-excel-google-sheets`
- Title: How to Calculate a 2/10 Net 30 Early Payment Discount in Excel and Google Sheets
- First paragraph: Find the discount with `=ROUND(D2*$M$2,2)` and the last day to take it with `=C2+$M$3`, where M2 is 2% and M3 is 10 days. On a $1,250.00 invoice dated 09/01/2026, the client saves $25.00 by paying $1,225.00 on or before 09/11/2026.
- Other title options: How to Calculate an Early Payment Discount on an Invoice in Excel and Google Sheets; How to Track Which Clients Take the 2 Percent Discount in Excel and Google Sheets
- Search wording: "2/10 net 30 formula excel", "early payment discount calculator", "invoice discount deadline google sheets"
- Template: tidy-tabs-early-payment-discount.xlsx (Discounts, How to use). Terms cells (rate, discount days, net days), discount deadline, discount amount, amount to pay by the deadline, full amount due date, Discount earned / Full amount / Open status, totals taken, missed and still available.
- Verify: Python Decimal rounding for 8 invoices; payment on the deadline day; one day late; amount typed as text; changed terms (1/15 net 45); empty sheet. Arithmetic from typed terms, no judgment about what a client owes.

3. Sell through rate: `sell-through-rate-inventory-excel-google-sheets`
- Title: How to Calculate Sell Through Rate in Excel and Google Sheets
- First paragraph: Divide units sold by units received with `=C2/B2` and format the cell as a percent. Selling 41 of 48 lavender candles is an 85.4% sell through rate.
- Other title options: How to Track What Percent of Your Inventory Sold in Excel and Google Sheets; How to Flag Slow Selling Products in Excel and Google Sheets
- Search wording: "sell through rate formula excel", "percent of inventory sold spreadsheet", "slow moving inventory google sheets"
- Template: tidy-tabs-sell-through-rate.xlsx (Products, How to use). Units received, sold, left, rate, reader-typed Strong and Slow cutoffs, status, overall rate, slow count.
- Verify: Python division for 8 products; received 0; sold more than received; sold typed as text; sold blank; empty sheet. Units only, cutoffs are the reader's own and the post says they are examples.

4. How many boxes: `how-many-boxes-needed-roundup-excel-google-sheets`
- Title: How to Calculate How Many Boxes You Need in Excel and Google Sheets
- First paragraph: Get the box count with `=ROUNDUP(B2/C2,0)`, where B2 is units to ship and C2 is units per box. Shipping 50 candles at 12 per box needs 5 boxes, not 4, because 50 divided by 12 is 4.17.
- Other title options: How to Round Up to the Next Whole Number in Excel and Google Sheets; How to Figure Out How Many Shipping Boxes an Order Fills in Excel and Google Sheets
- Search wording: "how many boxes needed formula excel", "roundup to whole number excel", "packaging calculator google sheets"
- Template: tidy-tabs-boxes-needed.xlsx (Orders, How to use). Units, units per box, boxes, units in the last box, empty slots, box cost, totals.
- Verify: Python math.ceil for 7 items; an exact multiple (48 at 12); zero units; units per box 0 or blank; units typed as text; a decimal quantity; empty sheet. One box size per row, no weights or carrier rules.

5. Median order value: `median-order-value-vs-average-excel-google-sheets`
- Title: How to Find the Median Order Value in Excel and Google Sheets
- First paragraph: Use `=MEDIAN(C2:C100)` for the middle order and `=AVERAGE(C2:C100)` for the mean. In the sample, one $640.00 wholesale order lifts the average to $83.35 while the median stays at $35.25.
- Other title options: How to Calculate Median vs Average Order Value in Excel and Google Sheets; How to Stop One Big Order From Skewing Your Average in Excel and Google Sheets
- Search wording: "median order value excel", "median vs average google sheets", "average ignoring outliers excel"
- Template: tidy-tabs-median-order-value.xlsx (Orders, How to use). Orders list, count, average, median, difference, largest order, average without orders above a typed cutoff, count above it.
- Verify: Python statistics.median and mean for 13 orders; an even count; an empty list; a total typed as text; one order; a blank row. The cutoff is the reader's choice.
