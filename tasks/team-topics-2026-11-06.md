# Team topics - 2026-11-06 queue (5 new post topics)

Posts dated 2026-11-06. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc 24.2, a download that gives a reason to visit, no tax, wage-law or legal judgment). Each slug and title was compared by grep with posts/ready, posts/published, tasks/topics-seed.md and every tasks/team-topics-*.md. Grepped for: parenthes, bracket, cutoff, cut-off, zone, change due, make change, bills and coins, tendered, denomination, new customer, new vs, returning, first order.

Closest existing posts and how they differ: split-sku-code-into-parts, format-phone-numbers and extract-domain-from-email (cut text at fixed characters or one known character; topic 1 cuts between two different characters and needs a guard when one is missing), business-days-ship-by-date (adds business days to a date; topic 2 first decides which day an order counts as, from the time of day, then adds the days), cash-drawer-count-over-short (counts the bills and coins already in the drawer; topic 3 works out which bills and coins to hand back for one sale), keep-leading-zeros-zip-codes and split-city-state-zip (keep or split a ZIP; topic 4 looks up a price band from the first three digits), repeat-customer-rate and flag-customers-no-order-in-90-days (customers with 2+ orders, customers gone quiet; topic 5 labels every order new or returning and counts by month).

Dropped before choosing: sales tax by ZIP (tax judgment, fails gate 3; topic 4 uses made-up zones and prices and says so), tip pooling by hours (wage-law adjacent), last 30 days of sales from an as-of date (too close to weekly-sales-totals and trailing-12-month-sales), pad SKU numbers with zeros (invoice-number-generator-text covers TEXT), real carrier zone charts (cannot be checked by us, belongs to the carrier).

Function families: MID/FIND/TRIM/IFERROR-style guards (1), WORKDAY/INT/MOD/ROUND/TEXT (2), ROUND/INT/SUMPRODUCT in whole cents (3), TEXT/LEFT/VALUE/INDEX/MATCH approximate with a text-or-number ZIP (4), MINIFS/COUNTIFS/SUMIFS/EDATE (5). No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, dynamic arrays. No tax, wage, contract or carrier-rate advice: arithmetic on fictional numbers only. Mix: one text how-to (1), one date/time calculator (2), one market cash tool (3), one lookup with messy input (4), one monthly summary table (5).

Ranked best first:

1. How to Extract Text Between Parentheses in Excel and Google Sheets
   - Slug: extract-text-between-parentheses-mid-find-excel-google-sheets
   - Search phrase: "extract text between parentheses excel", "get text inside brackets etsy variation".
   - Template: tidy-tabs-text-between-parentheses.xlsx (Listings, How to use). Listing title, text inside the first pair of parentheses, title without them, status, parts before and after the first comma.
   - Test: no parentheses, no closing ), empty (), ( ), nested, two pairs, ) before (, spaces around, comma cases, number in the cell, 300-character title, 100 rows.
   - Who ranks: Microsoft support (MID, FIND), Exceljet, Ablebits, Stack Overflow answers. Angle: Etsy variation titles, a guard for missing parentheses, and splitting the result at the comma.

2. How to Calculate a Ship Date With a Daily Cutoff Time in Excel and Google Sheets
   - Slug: ship-date-daily-cutoff-time-workday-excel-google-sheets
   - Search phrase: "ship date cutoff time excel", "order cutoff 2pm next business day formula".
   - Template: tidy-tabs-ship-date-cutoff-time.xlsx (Orders, Holidays, How to use). Cutoff time, business days to ship, order date and time, day the order counts, ship date, ship day, days after order date, closed-days list.
   - Test: orders at 13:59:59, 14:00:00 and 14:00:01, Friday after cutoff, Saturday, Sunday, holiday before and after cutoff, 0 and 20 days, midnight and 23:59:59 cutoffs, bad cutoff or days or holiday, time-only value, text value, year and leap-year boundaries, 100 rows.
   - Who ranks: Microsoft WORKDAY help, Exceljet, Ablebits. Angle: the time-of-day test before WORKDAY and why the template rounds to whole seconds.

3. How to Calculate Change Due and Count Bills and Coins in Excel and Google Sheets
   - Slug: change-due-count-bills-coins-int-sumproduct-excel-google-sheets
   - Search phrase: "change calculator bills and coins excel", "how much change to give back spreadsheet".
   - Template: tidy-tabs-change-due-bills-coins.xlsx (Change, How to use). Sale total, cash given, change due, change in cents, counts of $20, $10, $5, $1, quarters, dimes, nickels, pennies, row total to hand out.
   - Test: exact change, 1 cent, every amount 0.00 to 0.99, float traps (0.1+0.2, 9.99 from 10), short cash, negatives, text, blank, 10,000 in cash, 60 rows, an independent greedy count.
   - Who ranks: coin and change calculator sites, some Excel forum threads on "denomination breakdown". Angle: a sheet for the booth, no app, with the cents trick that keeps it exact.

4. How to Look Up a Shipping Zone From a ZIP Code in Excel and Google Sheets
   - Slug: shipping-zone-from-zip-code-index-match-excel-google-sheets
   - Search phrase: "shipping zone by zip code excel", "lookup zone from zip code spreadsheet".
   - Template: tidy-tabs-shipping-zone-from-zip.xlsx (Orders, Zones, How to use). ZIP as typed (number or text), 5-digit ZIP, first 3 digits, zone, price; a zone table with order counts and totals.
   - Test: 2108 typed as a number, text with leading zero, ZIP+4, spaces, letters, 4 and 6 digits, 1e345 and 12E45 text, zone boundaries 99/100/199/200, table with one row, 10 rows, gap, unsorted, not starting at 0, text price, duplicate zone names, empty zone table, 200 rows.
   - Who ranks: carrier zone chart pages (USPS, UPS, FedEx) and Excel lookup tutorials. Angle: not a real chart; the made-up table shows the approximate-match lookup and how to repair ZIPs that lost a zero.

5. How to Count New and Returning Customers Each Month in Excel and Google Sheets
   - Slug: new-vs-returning-customers-by-month-minifs-countifs-excel-google-sheets
   - Search phrase: "new vs returning customers excel", "count new customers per month spreadsheet".
   - Template: tidy-tabs-new-vs-returning-customers.xlsx (Orders, Months, How to use). Order list with first order date, New or Returning, a 1 on each customer's first row; a month table with orders, new customers, orders on first day, orders from returning, sales split, returning share.
   - Test: same-day repeat orders, unsorted list, capitals, trailing space, orders in other years, year 2027 and a leap year, text date, date with a time, blank customer, blank date, text amount, refund, one customer only, all new, 300 rows.
   - Who ranks: Shopify and analytics vendor blogs, Exceljet COUNTIFS. Angle: no pivot table or add-on, formulas that update when a row is added, and what "new" means when your list starts mid-history.
