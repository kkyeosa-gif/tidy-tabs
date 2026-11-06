# Team topics - 2026-11-07 queue (5 new post topics)

Posts dated 2026-11-07. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc 24.2, a download that gives a reason to visit, no tax, wage-law or legal judgment). Each slug and title was compared by grep with posts/ready, posts/published, tasks/topics-seed.md and every tasks/team-topics-*.md. Grepped for: last word, file name, extension, return window, return by, delivery date, days left, wildcard, partial, ordered once, orders per customer, how many customers, frequency, keystone, wholesale, landed, per use.

Closest existing posts and how they differ: split-sku-code-into-parts, pull-number-out-of-text and initials-from-full-name (cut at fixed positions or the first marker; topic 1 cuts at the LAST dot, which needs a count-the-dots trick), prorate-first-month and invoice-due-date-payment-terms (add days to an invoice date; topic 2 adds a return window to a delivery date, allows a per-order window and reports days left as of a typed date), sum-if-cell-contains-text-wildcard (adds amounts for names that contain a word; topic 3 finds one product row and its price from part of a name, and counts how many rows match), repeat-customer-rate and flag-customers-no-order-in-90-days (one rate, one flag per customer; topic 4 builds the table "how many customers ordered once, twice, 3 times" with the sales in each group), quantity-price-tiers-lookup, markup-vs-margin-calculator and price-increase-planner (wholesale tiers by quantity, markup vs margin, new prices; topic 5 derives a wholesale price from the retail price and a wholesale percent, and tests your margin at that price).

Dropped before choosing: warranty or refund deadlines as legal rights (legal judgment, fails gate 3; topic 2 is the shop's own policy and says so), gift card balances (same call as 10-31), estimated delivery windows by carrier (cannot be checked by us), sales tax and wage topics (gate 3), cost per use of equipment (depreciation judgment), a starting cash float list (too close to change-due-count-bills-coins), case-sensitive lookups (EXACT array tricks not reliable across apps).

Function families: SUBSTITUTE/LEN/FIND/MID/LEFT/LOWER with a CHAR(1) marker (1), INT/TEXT/date subtraction with status labels (2), COUNTIF/MATCH with * wildcards and ~ escapes plus INDEX (3), COUNTIF/SUMPRODUCT/COUNTIFS/SUMIFS in a frequency table (4), ROUND/ROUNDUP margin arithmetic in cents (5). No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, dynamic arrays. No tax, wage, contract or legal advice: arithmetic on fictional numbers only. Mix: one text how-to (1), one date calculator (2), one lookup (3), one summary table (4), one pricing tool (5).

Ranked best first:

1. How to Split a File Name Into Name and Extension in Excel and Google Sheets
   - Slug: split-file-name-and-extension-last-dot-substitute-excel-google-sheets
   - Search phrase: "get file extension excel", "separate file name and extension google sheets", "text after last period excel".
   - Template: tidy-tabs-split-file-name-extension.xlsx (Files, How to use). File name, cleaned name, position of the last dot, extension, name without extension, status, lowercase extension; a count of files per extension you type.
   - Test: no dot, one dot, several dots, leading dot (.gitignore), trailing dot, only a dot, spaces around, a space inside the extension, upper and lower case, numbers in the cell, 300-character name, 100 rows.
   - Who ranks: Exceljet, Ablebits, Stack Overflow answers for "text after last delimiter". Angle: Etsy digital download and design files, a status column for names that have no real extension.

2. How to Calculate a Return-By Date From the Delivery Date in Excel and Google Sheets
   - Slug: return-by-date-from-delivery-date-days-left-excel-google-sheets
   - Search phrase: "return deadline formula excel", "date plus 30 days excel", "days left until return window closes".
   - Template: tidy-tabs-return-deadline.xlsx (Orders, How to use). Default window, as-of date you type, closing-soon days; per order delivery date, optional window override, return-by date, weekday, days left, status (Open, Closing soon, Last day, Closed).
   - Test: window 0 and 30 and 365, override 0 and 14 and text and 1.5, delivery on month and year ends, leap day, date with a time, text date, blank, time-only value, as-of before and after, last day equals as-of, bad settings, 100 rows.
   - Who ranks: Microsoft date-add help, Exceljet, generic "date plus days" pages. Angle: the shop's own return policy (not a legal rule), a typed as-of date instead of TODAY so a printout stays the same.

3. How to Look Up a Price From Part of a Product Name in Excel and Google Sheets
   - Slug: look-up-price-by-part-of-product-name-wildcard-match-excel-google-sheets
   - Search phrase: "vlookup partial match excel", "lookup part of text in cell", "index match wildcard google sheets".
   - Template: tidy-tabs-lookup-by-part-of-name.xlsx (Search, Products, How to use). Products list (SKU, name, price); up to 20 search words with number of matches, first match, SKU, price and a status.
   - Test: capitals, spaces around, a term that matches none, one, many, the first and last rows, the whole name, * ? ~ typed in the term, a 250-character term, numbers as terms, blank product cells, duplicate names, text price, 100 products.
   - Who ranks: Exceljet, Ablebits, Stack Overflow for "vlookup partial match". Angle: a "Several matches" warning so a short word does not silently return the wrong item; wildcard characters typed in the term are escaped.

4. How to Count Customers by Number of Orders in Excel and Google Sheets
   - Slug: customers-by-number-of-orders-frequency-table-countif-excel-google-sheets
   - Search phrase: "count customers by number of orders excel", "how many customers ordered more than once", "order frequency table".
   - Template: tidy-tabs-customers-by-number-of-orders.xlsx (Orders, How to use). Order list (customer, amount) with each customer's order count and a first-row flag; a table for 1, 2, 3, 4 and 5 or more orders with customers, share of customers, orders, sales and share of sales.
   - Test: one customer only, all customers once, capitals, trailing space (matched after TRIM), blank customer, text amount, 5+ group, wildcard characters in names, 300 rows, random sets against a Python Counter.
   - Who ranks: Shopify and analytics vendor blogs, pivot table tutorials. Angle: no pivot table, no UNIQUE; two helper columns and a five-row table that updates when a row is added.

5. How to Set a Wholesale Price From the Retail Price and Check Your Margin in Excel and Google Sheets
   - Slug: wholesale-price-from-retail-price-margin-check-excel-google-sheets
   - Search phrase: "wholesale price from retail price formula", "50 percent wholesale pricing spreadsheet", "wholesale margin calculator excel".
   - Template: tidy-tabs-wholesale-from-retail-price.xlsx (Prices, How to use). Wholesale percent of retail, your target margin; per product unit cost, retail price, wholesale price, profit and margin at wholesale, margin at retail, status, wholesale and retail price that would reach the target.
   - Test: 50% and 40% and 100% and 1%, target 0% and 99%, margin exactly on target, cost above wholesale (Loss), cost zero, retail zero, text, half-cent rounding, 60 rows, random sets against an exact-decimal Python model.
   - Who ranks: pricing blogs for makers, Faire and Etsy seller guides, margin calculator sites. Angle: arithmetic only, on made-up products; the percent and target are yours to change, nothing says what they should be.
