# Team topics - 2026-11-08 queue (5 new post topics)

Posts dated 2026-11-08. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc, a download that gives a reason to visit, no tax, wage-law or legal judgment). Each slug and title was compared by grep with posts/ready, posts/published, tasks/topics-seed.md and every tasks/team-topics-*.md. Grepped for: slug, truncate, abbreviation, state name, items per order, one cell, comma-separated, profit by product, gross profit, last business day, last workday, milestone, crossed, hit, first date, weekend, lifetime, elapsed, turnaround.

Closest existing posts and how they differ: list-all-orders-for-one-customer and number-items-within-each-order (list rows for one customer, number lines inside an order; topic 1 joins every item of an order into ONE text cell), split-city-state-zip and orders-by-state (split an address, count by a state code you already have; topic 2 turns the full state name into the 2-letter code), weighted-average-cost, percent-of-total-product-sales and return-rate-by-product (average cost, share of sales, returns per product; topic 3 is gross profit and margin per product from a sales log with a separate price/cost list), recurring-billing-dates-twice-monthly (EOMONTH billing dates, weekend roll; topic 4 is the last business day of every month with a holiday list), monthly-sales-goal-progress and monthly-sales-column-chart-with-goal-line (progress inside one month; topic 5 finds the DATE a running total first passed a goal).

Dropped before choosing: pay-on-weekend roll (already in recurring-billing-dates), customer lifetime value (needs assumptions that read as advice), response time in hours between timestamps (too close to add-up-hours-over-24 and hours-worked-overnight), proportional stock allocation across channels (too close to split-shared-booth-fee), data validation to block weekend booking dates (no recalculated value to check).

Function families: helper-column IF/running text join (1), INDEX/MATCH exact with TRIM/clean input (2), SUMIFS/SUMPRODUCT/ROUND profit (3), EOMONTH/WORKDAY with a holiday range (4), cumulative SUM/MATCH/INDEX on dates (5). No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN-with-IF array formulas. Fictional US data only. Mix: one text joiner (1), one lookup on messy input (2), one profit summary (3), one date calculator (4), one goal tracker (5).

How to verify in LibreOffice (all five): build the .xlsx with openpyxl, recalculate headless (soffice --headless --convert-to xlsx / csv), read the values back and compare with an independent Python calculation; also type the edge cases in the Test lines.

Ranked best first:

1. How to Calculate Profit by Product From a Sales Log in Excel and Google Sheets
   - Slug: profit-by-product-from-sales-log-sumifs-excel-google-sheets
   - Search phrase: "calculate profit per product excel", "gross profit by product spreadsheet", "sales log profit margin google sheets".
   - Template: tidy-tabs-profit-by-product.xlsx (Sales, Products, How to use).
     - Products (A1:C7, headers Product | Price | Unit cost): Soy Candle 8 oz 18.00 / 6.40; Lip Balm 3.50 / 0.95; Tote Bag 24.00 / 9.10; Enamel Pin 8.00 / 2.30; Sticker Pack 4.50 / 0.70; Gift Box 32.00 / 14.25.
     - Sales (A1:E13, headers Date | Order | Product | Qty | Discount $): 12 rows 10/01/2026 to 10/28/2026, e.g. 10/01/2026, 1001, Soy Candle 8 oz, 3, 0.00; 10/02/2026, 1002, Tote Bag, 1, 2.00; ... one product sold twice, one product with no sales.
     - Product summary (Products D:H): D Units `=SUMIFS(Sales!$D$2:$D$13,Sales!$C$2:$C$13,$A2)`, E Revenue `=ROUND(D2*B2-SUMIFS(Sales!$E$2:$E$13,Sales!$C$2:$C$13,$A2),2)`, F Cost `=ROUND(D2*C2,2)`, G Profit `=E2-F2`, H Margin `=IF(E2=0,"",G2/E2)`; total row with SUM.
   - Test: product with no sales (margin blank), discount larger than line, misspelled product in Sales (add a check cell `=SUM(Sales!D2:D13)-SUM(D2:D7)` that must be 0), text qty, blank rows, 60 sales rows, compare to Python Decimal.
   - Who ranks: Exceljet SUMIFS/SUMPRODUCT, Shopify/Etsy profit calculator pages, Ablebits profit margin. Angle: two-sheet design (price list and sales log) with a reconciliation cell that catches unmatched product names; arithmetic on made-up products, not advice on what margin to aim for.

2. How to Look Up a State Abbreviation From the Full State Name in Excel and Google Sheets
   - Slug: state-name-to-abbreviation-lookup-index-match-excel-google-sheets
   - Search phrase: "convert state name to abbreviation excel", "state abbreviation lookup google sheets", "texas to tx formula".
   - Template: tidy-tabs-state-abbreviations.xlsx (Customers, States, How to use). States: 50 states plus DC (name | 2-letter code), the code list is the standard USPS list (link to USPS publication 28 as source, no judgment). Customers (A1:D13): Name | City | State as typed (mix: "Texas", " ohio ", "NEW YORK", "tx", "Calif.", "") | Code. Code `=IFERROR(INDEX(States!$B$2:$B$52,MATCH(TRIM(C2),States!$A$2:$A$52,0)),IF(AND(LEN(TRIM(C2))=2,COUNTIF(States!$B$2:$B$52,TRIM(C2))>0),UPPER(TRIM(C2)),"Check"))`; status column; count of "Check" rows.
   - Test: all 51 names, upper/lower case, spaces, nonbreaking space, already a valid code, invalid 2-letter "ZZ", "Calif.", blank, a number, "District of Columbia", Puerto Rico (not in list, shows Check), 200 rows, compare to a Python dict.
   - Who ranks: Stack Overflow, Exceljet VLOOKUP, free state list pages. Angle: messy typed input handled with TRIM and a fallback for codes already typed; the "Check" status instead of a silent blank.

3. How to Combine All Items in an Order Into One Cell in Excel and Google Sheets
   - Slug: combine-order-items-into-one-cell-running-join-excel-google-sheets
   - Search phrase: "combine multiple rows into one cell excel", "concatenate items for same order id", "join text by order number google sheets".
   - Template: tidy-tabs-items-per-order-one-cell.xlsx (Lines, Orders, How to use). Lines sorted by order (A1:E13): Order | Item | Qty | Running text | Last line? . 12 rows, orders 2001 to 2005 with 1 to 4 lines each. Running text D2 `=IF(A2=A1,D1&", "&C2&" x "&B2,C2&" x "&B2)`, Last line E2 `=IF(A2<>A3,1,0)`. Orders sheet: Order | Items; `=INDEX(Lines!$D$2:$D$13,MATCH(A2,Lines!$A$2:$A$13,1))` on sorted order numbers (last match of a sorted list returns the last row, which holds the full text).
   - Test: single-line order, 1 order only, header row guard (A2=A1 compares header text), order number as text vs number, unsorted input (shows what breaks and the sort step), cell over 255 characters, 100 lines, compare to Python groupby.
   - Who ranks: Microsoft TEXTJOIN/CONCAT help, Ablebits "combine rows", Exceljet TEXTJOIN IF. Angle: works in every version with no array formula; honest about the sort requirement.

4. How to Find the Last Business Day of Each Month in Excel and Google Sheets
   - Slug: last-business-day-of-month-eomonth-workday-excel-google-sheets
   - Search phrase: "last business day of month excel", "last working day of month formula", "month end invoice date weekday google sheets".
   - Template: tidy-tabs-last-business-day.xlsx (Months, Holidays, How to use). Start month entered once (11/01/2026); 12 rows: Month start `=EDATE(A2,1)` | Month end `=EOMONTH(A2,0)` | Last business day `=WORKDAY(B2+1,-1,Holidays!$A$2:$A$12)` | Weekday `=TEXT(C2,"dddd")` | Days earlier than month end `=B2-C2`. Holidays: 10 fictional closed dates incl. 11/26/2026 and 12/25/2026 and 12/31/2026 (user list, not a legal list).
   - Test: month ending Saturday/Sunday (Nov 30 2024 style cases), month end on a listed holiday, two holidays in a row at month end, leap Feb 2028, empty holiday list, text date in holidays, start date mid-month, 120 months against Python.
   - Who ranks: Microsoft WORKDAY/EOMONTH help, Exceljet, Ablebits. Angle: freelancer and shop month-end invoice or payout date, holiday list is yours.

5. How to Find the Date You Hit a Sales Goal in Excel and Google Sheets
   - Slug: date-you-hit-sales-goal-cumulative-total-match-excel-google-sheets
   - Search phrase: "date cumulative total reaches target excel", "when did sales pass 10000 spreadsheet", "first date running total exceeds value".
   - Template: tidy-tabs-date-hit-sales-goal.xlsx (Daily sales, How to use). Goal in B1 (10,000.00 fictional), daily sales A:B from 09/01/2026, 14 sample days with some zero days; Running total C `=SUM($B$4:B4)`; Goal reached? D `=IF(C4>=$B$1,"Yes","")`; Date hit `=IFERROR(INDEX($A$4:$A$17,MATCH(TRUE,INDEX($C$4:$C$17>=$B$1,0),0)),"Not yet")` (non-array via INDEX; alternative `=INDEX(A4:A17,COUNTIF(C4:C17,"<"&B1)+1)` valid because running total never decreases without refunds, document both); Days it took `=date-A4+1`; Amount over goal.
   - Test: goal exactly equal to a running total, goal 0, goal above total (Not yet), first day already over, refund making total fall below (why COUNTIF alternative breaks), text amounts, blank days, 365 rows, compare to Python itertools.accumulate.
   - Who ranks: Stack Overflow "first date cumulative sum exceeds", Exceljet MATCH. Angle: simple shop milestone ("when did we reach $10,000 this season"), non-array form that works in both apps.
