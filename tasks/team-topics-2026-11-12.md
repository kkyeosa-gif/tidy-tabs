# Team topics - 2026-11-12 queue (5 new post topics)

Posts dated 2026-11-12. tasks/search-stats.md does not exist: 데이터 없음 (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc, a download that gives a reason to visit, no tax, wage-law or legal judgment).

Duplicate check done against posts/ready, posts/published, tasks/topics-seed.md and tasks/team-topics-*.md (grep of file names, seed titles and topic files for rank, within, category, oldest, apply payment, military, 24-hour, 1430, upc, barcode, check digit, weekend, rental, sinking, saving). Nothing matched as a topic. Closest existing posts and how they differ:
- Topic 1 vs rank-top-customers (one overall RANK) and pareto-80-20-top-products: this one ranks each product only against the products in its own category, in an unsorted list, with ties.
- Topic 2 vs receivables-aging-report, running-balance-column and track-custom-order-deposits: this one takes one lump-sum payment from a client and spreads it over that client's open invoices, oldest invoice date first, without sorting the list.
- Topic 3 vs business-days-ship-by-date (WORKDAY), count-market-saturdays-in-month and find-overlapping-bookings: this one prices a rental from a pickup and return date with one rate for Monday to Friday and another for Saturday, Sunday and listed holidays.
- Topic 4 vs time-to-decimal-hours, add-up-hours-over-24-hours and convert-yyyymmdd-number-to-date: this one turns a time exported as a plain number or text (1430, 905, 0745) into a real time shown as 2:30 PM.
- Topic 5 vs check-zip-code-format (shape of a ZIP) and pull-number-out-of-text: this one does the UPC-A check digit arithmetic (weights 3 and 1, MOD 10) and says whether the 12th digit fits.

Dropped before choosing: sinking-fund monthly savings (dropped on 11-09 as reading like financial advice; still out), hide repeated names in a grouped report (too thin for a template), split shipping cost across items by weight (too close to split-shared-booth-fee), stacked coupon discounts (close to sale-price-discount-calculator), business-hours response time between two timestamps (hard to verify every weekend and night edge in one post).

Function families: COUNTIFS/SUMPRODUCT/FIND (1), SUMIFS/MIN/MAX/ROUND (2), NETWORKDAYS/MAX/INT (3), TIME/INT/MOD/RIGHT/TEXT (4), MID/SUMPRODUCT/MOD/SUBSTITUTE (5). Themes: ranking, money allocation, date pricing, time conversion, number check. No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays, and no function that needs an _xlfn prefix (NETWORKDAYS, not NETWORKDAYS.INTL).

How to verify in LibreOffice (all five): build the .xlsx with openpyxl, recalculate headless (soffice --headless --convert-to xlsx), read the values back and compare with an independent Python model; also type the edge cases listed per topic. Excel and Google Sheets are not run by us; the posts say LibreOffice Calc only unless someone actually tests there.

Ranked best first:

1. How to Rank Products Within Each Category in Excel and Google Sheets
   - Slug: 2026-11-12-team-01-rank-products-within-each-category-countifs-excel-google-sheets
   - Search phrase: "rank within group excel", "rank by category google sheets", "countifs rank formula", "rank products in each category without sorting".
   - Template: tidy-tabs-rank-within-category.xlsx (Sales, How to use). B1 "Mark the top" (3), header row 3, rows 4 to 203 (14 sample products in 4 categories, fictional, not sorted): A Product | B Category | C Sales | D Rank in category | E Products in category | F Overall rank | G Top N mark; summary I3:J7.
     - D4 core `=COUNTIFS($B$4:$B$203,B4,$C$4:$C$203,">"&C4)+1`.
     - Expected: the best seller in each category is 1; two products tied at $720.00 are both 2 and the next is 4.
   - Test: ties, 0 and negative sales, text sales, blank category, upper and lower case categories, a trailing space in a category, * ? ~ in a category, N = 1, 5, 200, 0, -1, 2.5, text and blank, 8 random lists of 200 rows against a Python model.
   - Who ranks: Exceljet and Microsoft RANK and COUNTIFS help, forum answers using RANK with arrays. Angle: one COUNTIFS that works on an unsorted list, with the tie rule shown.

2. How to Apply a Client Payment to the Oldest Invoices First in Excel and Google Sheets
   - Slug: 2026-11-12-team-02-apply-payment-to-oldest-invoices-first-sumifs-min-excel-google-sheets
   - Search phrase: "apply payment to oldest invoice first excel", "allocate payment across invoices formula", "distribute payment to invoices google sheets", "partial payment multiple invoices excel".
   - Template: tidy-tabs-apply-payment-oldest-first.xlsx (Invoices, How to use). B1 payment received, B2 client, header row 4, rows 5 to 204 (7 sample invoices for one fictional 1099 client, one older invoice typed last): A Invoice no. | B Invoice date | C Open amount | D Applied from payment | E Balance after | F Status; summary H4:I13.
     - D5 core `=ROUND(MAX(0,MIN(C5,$B$1-(SUMIFS(older dates)+SUMIFS(same date, rows above)))),2)`.
     - Expected: $2,400.00 pays INV-1039 and INV-1041 in full and $775.00 of INV-1047.
   - Test: payment 0, exactly one invoice, one cent under and over, more than all open invoices, negative, text, blank, fractions of a cent; same-day invoices; text dates; negative and text amounts; random lists of 200 invoices against a Python model with Decimal.
   - Who ranks: accounting software help pages (QuickBooks, Xero) and a few forum threads. Angle: a plain formula version for freelancers who track invoices in a sheet. It applies a payment by date; it does not say how a payment must be applied (the remittance decides).

3. How to Price a Rental With Weekday and Weekend Day Rates in Excel and Google Sheets
   - Slug: 2026-11-12-team-03-rental-price-weekday-weekend-day-rates-networkdays-excel-google-sheets
   - Search phrase: "count weekend days between two dates excel", "weekend rate calculator excel", "rental price weekday weekend rate google sheets", "networkdays count weekends".
   - Template: tidy-tabs-rental-weekend-rate.xlsx (Rentals, Holidays, How to use). B1 weekday rate, B2 weekend or holiday rate, header row 4, rows 5 to 204 (7 fictional event-equipment rentals): A Customer | B Item | C Pickup date | D Return date | E Days billed | F Weekend and holiday days | G Weekday days | H Price | I Status; summary K4:L10; Holidays A4:A23.
     - E5 `=MAX(1,D5-C5)`; F5 `=E5-NETWORKDAYS(C5,C5+E5-1,Holidays!$A$4:$A$23)`.
     - Expected: Friday pickup, Monday return is 3 days, 2 of them weekend.
   - Test: every start weekday with lengths 0, 1, 2, 3, 6, 7, 8 and 15 days, a holiday on a Saturday, a duplicate holiday, return before pickup, text date, date with a time, rates 0, negative, text and blank, random rentals and holiday lists against a Python model.
   - Who ranks: Exceljet NETWORKDAYS and NETWORKDAYS.INTL pages, Microsoft help. Angle: a ready price sheet with the billing rule (pickup day billed, return day not) spelled out.

4. How to Convert 24-Hour Times Like 1430 to AM and PM Times in Excel and Google Sheets
   - Slug: 2026-11-12-team-04-convert-24-hour-time-numbers-like-1430-to-am-pm-time-int-mod-excel-google-sheets
   - Search phrase: "convert 1430 to time excel", "military time number to standard time google sheets", "convert 4 digit number to time excel", "hhmm number to time".
   - Template: tidy-tabs-24-hour-time-numbers.xlsx (Times, How to use). Header row 3, rows 4 to 203 (13 fictional sales tickets): A Ticket | B Time as exported | C Time | D As text (12-hour) | E Status; summary G3:H7.
     - Core `=TIME(INT(B4/100),MOD(B4,100),0)`; the file pads with RIGHT("000"&TRIM(B4),4) and checks digits and ranges.
     - Expected: 1430 gives 2:30 PM, 905 gives 9:05 AM, 0 gives 12:00 AM; 1460 and 2400 show Check.
   - Test: 0 to 2459 boundaries, text with leading zeros, spaces, colon times, decimals, negatives, letters, full width digits, 5 and 6 digit values, random values against a Python model; the core formula on the bad values.
   - Who ranks: Exceljet, Ablebits and forum answers (often TEXT and TIMEVALUE tricks). Angle: the short formula plus what it silently gets wrong (1460, 2400, 14:30).

5. How to Check a UPC Barcode Number's Check Digit in Excel and Google Sheets
   - Slug: 2026-11-12-team-05-check-upc-barcode-check-digit-mid-sumproduct-mod-excel-google-sheets
   - Search phrase: "upc check digit formula excel", "calculate upc check digit google sheets", "validate barcode number excel", "upc-a check digit calculator".
   - Template: tidy-tabs-upc-check-digit.xlsx (Barcodes, How to use). Header row 3, rows 4 to 203 (12 fictional products and codes): A Product | B UPC as typed | C Digits only | D Length | E Check digit from first 11 | F Result | G Note; summary I3:J9.
     - E4 core `=MOD(10-MOD(SUMPRODUCT(--MID(C4,{1,2,3,4,5,6,7,8,9,10,11},1),{3,1,3,1,3,1,3,1,3,1,3}),10),10)`.
     - Expected: 084512300179 OK; 084512300328 Wrong check digit (should end in 2); 84512300407 (stored as a number) 11 digits, passes with a 0 in front.
   - Test: 10, 11, 12 and 13 digits, spaces and hyphens, letter O, full width digits, a number that lost its leading 0, swaps of neighboring digits, decimals, random codes against a Python model.
   - Who ranks: GS1 US check digit calculator, barcode vendor pages, Exceljet-style forum answers. Angle: a sheet that checks a whole product list at once. It checks the arithmetic only; it does not say a code is registered (GS1 US assigns prefixes).
