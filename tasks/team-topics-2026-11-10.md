# Team topics - 2026-11-10 queue (5 new post topics)

Posts dated 2026-11-10. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc, a download that gives a reason to visit, no tax, wage-law or legal judgment).

Duplicate check done against posts/ready, posts/published and tasks/topics-seed.md (grep of file names and seed titles for rank, unique, zip, time zone, break even, round to, 99, proper, networkdays, dropdown, invoice number, percent change, tip, largest, bundle, countdown, split, etsy, character, tag, calendar, budget, variance, cents, case, minimum order, latest, last price, between two dates, project, week). Closest existing posts and how they differ:
- Topic 1 vs lookup-to-the-left, price-in-effect-on-order-date (approximate match on a date table) and look-up-price-by-part-of-product-name (first match, wildcard): this one returns the LAST matching row of an unsorted purchase log with LOOKUP(2,1/...) and compares it with the newest date.
- Topic 2 vs how-many-boxes-you-need (packaging for an order), reorder-point and days-of-stock-left (when to order): this one is how many to order once the supplier's case size and minimum order are known.
- Topic 3 vs sum-cells-containing-text-wildcards (SUMIF with *), count-listing-characters-tags (length of tags) and combine-order-items-into-one-cell: this one counts orders that carry an exact tag inside a comma separated cell, where a wildcard over-counts gift as gifted and gift wrap.
- Topic 4 vs convert-text-to-numbers (dollar signs and commas) and convert-currency (exchange rates): this one is whole cents in a payments export, with fee, net and a payout check.
- Topic 5 vs weekly-sales-totals, week-ending-date and total-sales-by-quarter (fixed calendar periods): this one has a start date and a period length setting, so any run of N days is a billing period.

Dropped before choosing: Nth highest value with names (too close to rank-top-customers and pareto), first or last word of a cell (too close to split-full-name and split-sku), round prices to .99 (round-prices-nearest-nickel-99 exists), late fee calculator (legal limits, judgment), tip pool split (wage-law adjacent), next Friday payout date (too close to week-ending-friday).

Function families: LOOKUP/INDEX/SUMPRODUCT(MAX) (1), ROUNDUP/MAX/IF (2), FIND/SUBSTITUTE/SUMPRODUCT text matching (3), ROUND/INT arithmetic with a reconciliation total (4), INT date arithmetic with COUNTIF/SUMIFS (5). No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays.

How to verify in LibreOffice (all five): build the .xlsx with openpyxl, recalculate headless (soffice --headless --convert-to xlsx), read the values back and compare with an independent Python model; also type the edge cases in the Test lines. Excel and Google Sheets are not run by us; the posts say LibreOffice Calc only unless someone actually tests there.

Ranked best first:

1. How to Look Up the Most Recent Price Paid for an Item in Excel and Google Sheets
   - Slug: look-up-most-recent-price-paid-last-match-lookup-excel-google-sheets
   - Search phrase: "excel lookup last matching value", "get most recent price for item google sheets", "vlookup last match instead of first", "last purchase price spreadsheet".
   - Template: tidy-tabs-last-price-paid.xlsx (Purchases, Lookup, How to use). Purchases: header row 3, data rows 4 to 203 (13 sample purchases, fictional): A Date | B Item | C Unit cost | D Item (trimmed) | E Date as number. Lookup: header row 3, items typed in A4:A23, B Purchases found | C Last price entered | D On Purchases row | E Date on that row | F Newest date for this item | G Status; summary I3:J9.
     - C4 `=IF(B4="","",IF(B4=0,"No purchase",LOOKUP(2,1/((Purchases!$D$4:$D$203=TRIM(A4))*ISNUMBER(Purchases!$C$4:$C$203)),Purchases!$C$4:$C$203)))`; F4 uses `SUMPRODUCT(MAX(match*Purchases!$E$4:$E$203))` with E `=N(A4)`; G4 says OK or Newer date higher up.
     - Expected: Soy wax, 10 lb $61.00; Kraft mailer 6x6x4 $0.71; tissue paper (typed in lower case with spaces) $18.90 with Newer date higher up; Lavender oil $15.10; 8 oz tin $1.20; Wax melt molds No purchase.
   - Test: first, middle and last match, one match, none, capitals and spaces, blank or text cost on the last row (skipped), text or blank dates, same date twice, 200 rows, 20 items, random logs against a Python model.
   - Who ranks: Exceljet last match, Microsoft LOOKUP help, Stack Overflow. Angle: price log for supplies, no sorting needed, and a flag when the last row is not the newest date.

2. How to Round an Order Quantity Up to a Case Size and Minimum Order in Excel and Google Sheets
   - Slug: round-order-quantity-up-to-case-size-minimum-order-roundup-excel-google-sheets
   - Search phrase: "round up to nearest case size excel", "order quantity multiple of pack size formula", "minimum order quantity formula excel", "ceiling to multiple of 12 google sheets".
   - Template: tidy-tabs-order-quantity-case-size.xlsx (Order, How to use). Header row 3, rows 4 to 43 (8 sample items): A Item | B Need | C Case size | D Minimum order | E Cases | F Units to order | G Extra units | H Unit cost | I Order cost | J Row check; summary L3:M9.
     - E4 `=IF(J4="OK",IF(B4=0,0,ROUNDUP(MAX(B4,D4)/C4,0)),"")`; F4 `=E4*C4`; G4 `=F4-B4`; I4 `=ROUND(F4*H4,2)`.
     - Expected: mailers need 130, case 25, minimum 100 gives 6 cases, 150 units, $93.00; candle tin need 340, case 24, minimum 240 gives 360; labels need 900, case 500, minimum 1000 gives 1000; totals 1,495 needed, 1,660 ordered, 165 extra, $914.60.
   - Test: need 0, need below and above the minimum, minimum that is not a multiple of the case, decimals, negatives, text, blank, case size 0 and 1.5, large numbers, random lists against a Python model.
   - Who ranks: Exceljet round to multiple, Microsoft CEILING and ROUNDUP help. Angle: one row per supply, minimum and case size together, shows the extra units you end up with. The numbers come from the supplier; the post does not suggest what to buy.

3. How to Count Orders With an Exact Tag in a Comma Separated Cell in Excel and Google Sheets
   - Slug: count-orders-with-exact-tag-comma-separated-cell-find-sumproduct-excel-google-sheets
   - Search phrase: "count cells containing exact word in comma separated list excel", "countif exact match within comma separated values", "count tag in cell google sheets", "countif whole word not partial".
   - Template: tidy-tabs-count-exact-tag.xlsx (Orders, How to use). Orders: header row 3, rows 4 to 203 (12 sample orders): A Order | B Tags | C Amount | D Tags cleaned; tag table F3:J13 (Tag to count, Orders with this tag, Amount of those orders, COUNTIF with * partial match, Difference); summary L3:M7.
     - D4 `=IF(TRIM(B4)="","",","&SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B4)),", ",",")," ,",",")&",")`; G4 `=IF(TRIM(F4)="","",IF(ISNUMBER(FIND(",",F4)),"Check",SUMPRODUCT(--ISNUMBER(FIND(","&LOWER(TRIM(F4))&",",$D$4:$D$203)))))`; I4 `=IF(ISNUMBER(G4),COUNTIF($B$4:$B$203,"*"&TRIM(F4)&"*"),"")`.
     - Expected: gift 4 orders ($109.50) against 7 with COUNTIF; gift wrap 2; rush 3; repeat 3; wholesale 2; personalized 2.
   - Test: gift against gifted, a gift, gift wrap, gift, (comma), spaces around commas, capitals, the same tag twice in one order, empty tags (,,), blank rows, tag containing a comma, tags with * ? ~, 200 rows, random orders against a Python model.
   - Who ranks: Exceljet count cells that contain, Ablebits, Stack Overflow. Angle: Etsy and craft order notes with several tags per order; shows the over-count of COUNTIF side by side.

4. How to Convert Cents to Dollars in a Payments Export in Excel and Google Sheets
   - Slug: convert-cents-to-dollars-payments-export-round-divide-100-excel-google-sheets
   - Search phrase: "convert cents to dollars excel", "amount in cents to dollars google sheets", "divide by 100 payment export", "stripe export amount in cents spreadsheet".
   - Template: tidy-tabs-cents-to-dollars.xlsx (Payments, How to use). B1 payout reported (fictional 1160.94), header row 3, rows 4 to 203 (8 sample payments): A Date | B Order | C Amount (cents) | D Fee (cents) | E Amount ($) | F Fee ($) | G Net ($) | H Row check; dollars to cents block J3:K13; summary M3:N10.
     - E4 `=IF(H4="OK",ROUND(C4/100,2),"")`; G4 `=IF(H4="OK",ROUND(E4-F4,2),"")`; K4 `=IF(J4="","",IF(ISNUMBER(J4),ROUND(J4*100,0),"Check"))`.
     - Expected: 1850 cents is $18.50 with an 84 cent fee, net $17.66; 5 cents is $0.05; refund -2400 is -$24.00; 100000 is $1,000.00; net total $1,160.94 matches B1, difference $0.00.
   - Test: 0, 1, 5, 99, negative, 10 digits, fraction of a cent (Check), text, text number "1850" (Check), blank fee, text fee, empty list, 200 rows, payout off by one cent, dollars to cents for 2-decimal values, random lists against an exact Decimal model.
   - Who ranks: generic Microsoft and Google formula help pages, forum answers. Angle: a payments export with fee and net columns and a payout check, not just `/100`. Fee numbers are made up; the post does not state any processor's fees.

5. How to Group Dates Into 14-Day Billing Periods in Excel and Google Sheets
   - Slug: group-dates-into-14-day-billing-periods-int-excel-google-sheets
   - Search phrase: "group dates into 14 day periods excel", "biweekly period from date formula", "date to period number excel", "sum by two week period google sheets".
   - Template: tidy-tabs-billing-periods.xlsx (Entries, How to use). B1 first period starts (11/02/2026), B2 days in each period (14), header row 4, rows 5 to 204 (13 sample entries): A Date | B What | C Amount | D Period no. | E Period starts | F Period ends; period table H4:L16 (12 periods: Period no., Starts, Ends, Entries, Amount); summary rows H17:L22.
     - D5 `=IF(A5="","",...IF(A5<$B$1,"Before start",INT((A5-$B$1)/$B$2)+1))` (the guards return Check for a text date or a bad B1 or B2); E5 `=IF(ISNUMBER(D5),$B$1+(D5-1)*$B$2,"")`; K5 `=COUNTIF($D$5:$D$204,H5)`; L5 `=SUMIFS($C$5:$C$204,$D$5:$D$204,H5)`.
     - Expected: 11/15/2026 is period 1 (11/02 to 11/15), 11/16/2026 is period 2, 12/20/2026 is period 4; 10/30/2026 is Before start; periods 1 to 4 hold 4, 3, 3 and 2 entries, $655.00, $435.00, $585.00, $145.00; total $1,940.00 with $120.00 before the start.
   - Test: a date on the first day and last day of a period, the day before the start, period length 1, 7, 30, 0, negative, 14.5, text and blank settings, dates with a time of day, leap day, year change, text dates, blank rows, 200 rows, random sets against a Python model.
   - Who ranks: Exceljet group by, Microsoft date help. Angle: one start date and one length cell, no helper tables; it groups dates and does not decide when anyone must bill or pay.
