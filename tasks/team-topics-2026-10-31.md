# Team topics - 2026-10-31 queue (5 new post topics)

Posts dated 2026-10-31 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc 24.2, a download that gives a reason to visit, no tax, wage-law or legal judgment). Each was checked by grep against slugs and titles in posts/ready and posts/published, tasks/topics-seed.md and every tasks/team-topics-*.md. Grepped for: deposit, balance due, consign, payout, commission, gift card, store credit, landed cost, freight, bundle, kit, cash flow, opening balance, days of stock or supply, stale or slow-moving, lookup to the left, entire row. Hits were different jobs (see below). No topic is a duplicate. Considered and dropped: gift card or store credit balances (unredeemed balances can carry legal rules, so it fails gate 3), late fee calculator (legal flavor), tip pooling (wage-law flavor), landed cost with freight split (same method as the 10-23 split-shared-booth-fee post), highlight the entire row by status (already shown in the budget-vs-actual variance post), banded rows and print titles (covered by the freeze-header and one-page print posts).

Closest existing posts and how these differ: running-balance-column and months-to-pay-off (a single running total; topic 5 buckets dated expected money in and out into weeks and finds the lowest week); invoice-due-date-payment-terms, receivables-aging-report and average-days-to-get-paid (invoices only; topic 1 tracks several part payments per custom order and a balance that can be zero, open, past due or overpaid); marketplace-fee-net-payout-calculator (one sale, fee percent; topic 2 sums many sales per shop at a different commission rate each); reorder-point-calculator and sell-through-rate (reorder level from lead time; percent sold; topic 4 is days of stock left from the last 30 days of sales, with a run-out date); vlookup-price-list and two-way-rate-card-lookup (return a column to the right of the key; topic 3 returns a value from a column LEFT of the key, which VLOOKUP cannot do).

Function families vary from the last 3 days (ROUNDUP, COUNTIF/INDEX/MATCH multi-match, MOD, MAXIFS, SUMIFS with EDATE on 10-30; IFERROR, ROUND, RAND/RANK, AVERAGEIF, IFS on 10-29; IPMT/PPMT, EOMONTH, ABS/MIN, MROUND/TEXT, INT/MOD/INDEX on 10-28): SUMIF with status IFs, SUMIFS plus a per-shop rate lookup, INDEX/MATCH to the left with TRIM, SUMIFS on a date window with a typed as-of date, SUMIFS weekly buckets with MIN. Mix: two trackers with sample data (1, 5), one formula how-to (3), two calculators fed by a log (2, 4). No FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TODAY, NOW, dynamic arrays or WEBSERVICE. No tax, wage, contract or financial-planning advice: arithmetic on fictional numbers only.

Formulas were NOT run in LibreOffice while picking these topics. Nothing below is a recorded test. Each topic's Verify line says what the build must run; only claim LibreOffice in `tested_in` after that run. Excel and Google Sheets are not tested for any topic; link the Microsoft and Google help pages for the functions.

Ranked best first. Titles are final picks from the hook-writer step; opening paragraphs are the answer-first paragraph for the script-writer to open the post with. The openings state formulas that the build must confirm before publishing.

---

## 1. Track deposits and balance due on custom orders

- Title: How to Track Deposits and Balance Due on Custom Orders in Excel and Google Sheets
- Slug: `track-custom-order-deposits-balance-due-sumif-excel-google-sheets`
- Opening paragraph: Add up what each customer has paid with `=SUMIF(Payments!$A$5:$A$104,A5,Payments!$C$5:$C$104)`, then subtract it from the order total with `=ROUND(D5-F5,2)` to get the balance due. Keep payments on their own sheet so a $50.00 deposit and a later $120.00 final payment both count toward the same order number.
- Other title options considered: How to Track Customer Deposits and Remaining Balances in Excel and Google Sheets; How to Show What a Customer Still Owes You in Google Sheets
- Search wording people type: "deposit and balance due spreadsheet", "track partial payments excel", "customer balance remaining formula excel", "custom order tracker google sheets", "sumif payments by order number"
- Template spec: `tidy-tabs-custom-order-deposits.xlsx` (Orders, Payments, Settings, How to use). All people and orders are fictional; the post says so once.
  - Settings (yellow): B1 As-of date `10/31/2026` (typed, not TODAY, so the status does not change when the file is opened later; the post mentions TODAY() as the swap and its effect).
  - Payments (yellow input, A4:D4 headers, rows 5-104): Order # (`CO-1001`), Date, Amount, Method (Cash, Venmo, Card, Check). 20 fictional rows, for example a $60.00 deposit on `CO-1003` on 10/05/2026 and a $140.00 final payment on 10/27/2026.
  - Orders (A4:I4 headers, rows 5-64, yellow input A:E): Order # `CO-1001`, Customer, Item (for example "Wedding sign, cedar", "Cake topper set", "Embroidered tote x 12"), Order total (`$` amounts from $45.00 to $480.00), Due date. 12 fictional rows.
    - F Paid so far `=IF(A5="","",SUMIF(Payments!$A$5:$A$104,A5,Payments!$C$5:$C$104))`
    - G Balance due `=IF(A5="","",ROUND(D5-F5,2))`
    - H Status `=IF(A5="","",IF(G5<0,"Overpaid",IF(G5=0,"Paid in full",IF(F5=0,IF(E5<Settings!$B$1,"Past due","Awaiting deposit"),IF(E5<Settings!$B$1,"Past due","Deposit received")))))`
    - I Deposit percent `=IF(OR(A5="",D5=0),"",ROUND(F5/D5,4))` formatted 0%
  - Summary (Orders K1:L8): orders count `=COUNTA(A5:A64)`, total ordered `=SUM(D5:D64)`, total received `=SUM(F5:F64)`, total balance due `=ROUND(SUM(G5:G64),2)`, past due count `=COUNTIF(H5:H64,"Past due")`, past due balance `=SUMIF(H5:H64,"Past due",G5:G64)`, payments not matching any order `=SUMPRODUCT((Payments!A5:A104<>"")*(COUNTIF(A5:A64,Payments!A5:A104)=0))` (if that SUMPRODUCT form fails in the build, use a helper column on Payments `=IF(A5="","",IF(COUNTIF(Orders!$A$5:$A$64,A5)=0,"No order",""))` and COUNTIF it).
  - CHECK cell: `=IF(AND(ROUND(L3-L4-L5,2)=0,L8=0),"OK","Check")` (ordered minus received equals balance due, and every payment has an order).
  - Sample must show all five statuses: Paid in full, Deposit received, Awaiting deposit, Past due, Overpaid (one order overpaid by $5.00 to show the negative balance), plus one payment with a typo order number to trip the unmatched count in a separate edited copy (the committed sample stays OK).
  - The post explains arithmetic only. It does not say when a deposit is refundable or what a contract requires.
- Verify plan: Python model with `Decimal` cents for the sample and 12 edited copies (order with 3 part payments, no payments, payment exactly equal to total, payment above total, due date equal to the as-of date stays not Past due, due date one day earlier is Past due, order number with a trailing space or different capitalization (SUMIF ignores case, not spaces), payment with a typo order number, blank rows in the middle, a text amount "60", a total of $0.00). LibreOffice headless: recalculate every copy, read with `data_only`, compare every Paid, Balance, Status and summary cell, confirm no error values. Note the TRIM/space and text-amount behavior found and add guards the same way the 10-30 builds did (ISNUMBER on the amount). Not run: Excel, Sheets, TODAY().
- Page-1 competitors angle: Page 1 will likely be Microsoft invoice and payment tracker templates, Vertex42, Smartsheet and template-gallery sites that sell or gate whole workbooks. Angle: one small payments sheet plus a SUMIF per order, a typed as-of date so the status is reproducible, and a CHECK that catches payments typed against a wrong order number. For makers who take a deposit before they start.

---

## 2. Calculate consignment payouts for each shop

- Title: How to Calculate Consignment Payouts for Each Shop in Excel and Google Sheets
- Slug: `consignment-payout-by-shop-sumifs-excel-google-sheets`
- Opening paragraph: Multiply each sale by the shop's commission rate, round to cents with `=ROUND(D5*E5*rate,2)`, and add the results per shop with `=SUMIFS(Sales!$H$5:$H$204,Sales!$B$5:$B$204,A5)`. What you are owed from a shop is its sales total minus its cut minus anything it already paid you.
- Other title options considered: How to Track What Each Consignment Shop Owes You in Google Sheets; How to Split Sales Between You and a Shop With a Commission Percent in Excel
- Search wording people type: "consignment sales tracker spreadsheet", "consignment commission calculator excel", "how much does the shop owe me consignment", "sumifs by store name google sheets", "artist consignment payout spreadsheet"
- Template spec: `tidy-tabs-consignment-payouts.xlsx` (Shops, Sales, How to use). Fictional shops and items; the post says so once.
  - Shops (yellow input A5:C14, up to 10 shops): Shop name, Commission rate (the shop keeps), Paid to you so far. 3 filled rows: "Maple Street Gifts" 30%, "Harbor Art Co-op" 40%, "Cedar & Stone Boutique" 35%, with paid-so-far `$150.00`, `$0.00`, `$85.00`. Commission cells formatted 0%.
  - Sales (yellow input A:E, header row 4, rows 5-204): Sale date, Shop, Item (for example "Soy candle, large", "Ceramic mug", "Greeting card 6-pack"), Price each, Qty. 30 fictional rows between 09/01/2026 and 10/28/2026.
    - F Gross `=IF(A5="","",ROUND(D5*E5,2))`
    - G Shop's cut `=IF(A5="","",IF(COUNTIF(Shops!$A$5:$A$14,B5)=0,"Check shop",ROUND(F5*INDEX(Shops!$B$5:$B$14,MATCH(B5,Shops!$A$5:$A$14,0)),2)))`
    - H Your share `=IF(ISNUMBER(G5),ROUND(F5-G5,2),"")`
  - Shops summary columns (D:H on the Shops sheet): Gross sales `=SUMIFS(Sales!$F$5:$F$204,Sales!$B$5:$B$204,A5)`, Shop's cut `=SUMIFS(Sales!$G$5:$G$204,Sales!$B$5:$B$204,A5)`, Your share `=SUMIFS(Sales!$H$5:$H$204,Sales!$B$5:$B$204,A5)`, Still owed to you `=ROUND(F5-C5,2)` (the your-share column minus paid so far, shown negative-safe with a Check flag when the shop has been paid more than it owes), Items sold `=SUMIFS(Sales!$E$5:$E$204,Sales!$B$5:$B$204,A5)`. Totals row.
  - Per-line rounding is deliberate: the cut is rounded on each sale so the cut plus your share equals the gross to the cent; the post links to the 10-29 totals-off-by-a-cent post.
  - CHECK cell: `=IF(AND(ROUND(SUM(Sales!F5:F204)-SUM(Sales!G5:G204)-SUM(Sales!H5:H204),2)=0,COUNTIF(Sales!G5:G204,"Check shop")=0),"OK","Check")`.
  - Post says nothing about agreement terms, ownership of unsold stock, or sales tax collected by the shop.
- Verify plan: Python `Decimal` ROUND_HALF_UP model for the sample and 12 edited copies (commission 0% and 100%, a sale with price $12.50 x 3 at 35% that rounds on a half cent, a shop name typed with a different capitalization (MATCH is not case-sensitive) and with a trailing space (must show Check shop, not a wrong rate), a shop with no sales, a shop paid more than owed, qty blank, price as text, 200 rows). LibreOffice headless to recalculate and compare per-shop sums and the CHECK. Not run: Excel, Sheets.
- Page-1 competitors angle: Page 1 will likely be consignment inventory app and software pages, a few template sites (Vertex42-style), and forum threads. Angle: no app, one Sales log, one commission lookup per sale and a SUMIFS per shop, with a visible "still owed to you" number for the artist or maker who sells through several shops.

---

## 3. Look up a value to the left of the key column

- Title: How to Look Up a Value to the Left in Excel and Google Sheets
- Slug: `look-up-value-to-the-left-index-match-excel-google-sheets`
- Opening paragraph: Use `=INDEX($A$5:$A$44,MATCH(E5,$B$5:$B$44,0))` to return the product name from column A when the SKU you are searching sits in column B. VLOOKUP only returns columns to the right of the search column, so INDEX and MATCH fills the gap in both Excel and Google Sheets.
- Other title options considered: How to Use INDEX and MATCH When VLOOKUP Can't Look Left in Excel and Google Sheets; How to Find a Product Name From a SKU When the SKU Column Is on the Right
- Search wording people type: "vlookup to the left", "excel lookup column to the left of the search column", "index match instead of vlookup", "google sheets vlookup return value from column before", "find name from sku excel"
- Template spec: `tidy-tabs-lookup-to-the-left.xlsx` (Products, Lookup, How to use). Fictional shop catalog.
  - Products (yellow input A4:D4 headers, rows 5-44): Product name, SKU, Price, Stock. 25 fictional rows for a candle and soap shop, name in A, SKU in B (`CND-001`, `SOP-014`), price in C, stock in D. Column order deliberately puts the key to the right of the name.
  - Lookup (A4:H4 headers, rows 5-64, yellow input A): SKU typed or pasted, 15 fictional lookups. Include two SKUs not in the list, one with a trailing space (`CND-004 `), one lower case (`snd-010`), one blank row in the middle.
    - B Product name (INDEX/MATCH left lookup) `=IF(A5="","",IFERROR(INDEX(Products!$A$5:$A$44,MATCH(TRIM(A5),Products!$B$5:$B$44,0)),"Not found"))`
    - C Price (right of the key, the same lookup form) `=IF(A5="","",IFERROR(INDEX(Products!$C$5:$C$44,MATCH(TRIM(A5),Products!$B$5:$B$44,0)),"Not found"))`
    - D Stock, same form on column D.
    - E "VLOOKUP attempt to the left" demo `=IF(A5="","",IFERROR(VLOOKUP(A5,Products!$A$5:$D$44,1,FALSE),"Not found"))` with a header note: this searches column A, so a SKU is never found. Column F shows `=IF(A5="","",IFERROR(VLOOKUP(TRIM(A5),Products!$B$5:$D$44,2,FALSE),"Not found"))` as the right-side form that works for price, so the reader sees what VLOOKUP can and cannot do.
    - G Row found `=IF(A5="","",IFERROR(MATCH(TRIM(A5),Products!$B$5:$B$44,0),"Not found"))` (position in the list, to explain what MATCH returns).
  - Summary: lookups entered `=COUNTA(A5:A64)`, found `=COUNTIF(B5:B64,"<>Not found")-COUNTBLANK-style` (use `=SUMPRODUCT((A5:A64<>"")*(B5:B64<>"Not found"))` if the simpler form miscounts blanks, and confirm in the build), not found count, duplicate SKUs in Products `=SUMPRODUCT((Products!B5:B44<>"")*(COUNTIF(Products!B5:B44,Products!B5:B44)>1))` (MATCH returns the first one only, so the post warns).
  - CHECK cell: `=IF(AND(found+notfound=entered,duplicate SKUs=0),"OK","Check")` with the exact cell addresses fixed by the build.
- Verify plan: Python dictionary lookup for the sample and 12 edited copies (key in the first row, last row, not in the list, lower case, trailing space, leading apostrophe or SKU stored as a number versus text (`1001` against `"1001"`), duplicate SKU (first match wins), blank key, a name that contains a comma, list with a blank row inside). LibreOffice headless to confirm INDEX/MATCH results and that the VLOOKUP-left demo returns Not found for every SKU. Post may say that XLOOKUP does this in newer Excel and Google Sheets and that the team did not test it. Note: MATCH with 0 is case-insensitive in Calc; check and state it as the build finds it. Not run: Excel, Sheets.
- Page-1 competitors angle: Page 1 will likely be Exceljet and Ablebits INDEX/MATCH tutorials, Microsoft VLOOKUP help, Stack Overflow, and Google's own VLOOKUP help. Angle: a shop catalog where the SKU column is on the right, a side-by-side that shows exactly what VLOOKUP returns for the left case, and a TRIM and Not found guard for pasted SKUs.

---

## 4. Days of stock left from recent sales

- Title: How to Calculate Days of Stock Left From Recent Sales in Excel and Google Sheets
- Slug: `days-of-stock-left-from-recent-sales-sumifs-excel-google-sheets`
- Opening paragraph: Divide the units on hand by the average units sold per day, using `=ROUND(C5/E5,0)` where E5 is the last 30 days of sales divided by 30. Pull the 30-day total from your sales log with `=SUMIFS(Sales!$C$5:$C$204,Sales!$B$5:$B$204,A5,Sales!$A$5:$A$204,">"&Settings!$B$1-Settings!$B$2,Sales!$A$5:$A$204,"<="&Settings!$B$1)`.
- Other title options considered: How to Estimate When You Will Run Out of a Product in Excel and Google Sheets; How to Find Out How Many Days of Inventory You Have Left in Google Sheets
- Search wording people type: "days of inventory on hand formula excel", "how many days until I run out of stock spreadsheet", "days of supply calculator", "stock cover days google sheets", "average daily sales sumifs last 30 days"
- Template spec: `tidy-tabs-days-of-stock-left.xlsx` (Products, Sales, Settings, How to use). Fictional shop.
  - Settings (yellow): B1 As-of date `10/31/2026` (typed, not TODAY), B2 Days of sales to look back `30`, B3 Warn when days left is below `21`.
  - Sales (yellow input, header row 4, rows 5-204): Date, SKU, Qty sold. 60 fictional rows from 09/20/2026 to 10/30/2026, so some rows fall outside the 30-day window (09/20 to 09/30 rows must not count).
  - Products (yellow input A:C, header row 4, rows 5-44): SKU, Product, On hand. 12 fictional rows (for example `TEE-S` "Logo tee, small", `MUG-01` "Enamel mug", `STK-10` "Sticker pack"), on hand from 0 to 240.
    - D Sold in window `=IF(A5="","",SUMIFS(Sales!$C$5:$C$204,Sales!$B$5:$B$204,A5,Sales!$A$5:$A$204,">"&Settings!$B$1-Settings!$B$2,Sales!$A$5:$A$204,"<="&Settings!$B$1))`
    - E Average per day `=IF(A5="","",ROUND(D5/Settings!$B$2,2))`
    - F Days left `=IF(A5="","",IF(E5=0,"No recent sales",INT(C5/E5)))` (rounded down, stated in the post)
    - G Runs out about `=IF(ISNUMBER(F5),Settings!$B$1+F5,"")` formatted MM/DD/YYYY
    - H Status `=IF(A5="","",IF(C5=0,"Out of stock",IF(F5="No recent sales","No recent sales",IF(F5<Settings!$B$3,"Order soon","OK"))))`
  - Summary: products counted, Order soon count `=COUNTIF(H5:H44,"Order soon")`, Out of stock count, fewest days left `=MIN(F5:F44)` with the product `=INDEX(B5:B44,MATCH(MIN(F5:F44),F5:F44,0))` (guarded when no numbers).
  - CHECK cell: units sold in window summed over products equals `SUMIFS` over all of Sales in the same window (flags sales rows whose SKU is not in Products).
  - Post says: this is a simple average and ignores seasons and lead time; link to the reorder-point post for lead time. No forecasting claims.
- Verify plan: Python model with `datetime` for the sample and 12 edited copies (sale dated exactly 30 days before the as-of date is excluded, 29 days before is included, sale dated on the as-of date is included, sale dated after the as-of date is excluded, window 7 and 60 days, product with zero sales, product with zero stock, SKU in Sales not in Products, SKU with a different capitalization, text date in Sales, on hand blank). LibreOffice headless to recalculate and compare D:H and the summary. Confirm the `">"&date` criteria concatenation works in Calc the same way as in the 10-30 SUMIFS builds. Not run: Excel, Sheets, TODAY().
- Page-1 competitors angle: Page 1 will likely be inventory-software blogs (days of inventory on hand, accounting definitions), Exceljet SUMIFS pages and Vertex42 inventory templates. Angle: a sales log of single-product units for a small shop, a window set in a cell, a run-out date, and an Order soon flag; not an accounting ratio.

---

## 5. Weekly cash flow forecast with the lowest balance

- Title: How to Build a Weekly Cash Flow Forecast in Excel and Google Sheets
- Slug: `weekly-cash-flow-forecast-lowest-balance-sumifs-excel-google-sheets`
- Opening paragraph: List each expected payment in or out with a date, then total each week with `=SUMIFS(Items!$D$5:$D$104,Items!$C$5:$C$104,"In",Items!$A$5:$A$104,">="&A5,Items!$A$5:$A$104,"<"&A5+7)`. Ending balance is the opening balance plus the week's money in minus money out, and `=MIN(F5:F17)` shows the lowest point.
- Other title options considered: How to Forecast Cash Week by Week for a Small Business in Google Sheets; How to Find the Week You Might Run Short of Cash in Excel
- Search wording people type: "13 week cash flow template excel", "weekly cash flow forecast google sheets", "small business cash flow spreadsheet", "ending balance each week formula", "sumifs by week start date"
- Template spec: `tidy-tabs-weekly-cash-flow.xlsx` (Items, Forecast, How to use). Fictional freelancer/shop.
  - Items (yellow input, header row 4, rows 5-104): Date, Description, Type (dropdown In or Out), Amount (always positive). 30 fictional rows between 11/02/2026 and 01/29/2027: "Client invoice 2041" In $1,800.00, "Studio rent" Out $650.00 (monthly), "Etsy payout" In, "Supplies" Out, "Insurance payment" Out $210.00, and similar; plus one row dated before the first week and one row dated after the 13th week to exercise the CHECK in an edited copy only (committed sample stays OK).
  - Forecast: B1 Opening balance `$2,400.00` (yellow), B2 First Monday `11/02/2026` (yellow), B3 Low-balance line `$500.00` (yellow, a reader-chosen number, not advice). Rows 6-18 (13 weeks): A week start `=IF(ISNUMBER($B$2),$B$2+7*(ROWS($A$6:A6)-1),"")`, B week end `=A6+6`, C Money in `=SUMIFS(Items!$D$5:$D$104,Items!$C$5:$C$104,"In",Items!$A$5:$A$104,">="&A6,Items!$A$5:$A$104,"<"&A6+7)`, D Money out the same with "Out", E Net `=ROUND(C6-D6,2)`, F Ending balance `=ROUND(IF(ROWS($A$6:A6)=1,$B$1,F5)+E6,2)` (first row refers to B1 explicitly so the header text is never used), G Flag `=IF(F6<$B$3,"Below line","")`. Totals row 19.
  - Summary (I1:J7): lowest ending balance `=MIN(F6:F18)`, week it happens `=INDEX(A6:A18,MATCH(MIN(F6:F18),F6:F18,0))` (first week if tied), weeks below the line `=COUNTIF(G6:G18,"Below line")`, ending balance after week 13 `=F18`, items outside the 13 weeks `=COUNTIFS(Items!A5:A104,"<"&B2)+COUNTIFS(Items!A5:A104,">="&B2+91)`, rows with a missing Type or Amount `=SUMPRODUCT((Items!A5:A104<>"")*((Items!C5:C104<>"In")*(Items!C5:C104<>"Out")+(1-ISNUMBER(Items!D5:D104))>0))` (the build may replace it with a helper column if that form is fragile).
  - CHECK cell: `=IF(AND(ROUND(B1+SUM(C6:C18)-SUM(D6:D18)-F18,2)=0,outside=0,bad rows=0),"OK","Check")`.
  - The post states that the forecast is only as good as the dates and amounts typed in and gives no financial, tax or lending advice.
- Verify plan: Python model with `datetime` and `Decimal` for the sample and 12 edited copies (item on the first and last day of the 13-week window, item dated one day after the window, item before the first Monday, a first-day date that is not a Monday (weeks still run 7 days from it), Type typed "in" lowercase (SUMIFS is case-insensitive in Calc) and "Income" (must trip the bad-rows count), a negative amount, a week with no items, opening balance negative, all weeks below the line, ties for the lowest balance). LibreOffice headless to recalculate and compare every C:G cell and the summary; confirm there are no error values. Not run: Excel, Sheets, a chart (none is promised).
- Page-1 competitors angle: Page 1 will likely be Vertex42, Smartsheet and other template galleries for "13 week cash flow", plus accounting-software blogs. Angle: a short Items list rather than a locked-grid template, week buckets by SUMIFS with the formula shown, and the lowest balance and its week called out, for a freelancer or maker with a handful of expected payments.
