# Team topics - 2026-10-16 queue (5 new post topics)

Posts dated 2026-10-16 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us, a download that gives a reason to visit, no tax, insurance or legal judgment). Each was checked against every title and slug in posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md and the names in templates/.

Closest existing posts and how these differ: sales-by-day-of-week-sumif and count-market-saturdays (sum or count by weekday name, this one moves a date to the Friday that ends its week), monthly-sales-goal-progress (a straight-line pace that its post says is not a forecast, this one fits a linear trend with FORECAST), rolling-3-month-average (smooths past months, this one projects the next), repeat-customer-rate and last-order-date-per-customer (order behavior, this one measures how long a client has been on the list), subscription-renewal-tracker (days until a renewal, this one counts elapsed years and months), clean-customer-list-trim-proper and compare-two-lists (fix text or compare lists, this one counts empty required fields per row), cheapest-supplier-cost-per-ounce (mentions free shipping only as a limit, this one is the threshold check itself), invoice-due-date-payment-terms (due date from terms, this one is a week-ending date for weekly billing).

Function families vary: WEEKDAY/MOD date math, IF/AND/MAX logic, FORECAST, DATEDIF, COUNTBLANK. Formulas were run once in LibreOffice Calc 24.2 headless while picking topics (MOD(6-WEEKDAY(A1,1),7) gave 10/16/2026 for 10/14, 10/16 stays, 10/17 gives 10/23; IF/AND/MAX; FORECAST gave 1630 for the points 1200, 1350, 1500, 1480; DATEDIF "y" and "ym"; COUNTBLANK). None uses FILTER, LET, LAMBDA or XLOOKUP. The template builds still need the full check in scripts/build-templates.py. Only claim LibreOffice in `tested_in`. The search results page is assumed to be led by Microsoft and Google help pages and big sites (not checked); each topic takes a small business angle.

Ranked best first. Titles are examples for the hook-writer, no post body or first paragraph yet.

1. Week ending date: `week-ending-friday-date-weekday-excel-google-sheets`
- Title: How to Find the Week Ending Date for Any Date in Excel and Google Sheets
- Other title options: How to Group Hours by Week Ending Friday in Excel and Google Sheets; How to Find the Friday of a Week in Excel and Google Sheets
- Search wording people type: "week ending date formula excel", "find friday of the week excel", "group by week ending date google sheets"
- Template: tidy-tabs-week-ending-date.xlsx (Log, Weeks, How to use). Work log with date and hours, week-ending date per row (the reader types which weekday ends the week, default Friday), a Weeks tab summing hours with SUMIFS per week ending, and a CHECK cell that hours add up to the log total.
- Verify: Python weekday math for 12 dates; a date that is already the week-ending day (stays the same); the day after it (jumps 6 days); a Sunday; a date typed as text; a blank date row; a year boundary (12/30/2026). Date math only, no pay period or overtime rules.
- Page 1 owner: Microsoft WEEKDAY help page, Exceljet and Stack Overflow answers. Angle: freelancer billing by week with a ready log and weekly totals.

2. Free shipping threshold: `free-shipping-threshold-amount-to-add-if-and-excel-google-sheets`
- Title: How to Check a Free Shipping Threshold and Show the Amount to Add in Excel and Google Sheets
- Other title options: How to Flag Orders That Qualify for Free Shipping in Excel and Google Sheets; How to Use IF and AND to Test Two Conditions in Excel and Google Sheets
- Search wording: "free shipping threshold formula excel", "if cart total over 50 excel", "if and function two conditions excel"
- Template: tidy-tabs-free-shipping-threshold.xlsx (Orders, How to use). Order subtotal, delivery type (Ship or Pickup), threshold cell the reader types, Free / Pay status, amount to add with MAX(0, ...), count qualifying, count within a typed "close" gap, a CHECK cell.
- Verify: Python comparison for 10 orders; subtotal exactly at the threshold (qualifies); one cent under; Pickup rows; subtotal typed as text; blank subtotal; threshold changed. Reader's own threshold, no carrier rates or marketplace rules.
- Page 1 owner: Microsoft IF/AND help pages and Exceljet. Angle: Etsy or Shopify-style seller testing a "free over $X" offer.

3. Forecast next month sales: `forecast-next-month-sales-forecast-linear-excel-google-sheets`
- Title: How to Forecast Next Month's Sales With FORECAST in Excel and Google Sheets
- Other title options: How to Project Sales From Past Months With a Trend Line in Excel and Google Sheets; How to Use FORECAST.LINEAR for a Small Business in Excel and Google Sheets
- Search wording: "forecast next month sales excel", "forecast function excel example", "sales projection google sheets"
- Template: tidy-tabs-forecast-next-month.xlsx (Sales, How to use). Month numbers and monthly sales for 8 months, FORECAST for the next 3 months, slope per month (SLOPE), a CHECK cell comparing the first projection to a hand calculation.
- Verify: Python least-squares fit on the sample data; flat sales (slope 0); falling sales; only 2 months; a month typed as text; a blank month in the middle; mismatched range lengths (#N/A shown as a Troubleshooting case). Use the older FORECAST name so LibreOffice recalculates it. The post says it is a straight-line estimate, not a prediction, and gives no advice.
- Page 1 owner: Microsoft FORECAST help page, Exceljet and large tutorial sites. Angle: 8 months of fictional shop sales with a download.

4. Client tenure: `client-tenure-years-months-datedif-excel-google-sheets`
- Title: How to Calculate How Long a Client Has Been With You in Years and Months in Excel and Google Sheets
- Other title options: How to Use DATEDIF to Count Years and Months Between Dates in Excel and Google Sheets; How to Show Customer Since Dates as Years and Months in Excel and Google Sheets
- Search wording: "datedif years and months excel", "calculate tenure in years and months", "how long since date excel"
- Template: tidy-tabs-client-tenure.xlsx (Clients, How to use). Client name, first project date, as-of date cell the reader types (so results do not change daily), years, months, text label like "3 yr 7 mo", total months, count of clients over a typed number of years.
- Verify: Python month-difference logic for 10 clients; a start date later in the month than the as-of date; same day exactly; start after as-of (DATEDIF returns #NUM!, shown as a Troubleshooting case); Feb 29 starts; a date typed as text. Note that DATEDIF is not listed in Excel's function picker, so the post links Microsoft's own DATEDIF page and explains that.
- Page 1 owner: Microsoft DATEDIF help page, Exceljet. Angle: freelancer or shop owner with a client list.

5. Find incomplete rows: `find-incomplete-customer-rows-countblank-excel-google-sheets`
- Title: How to Find Rows With Missing Information in a Customer List in Excel and Google Sheets
- Other title options: How to Count Blank Cells in Each Row in Excel and Google Sheets; How to Flag Customers With No Email or Address in Excel and Google Sheets
- Search wording: "count blank cells in row excel", "find rows with missing data excel", "flag empty cells google sheets"
- Template: tidy-tabs-incomplete-rows.xlsx (Customers, How to use). Name, email, phone, street, ZIP; blanks per row with COUNTBLANK; Complete / Missing label; a "Missing" column naming the first empty field; totals of complete rows and rows missing a ZIP; a CHECK cell.
- Verify: Python blank count for 12 rows; a cell holding only a space (COUNTBLANK treats it as not empty, shown as a Troubleshooting case); a cell with a formula returning "" (counted as blank in LibreOffice and Excel); an all-blank row; a fully complete row; leading zero ZIP kept as text. Counts empty cells only, no validation of email or address correctness.
- Page 1 owner: Microsoft COUNTBLANK help page and Exceljet. Angle: mailing list or order-prep cleanup before printing labels.
