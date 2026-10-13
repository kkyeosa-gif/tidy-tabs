# Team topics - 2026-10-14 queue (5 new post topics)

Posts dated 2026-10-14 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only
(buildable and checkable by us, a download that gives a reason to visit, no tax, insurance or legal judgment). Each was checked against every title and slug in posts/ready, posts/published,
tasks/topics-seed.md, tasks/team-topics-*.md and the names in templates/. Closest existing posts and how these differ: shipping-cost-weight-tier-lookup (looks up a rate from a weight, this one only converts ounces to pounds and ounces for the label),
business-days-ship-by-date (WORKDAY with a holiday list, this one counts one chosen weekday with NETWORKDAYS.INTL) and sales-by-day-of-week (sums sales by weekday, this one counts how many of that weekday fall in a month),
running-balance-column (builds the balance, this one pulls the latest value out of a log), split-sku-code and combine-address-columns (split or join text, this one measures text length and counts items in a cell),
time-to-decimal-hours (clock times to hours worked, this one shifts a date and time between US time zones).

Formulas were run once in LibreOffice Calc 24.2 headless while picking topics (INT/MOD text, LEN/SUBSTITUTE, LOOKUP(9.99E+307,...), A+TIME(3,0,0), NETWORKDAYS.INTL with "1111101"); none uses FILTER, LET, LAMBDA or XLOOKUP. The template builds still need the full check in scripts/build-templates.py. Only claim LibreOffice in `tested_in`. The search results page is assumed to be led by Microsoft and Google help pages and big sites (not checked); each topic takes a small business angle.

Ranked best first. Each entry has the topic, then the hook-writer's final title and first paragraph.

1. Ounces to pounds and ounces: `convert-ounces-to-pounds-and-ounces-int-mod-excel-google-sheets`
- Title: How to Convert Ounces to Pounds and Ounces in Excel and Google Sheets
- First paragraph: Get the whole pounds with `=INT(A2/16)` and the leftover ounces with `=MOD(A2,16)`, or join them with `=INT(A2/16)&" lb "&MOD(A2,16)&" oz"`. A 50 oz package shows as 3 lb 2 oz.
- Other title options: How to Show Package Weight as Pounds and Ounces in Excel and Google Sheets; How to Split Ounces Into Pounds and Ounces for Shipping Labels in Excel and Google Sheets
- Search wording people type: "convert ounces to pounds and ounces excel", "oz to lb oz formula", "shipping weight pounds and ounces google sheets"
- Template: tidy-tabs-ounces-to-pounds-ounces.xlsx (Weights, How to use). Item, packed ounces, pounds, leftover ounces, text label, total shipment weight, a CHECK cell that turns pounds and ounces back into ounces.
- Verify: Python divmod(oz,16) for 10 packages; exactly 16 oz (1 lb 0 oz); 15.5 oz; 0 and blank cells; text-typed numbers; a package over 100 lb. Arithmetic only, no carrier rules or rates.

2. Time zones for client calls: `convert-time-zones-client-calls-excel-google-sheets`
- Title: How to Convert Time Zones for Client Calls in Excel and Google Sheets
- First paragraph: Add the hour difference divided by 24 with `=A2+B2/24`, where A2 holds the call date and time and B2 holds the offset, such as -3 for Eastern to Pacific. A call at 10/14/2026 4:30 PM Eastern shows as 10/14/2026 1:30 PM Pacific.
- Other title options: How to Add Hours to a Date and Time in Excel and Google Sheets; How to Schedule a Call Across Eastern and Pacific Time in Excel and Google Sheets
- Search wording: "convert est to pst excel formula", "time zone converter google sheets", "add hours to datetime excel"
- Template: tidy-tabs-time-zone-converter.xlsx (Calls, Offsets, How to use). Offset table (Eastern, Central, Mountain, Pacific as typed hours), client zone dropdown, your time, client time, a Next day / Previous day flag.
- Verify: Python datetime plus timedelta for 8 calls; a call that crosses midnight in both directions; a 0 offset; a blank time; a time typed as text. Fixed hour offsets only: Daylight Saving Time changes and the zones that do not observe it are not handled, and the post says so and links the official source for rules.

3. Listing characters and tags: `count-listing-characters-tags-len-substitute-excel-google-sheets`
- Title: How to Count Characters and Tags in a Product Listing in Excel and Google Sheets
- First paragraph: Count characters with `=LEN(A2)` and count comma separated tags in one cell with `=LEN(B2)-LEN(SUBSTITUTE(B2,",",""))+1`. The title "Handmade Lavender Soy Candle 8 oz Gift for Her" is 46 characters long.
- Other title options: How to Check a Title Length Limit With LEN in Excel and Google Sheets; How to Count Items in a Comma Separated Cell in Excel and Google Sheets
- Search wording: "count characters in cell excel", "count comma separated values in a cell", "etsy title character counter spreadsheet"
- Template: tidy-tabs-listing-length-check.xlsx (Listings, How to use). Title, tags cell, character count, tag count, limit cells the reader types in, OK / Too long / Too many flags, a CHECK cell.
- Verify: Python len and str.count for 10 listings; an empty tags cell (the formula returns 1, handled with IF); a trailing comma; spaces after commas; a title exactly at the limit. The limits are typed by the reader from the marketplace's own help page; the post does not state them as facts.

4. Last number in a column: `last-number-in-column-lookup-excel-google-sheets`
- Title: How to Get the Last Number in a Column in Excel and Google Sheets
- First paragraph: Return the last number in a growing list with `=LOOKUP(9.99E+307,B2:B200)`, and the last date the same way from the date column. In a cash log whose balances end at $1,402.75, the formula returns $1,402.75 even with blank rows in between.
- Other title options: How to Pull the Latest Balance From a Running Log in Excel and Google Sheets; How to Find the Last Filled Cell in a Column in Excel and Google Sheets
- Search wording: "get last value in column excel", "last number in a column formula", "latest entry in a list google sheets"
- Template: tidy-tabs-last-number-in-column.xlsx (Log, Summary, How to use). Dated cash log with blank rows, a Summary tab with latest balance, latest date, count of entries, and an IFERROR version for an empty log.
- Verify: Python last non-blank numeric value; blank rows in the middle; a text note in the column (ignored); a number typed as text (ignored, shown as a Troubleshooting case); a 0 as the last entry (still returned); an empty column (#N/A, wrapped with IFERROR); a new row added below. Numbers only; the post says a text version needs a different formula.

5. Count market Saturdays: `count-market-saturdays-in-month-networkdays-intl-excel-google-sheets`
- Title: How to Count Market Saturdays in a Month With NETWORKDAYS.INTL in Excel and Google Sheets
- First paragraph: Count one weekday between two dates with `=NETWORKDAYS.INTL(A2,EOMONTH(A2,0),"1111101")`, where the seven digits mark Monday to Sunday and 0 is the day to count. For 10/01/2026 it returns 5 Saturdays.
- Other title options: How to Count How Many Saturdays Are in a Month in Excel and Google Sheets; How to Count a Certain Weekday Between Two Dates in Excel and Google Sheets
- Search wording: "count saturdays in a month excel", "count specific weekday between two dates", "networkdays.intl weekend string"
- Template: tidy-tabs-count-weekdays-in-month.xlsx (Months, How to use). Twelve months of 2026 with the first day, last day, Saturdays, Sundays, a booth fee per day, booth cost for the month, and a day-picker cell that changes the weekend string.
- Verify: Python weekday count for all 12 months (2026 Saturdays by month: 5,4,4,4,5,4,4,5,4,5,4,4); a month starting on a Saturday; February; a start date after the end date (negative count, shown as a Troubleshooting case); a date stored as text. Counts calendar days only, no booth rules or holidays.
