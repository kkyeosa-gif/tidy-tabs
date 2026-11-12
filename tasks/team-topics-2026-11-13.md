# Team topics - 2026-11-13 queue (5 new post topics)

Posts dated 2026-11-13. tasks/search-stats.md does not exist: 데이터 없음 (no data; no search volume or competition numbers, none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc, a download that gives a reason to visit, no tax, insurance or legal judgment).

Duplicate check done against every file name in posts/ready and posts/published (221 files), every title in tasks/topics-seed.md and the tasks/team-topics-*.md files (grep for duplicate, possible duplicate, unix, epoch, timestamp, in words, number to words, bill of materials, materials needed, shopping list, variation, option, label, feet and inches). Nothing matched as a topic. Closest existing posts and how these differ:
- Topic 1 vs highlight-duplicate-values (one column, conditional formatting), stop-duplicate-invoice-numbers (data validation that blocks typing), find-missing-invoice-numbers and remove-duplicate-rows (zauto): this one flags orders that match on three fields at once (customer, the day of a date-and-time stamp, amount), marks the first copy and every later copy, and only says "look at this"; it deletes nothing.
- Topic 2 vs how-many-batches-from-materials (what you can make from what you have, MIN/INT), scale-recipe-batch-quantities (one recipe times a factor), round-order-quantity-up-to-case-size (rounds one known quantity) and batch-cost-per-unit: this one starts from a list of orders, totals each material over every product that uses it through a recipe (bill of materials) table, and returns the shortfall, packs to buy and cost.
- Topic 3 vs convert-yyyymmdd-number-to-date, convert-24-hour-time-numbers, convert-time-zones-client-calls and time-to-decimal-hours: this one turns seconds or milliseconds since 01/01/1970 UTC (payment and API exports) into a local date and time, tells 10-digit seconds from 13-digit milliseconds, and flags cut-off values.
- Topic 4 vs count-words-in-cell (counts words), payment-reminder-message (builds a sentence with TEXT) and invoice-number-generator: this one writes a dollar amount out in words with the cents as a fraction (One Thousand Two Hundred Fifty and 00/100) for checks and receipts.
- Topic 5 vs extract-text-between-parentheses, split-sku-code-into-parts (fixed delimiters and positions), count-orders-with-exact-tag (counts tags) and pull-number-out-of-text: this one pulls the value that follows a named label (Size:, Scent:) from a variations string whose options come in any order, case and separator.

Dropped before choosing: convert feet-and-inches text such as 5' 7" to inches (too close to convert-decimal-inches-to-fractions, which already outputs feet and inches, and VALUE("1/2") turned into a date in a LibreOffice probe), build SKUs from name, color and size (reverse of split-sku and close to invoice-number-generator), average star rating from review counts (same SUMPRODUCT/SUM math as weighted-average-cost), late fee on overdue invoices (late fee limits are a legal question), mask account numbers to the last 4 digits (too thin for a template).

Function families: COUNTIFS with an expanding range, INT, COUNTIF (1); SUMIFS across three tabs, MAX, ROUNDUP, ROUND (2); DATE arithmetic, INT, SMALL, SUMPRODUCT/SUBSTITUTE digit test (3); INDEX on a word list, INT, MOD, TEXT, ROUND (4); FIND, MID, LOWER, SUBSTITUTE, CHAR, IFERROR (5). Themes: data check, production planning, import cleanup (dates), money text, text parsing. No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, MINIFS/MAXIFS, dynamic arrays, and no function that needs an _xlfn prefix.

How to verify in LibreOffice (all five): build the .xlsx with openpyxl, recalculate headless (soffice --headless --convert-to xlsx), read the values back and compare with an independent Python model; also type the edge cases listed per topic. Excel and Google Sheets are not run by us; the posts say LibreOffice Calc only unless someone actually tests there.

Ranked best first:

1. How to Flag Possible Duplicate Orders by Customer, Date and Amount in Excel and Google Sheets
   - Slug: 2026-11-13-team-01-flag-possible-duplicate-orders-same-customer-day-amount-countifs-excel-google-sheets
   - Search phrase: "find duplicate orders excel", "duplicate based on multiple columns excel", "countifs duplicates two columns google sheets", "flag double charged orders same day".
   - Template: tidy-tabs-possible-duplicate-orders.xlsx (Orders, How to use). Header row 3, rows 4 to 203 (15 fictional orders): A Order no. | B Order date and time | C Customer | D Amount | E Order day | F Same customer, day and amount | G Copy no. | H Flag | I Times this order no. is listed; summary K3:L9.
     - E4 core `=INT(B4)`; F4 core `=COUNTIFS($C$4:$C$203,C4,$E$4:$E$203,E4,$D$4:$D$203,D4)`; G4 core `=COUNTIFS($C$4:C4,C4,$E$4:E4,E4,$D$4:D4,D4)`.
     - Expected: Harbor Dental $86.50 at 9:14 AM and 9:15 AM on 11/02/2026 are First of 2 and Possible duplicate; the same customer and amount on 11/07/2026 is not flagged.
   - Test: wildcards * ? ~ and a leading < > = in names, upper and lower case, trailing spaces, minutes apart vs across midnight, date only, time only, text dates, text and fraction-of-a-cent amounts, negatives and 0, repeated order numbers (text and number), 200 identical rows, 8 random sets of up to 200 rows against a Python model.
   - Who ranks: Microsoft and Google COUNTIF/COUNTIFS help, Exceljet and Ablebits duplicate articles (mostly single-column or Remove Duplicates). Angle: a three-field check that keeps every row and marks first copy vs repeats.

2. How to Make a Materials Shopping List From Your Orders in Excel and Google Sheets
   - Slug: 2026-11-13-team-02-materials-shopping-list-from-orders-bill-of-materials-sumifs-roundup-excel-google-sheets
   - Search phrase: "bill of materials excel template small business", "calculate raw materials needed from orders excel", "materials needed calculator google sheets", "how much wax do I need for candles spreadsheet".
   - Template: tidy-tabs-materials-from-orders.xlsx (Materials, Orders, Recipe, How to use). Orders A4:C203 (For, Product, Quantity) + D Status; Recipe A4:C203 (Product, Material, Amount per item) + D Items ordered, E Material needed, F Status; Materials A5:E104 (Material, Unit, On hand, Pack size, Pack price) + F Needed, G Short, H Packs to buy, I Cost to buy, J Left after buying, K Status; summary M4:N9.
     - Recipe D4 core `=SUMIFS(Orders!$C$4:$C$203,Orders!$B$4:$B$203,A4)`; Materials F5 core `=SUMIFS(Recipe!$E$4:$E$203,Recipe!$B$4:$B$203,A5)`; H5 `=ROUNDUP(ROUND(G5/D5,9),0)`.
     - Expected: soy wax needed 595 oz, 400 on hand, 2 packs of 160 oz; six materials to buy for $307.00; the Holiday gift set line shows No recipe.
   - Test: product and material names with wildcards, case and trailing spaces, products with no recipe, materials missing from the Materials tab or listed twice, fractional quantities, negative and text amounts, on hand blank/text/negative, pack size 0/blank/negative, pack price blank/text/negative, floating cases (0.9/0.3, 0.7/0.1), 200 order lines, 8 random sets against an exact-fraction Python model.
   - Who ranks: MRP/inventory software pages and a few BOM template sites (often manufacturing jargon). Angle: a three-tab sheet for a maker selling at markets, wholesale and online, with packs to buy.

3. How to Convert a Unix Timestamp to a Date and Time in Excel and Google Sheets
   - Slug: 2026-11-13-team-03-convert-unix-timestamp-to-date-and-time-seconds-milliseconds-excel-google-sheets
   - Search phrase: "convert unix timestamp to date excel", "epoch time to date google sheets", "13 digit timestamp to date excel", "unix time milliseconds to date".
   - Template: tidy-tabs-unix-timestamp-to-date.xlsx (Timestamps, How to use). B1 hours from UTC (-5), header row 3, rows 4 to 203 (12 fictional payment IDs): A Payment ID | B Timestamp as exported | C Unit | D Date and time (local) | E Date | F Time | G Status; summary I3:J9.
     - Core UTC `=B4/86400+DATE(1970,1,1)`; with the offset `=(B4+$B$1*3600)/86400+DATE(1970,1,1)`; milliseconds `=B4/86400000+DATE(1970,1,1)`.
     - Expected: 1793630100 is 11/02/2026 9:35:00 AM at -5; 1793725200000 (milliseconds) is 11/03/2026 12:00:00 PM; 17945208 (cut off) shows Before 2000.
   - Test: 0, 1, -1, day boundaries, local midnight and 11:59:59 PM for several offsets, the seconds/milliseconds threshold, the 2100 limit, decimals, text digits with spaces and leading zeros, text with a minus, a decimal point, spaces inside or commas, ISO text, offsets -12, 14, 5.5, 0.25, -9.5, out of range and text, 8 random sets against a Python model.
   - Who ranks: Exceljet, Stack Overflow, epoch converter sites. Angle: one sheet that does seconds and milliseconds, a local offset and flags the values that are not timestamps.

4. How to Write a Check Amount in Words in Excel and Google Sheets
   - Slug: 2026-11-13-team-04-write-check-amount-in-words-dollars-and-cents-index-mod-excel-google-sheets
   - Search phrase: "convert number to words excel dollars and cents", "spell out numbers for checks excel", "number to words google sheets without script", "check amount in words formula".
   - Template: tidy-tabs-check-amount-in-words.xlsx (Checks, Words, How to use). Header row 3, rows 4 to 203 (12 fictional checks): A Check no. | B Date | C Pay to | D Amount | E Dollars | F Cents | G Amount in words | H Status; summary J3:K6. Words tab: B4:B23 (0 to 19) and E4:E13 (tens).
     - G4 splits the dollars into INT(E4/1000) and MOD(E4,1000) and writes each part with INDEX on the Words list; cents `TEXT(F4,"00")&"/100"`.
     - Expected: $1,250.00 is One Thousand Two Hundred Fifty and 00/100; $2,019.07 is Two Thousand Nineteen and 07/100; 12.345 shows Check.
   - Test: every dollar value 0 to 999 and every thousands group 0 to 999, all teens and tens, 1,000 / 1,001 / 1,010 / 1,100 / 100,100 / 999,999.99, 1,000,000, fractions of a cent, negatives, text, blank amounts, 6 random sets against an independent recursive Python speller.
   - Who ranks: Microsoft's SpellNumber VBA article, add-in vendors, Ablebits. Angle: no macro and no add-in, plain formulas that also open in Google Sheets.

5. How to Pull Size and Scent Out of an Order Variations Column in Excel and Google Sheets
   - Slug: 2026-11-13-team-05-pull-size-and-scent-from-variations-column-by-label-find-mid-excel-google-sheets
   - Search phrase: "extract text after a word excel", "extract value after colon excel", "split product options column google sheets", "get size from variation text".
   - Template: tidy-tabs-variation-values-by-label.xlsx (Orders, How to use). Header row 3 with the option names in D3:F3 (yellow inputs: Scent, Size, Gift wrap), rows 4 to 203 (12 fictional orders): A Order no. | B Variations as exported | C Cleaned (helper) | D to F one column per option | G Options not found; summary I3:J6.
     - D4 core: find ","&label&":" in the cleaned text and return what is up to the next comma, with FIND (no wildcards) on LOWER text.
     - Expected: "Size: 12 oz, Scent: Cedar & Sage, Gift wrap: Yes" gives Cedar & Sage / 12 oz / Yes; "Wick size: Double, Scent: Lavender" gives no Size.
   - Test: labels in any order and case, no spaces, spaces around the colon, line breaks, CR LF, semicolons, empty parts, a label that is part of another label, a label twice, empty values, values with a colon (2:30 PM), values with a comma, tabs, a number in B, blank B with and without an order number, header labels in other cases, blank or with extra spaces, 8 random sets against a split-based Python model.
   - Who ranks: Exceljet "extract text after", Ablebits, forum answers using TEXTAFTER/REGEXEXTRACT (newer functions we do not use). Angle: one formula that copies across for any option name and works in older Excel.
