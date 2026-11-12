# Handoff 2026-11-13: five topics researched, hooks written, templates built and tested, heroes made (LibreOffice Calc 24.2.7.2 only), for the script-writers

Stage 1 of the content team (researcher, hook-writer, builder, designer) was done by one agent. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: tasks/team-topics-2026-11-13.md, tasks/hooks-2026-11-13.md, scripts/build-templates.py (`_crit_text`, `possible_duplicate_orders`, `materials_from_orders`, `unix_timestamp_to_date`, `check_amount_in_words`, `variation_values_by_label`), scripts/make-heroes.py (`heroes_2026_11_13`, run with `--only-2026-11-13`), five new files in templates/, five heroes in images/, results in tasks/briefs.md ("2026-11-13 topics (researcher)" and "2026-11-13 builder results"), five titles appended to tasks/topics-seed.md.

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 77 edited workbooks compared (75,831 comparisons, 0 mismatches against independent Python models), plus 6 probe workbooks read by hand for the short formulas and the failure cases below; no error value in any recalculated workbook. All five files exported to PDF: every page US Letter, no error values or ### in the PDF text. Bugs found and fixed are in tasks/briefs.md (the variations cleanup changed text inside values; the Unix file flagged digits-only text that the short formula handles, and let a cut-off timestamp through as 1970; sample timestamps were first a year off).

Repo note: only the five new builder functions were run, so no older template was rewritten; `git status` shows the two scripts, the topics, hooks, handoff, briefs and seed files, five new .xlsx files and five image folders. No older file and no __pycache__ changed.

Post file names: `posts/ready/2026-11-13-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; a red fill marks Check and the other flagged results named below, green marks OK-type results (Enough, 0 options not found), orange marks Buy. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, MINIFS, dynamic arrays. The duplicates post must not tell readers to cancel, refund or dispute anything (a flag is a reason to look). The check post must not say what a bank accepts or that a check is valid; it writes words only. The variations post must not claim a specific marketplace's export format (we did not open any Etsy or Shopify export); say "for example".

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-13-team-01-flag-possible-duplicate-orders-same-customer-day-amount-countifs-excel-google-sheets | tidy-tabs-possible-duplicate-orders.xlsx | 12 workbooks, 12,084 comparisons |
| 2 | 2026-11-13-team-02-materials-shopping-list-from-orders-bill-of-materials-sumifs-roundup-excel-google-sheets | tidy-tabs-materials-from-orders.xlsx | 12 workbooks, 16,872 comparisons |
| 3 | 2026-11-13-team-03-convert-unix-timestamp-to-date-and-time-seconds-milliseconds-excel-google-sheets | tidy-tabs-unix-timestamp-to-date.xlsx | 21 workbooks, 21,147 comparisons |
| 4 | 2026-11-13-team-04-write-check-amount-in-words-dollars-and-cents-index-mod-excel-google-sheets | tidy-tabs-check-amount-in-words.xlsx | 18 workbooks, 14,472 comparisons |
| 5 | 2026-11-13-team-05-pull-size-and-scent-from-variations-column-by-label-find-mid-excel-google-sheets | tidy-tabs-variation-values-by-label.xlsx | 14 workbooks, 11,256 comparisons |

tested_in text pattern (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets."

Hero files (1600x900 PNG, real LibreOffice PDF render of the template with its sample data). Hero alt texts below are 8 to 15 words and differ from the titles:
1. images/2026-11-13-team-01-flag-possible-duplicate-orders-same-customer-day-amount-countifs-excel-google-sheets/possible-duplicate-orders-template.png
   - caption: "Fifteen orders with five possible duplicates in four groups, a repeated order number and a text date. PDF export from LibreOffice Calc 24.2."
   - alt: "Order list with Possible duplicate flags on repeat orders from the same customer and day"
2. images/2026-11-13-team-02-materials-shopping-list-from-orders-bill-of-materials-sumifs-roundup-excel-google-sheets/materials-shopping-list-template.png
   - caption: "Ten candle and soap supplies totaled from the week's orders: six to buy for $307.00. PDF export from LibreOffice Calc 24.2."
   - alt: "Materials tab with amount needed, shortfall, packs to buy and cost for candle supplies"
3. images/2026-11-13-team-03-convert-unix-timestamp-to-date-and-time-seconds-milliseconds-excel-google-sheets/unix-timestamp-to-date-template.png
   - caption: "Twelve payment timestamps in seconds and milliseconds converted at -5 hours from UTC; one is cut off and one is text. PDF export from LibreOffice Calc 24.2."
   - alt: "Ten and thirteen digit Unix timestamps turned into Eastern dates, times and status flags"
4. images/2026-11-13-team-04-write-check-amount-in-words-dollars-and-cents-index-mod-excel-google-sheets/check-amount-in-words-template.png
   - caption: "Twelve checks with the amount written out; fractions of a cent, a negative and a text amount show Check. PDF export from LibreOffice Calc 24.2."
   - alt: "Check register with each dollar amount spelled out in words and cents over 100"
5. images/2026-11-13-team-05-pull-size-and-scent-from-variations-column-by-label-find-mid-excel-google-sheets/variation-values-by-label-template.png
   - caption: "Scent, size and gift wrap pulled from twelve variation texts written in different ways. PDF export from LibreOffice Calc 24.2."
   - alt: "Order list with scent, size and gift wrap pulled from mixed variation text"
The summary block is in the hero for posts 1, 2 and 3; it is hidden in the hero (not in the file) for posts 4 and 5, and so are the Date and helper columns of post 4 (B, E, F) and the Cleaned helper of post 5 (C).

AI photo (`image_prompts`, `image_alt`), one scene each, 2 or 3 specific objects, no screens, no readable text, no logos, no laptop and coffee:
1. Duplicates: image_prompts "Two identical kraft mailer boxes side by side on a home packing table, each with the same folded packing slip tucked under a strip of tape, a roll of packing tape between them, address labels blank and unreadable." image_alt "Two identical kraft mailer boxes side by side on a packing table"
2. Materials: image_prompts "A candle maker's workbench with an open bag of soy wax flakes, a row of empty amber glass jars and a bundle of cotton wicks tied with string, a small kitchen scale at the edge, no labels readable." image_alt "Soy wax flakes, empty amber jars and cotton wicks on a workbench"
3. Unix: image_prompts "A plain analog wall clock above a home office shelf holding a short stack of folded payout statements and a flip desk calendar, all print blurred and unreadable." image_alt "Wall clock over a shelf with folded payout statements and a desk calendar"
4. Check words: image_prompts "An open business checkbook on a wooden desk with a fountain pen resting across a blank check, a small binder-clipped stack of supplier invoices beside it, all handwriting and print unreadable." image_alt "Open business checkbook with a fountain pen resting across a blank check"
5. Variations: image_prompts "Three soy candles in different jar sizes on a packing table, each with a small blank kraft hang tag on twine, next to a spool of gift ribbon and folded tissue paper." image_alt "Three soy candles in different jar sizes beside gift ribbon and tissue paper"

Every workbook has a "How to use" tab and a US Letter page setup with the header row repeated. Titles and openings: tasks/hooks-2026-11-13.md. Deviations from tasks/team-topics-2026-11-13.md: none.

---
## 1. flag-possible-duplicate-orders-same-customer-day-amount-countifs-excel-google-sheets
Title: How to Flag Possible Duplicate Orders by Customer, Date and Amount in Excel and Google Sheets
Template: templates/tidy-tabs-possible-duplicate-orders.xlsx (tabs: Orders, How to use)

Layout
- Orders: A1 note. Header row 3, data rows 4:203 (15 sample orders). A Order no., B Order date and time (mm/dd/yyyy h:mm AM/PM), C Customer, D Amount (inputs); E Order day, F Same customer, day and amount, G Copy no., H Flag, I Times this order no. is listed (formulas). Summary K3:L9: Orders listed (L3), Possible duplicates (2nd copy on) (L4), Groups with more than one order (L5), Amount in possible duplicates (L6), Rows with a repeated order no. (L7), Rows showing Check (L8), Orders with no flag (L9). H red for Check and Possible duplicate; I red when above 1. Landscape.

Key formulas (row 4)
- E4 `=IF(AND(TRIM(A4)="",B4="",TRIM(C4)="",D4=""),"",IF(OR(NOT(ISNUMBER(B4)),TRIM(C4)="",NOT(ISNUMBER(D4))),"Check",IF(OR(B4<1,ROUND(D4,2)<>D4),"Check",INT(B4))))` (core `=INT(B4)`)
- F4 `=IF(ISNUMBER(E4),COUNTIFS($C$4:$C$203,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(C4,"~","~~"),"*","~*"),"?","~?"),$E$4:$E$203,E4,$D$4:$D$203,D4),"")` (core `=COUNTIFS($C$4:$C$203,C4,$E$4:$E$203,E4,$D$4:$D$203,D4)`)
- G4 `=IF(ISNUMBER(E4),COUNTIFS($C$4:C4,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(C4,"~","~~"),"*","~*"),"?","~?"),$E$4:E4,E4,$D$4:D4,D4),"")` (core `=COUNTIFS($C$4:C4,C4,$E$4:E4,E4,$D$4:D4,D4)`; the range grows as it is copied down)
- H4 `=IF(E4="","",IF(E4="Check","Check",IF(F4=1,"",IF(G4=1,"First of "&F4,"Possible duplicate"))))`
- I4 `=IF(OR(E4="",TRIM(A4)=""),"",IF(ISNUMBER(A4),COUNTIF($A$4:$A$203,A4),COUNTIF($A$4:$A$203,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A4,"~","~~"),"*","~*"),"?","~?"))))`
- L3 `=SUMPRODUCT(--(E4:E203<>""))`; L4 `=COUNTIF(H4:H203,"Possible duplicate")`; L5 `=COUNTIF(H4:H203,"First of*")`; L6 `=SUMIFS(D4:D203,H4:H203,"Possible duplicate")`; L7 `=COUNTIF(I4:I203,">1")`; L8 `=COUNTIF(H4:H203,"Check")`; L9 `=L3-L4-L5-L8`
- Why the wrapper: COUNTIFS reads * ? ~ in a criterion as wildcards and a leading < > = as an operator; "="& plus the escapes match the name as plain text.

Sample values (fictional; recalculated)
| Order no. | Order date and time | Customer | Amount | Order day | Same cust/day/amt | Copy no. | Flag | Order no. listed |
|---|---|---|---|---|---|---|---|---|
| #1001 | 11/02/2026 9:14 AM | Harbor Dental | $86.50 | 11/02/2026 | 2 | 1 | First of 2 | 1 |
| #1002 | 11/02/2026 9:15 AM | Harbor Dental | $86.50 | 11/02/2026 | 2 | 2 | Possible duplicate | 1 |
| #1003 | 11/02/2026 1:40 PM | Lakeside Yoga | $42.00 | 11/02/2026 | 1 | 1 | | 1 |
| #1004 | 11/03/2026 10:05 AM | Jordan Reed | $18.00 | 11/03/2026 | 2 | 1 | First of 2 | 1 |
| #1005 | 11/03/2026 4:22 PM | Maria Lopez | $64.25 | 11/03/2026 | 1 | 1 | | 1 |
| #1006 | 11/03/2026 6:10 PM | jordan reed | $18.00 | 11/03/2026 | 2 | 2 | Possible duplicate | 1 |
| #1007 | 11/04/2026 11:00 AM | Cedar Print Shop | $120.00 | 11/04/2026 | 2 | 1 | First of 2 | 2 |
| #1007 | 11/04/2026 11:00 AM | Cedar Print Shop | $120.00 | 11/04/2026 | 2 | 2 | Possible duplicate | 2 |
| #1009 | 11/05/2026 2:15 PM | Pine & Co. Realty | $250.00 | 11/05/2026 | 1 | 1 | | 1 |
| #1010 | 11/05/2026 2:50 PM | Pine & Co. Realty | $25.00 | 11/05/2026 | 1 | 1 | | 1 |
| #1011 | 11/06/2026 9:00 AM | Lopez family | $33.75 | 11/06/2026 | 3 | 1 | First of 3 | 1 |
| #1012 | 11/06/2026 9:02 AM | Lopez family | $33.75 | 11/06/2026 | 3 | 2 | Possible duplicate | 1 |
| #1013 | 11/06/2026 9:05 AM | Lopez family | $33.75 | 11/06/2026 | 3 | 3 | Possible duplicate | 1 |
| #1014 | 11/07/2026 12:30 PM | Harbor Dental | $86.50 | 11/07/2026 | 1 | 1 | | 1 |
| #1015 | 11/08/2026 (text) | Main St. Gift Shop | $54.00 | Check | | | Check | 1 |
Summary L3:L9: 15, 5, 4, $292.00, 2, 1, 5. ($292.00 = 86.50 + 18.00 + 120.00 + 33.75 + 33.75.)

Edge cases tested (LibreOffice 24.2.7.2, all matched the model)
- Names: "A*B Studio" vs "AxB Studio" and "Shop?" vs "Shop1" (not matched), "Tilde~Co" twice, "<Bob>" twice, ">Ann" vs "Ann", "=Sam" vs "Sam" (matched only to themselves), "Harbor Dental " with a trailing space (separate), "HARBOR DENTAL" (same as Harbor Dental), "  Lopez   family " (separate from "Lopez family"), blank or spaces-only customer (Check).
- Dates: minutes apart on one day (match), 11:59 PM vs 12:01 AM the next day (no match), a date with no time, a time with no date (Check), a text date (Check), a negative date (Check), no date (Check).
- Amounts: "$86.50" as text (Check), 12.345 (Check), two $0.00, two -$86.50 refunds (matched), 0.1 vs 0.10 (matched), blank (Check).
- Order numbers: blank (I empty), 1001 as a number twice (2), "ab1" vs "AB1" (2, upper and lower case ignored).
- 200 identical rows (First of 200, 199 possible duplicates), an empty sheet, a row on row 203, 8 random sets of up to 200 rows.

Pitfalls the post must state
- A flag means "look at this", not "this is a duplicate": a customer can really order the same thing twice in a day. Compare order numbers and items before doing anything; the post gives no refund or dispute advice.
- The day comes from the date and time as it is in the export (INT drops the time). If the export is in UTC, two orders late in the evening US time can fall on different UTC days; the file does not convert time zones (link the existing time-zone post if mentioned).
- Upper and lower case do not matter, extra spaces do ("Harbor Dental " is a different customer). Link the existing clean-customer-list (TRIM) post.
- Amounts must match to the cent and be whole cents (12.345 shows Check). Same customer and day with a different amount is not flagged (Pine & Co. Realty $250.00 and $25.00).
- The core COUNTIFS treats * ? ~ in a name as wildcards; the file's wrapper avoids that (see Troubleshooting).
- Not tested: Excel, Google Sheets (the "="& criterion with ~ escapes in COUNTIFS is the open question).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Two different customers counted as one: the core formula with C19 = "A*B Studio" returned 2 because * matched "AxB Studio" (probe); the file returns 1.
- Orders minutes apart are not flagged: counting on column B (date and time) instead of the Order day; the probe COUNTIFS on B returned 1 for #1001 where the file returns 2.
- An obvious duplicate is not flagged: "Harbor Dental " with a trailing space was not flagged next to "Harbor Dental" (probe).
- A row shows Check: the date was text ("11/08/2026" typed as text), or the amount was text or had fractions of a cent, or the customer was blank.
- First of 2 and Possible duplicate on the same order number: the row was pasted twice (#1007; Times this order no. is listed shows 2).

Help links (checked 2026-11-12)
- https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842 (200)
- https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef (200)
- https://support.microsoft.com/en-us/office/using-wildcard-characters-in-searches-ef94362e-9999-4350-ad74-4d2371110adb (200)
- https://support.google.com/docs/answer/3256550 (COUNTIFS; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3093490 (INT; HEAD 404, GET 200)

---
## 2. materials-shopping-list-from-orders-bill-of-materials-sumifs-roundup-excel-google-sheets
Title: How to Make a Materials Shopping List From Your Orders in Excel and Google Sheets
Template: templates/tidy-tabs-materials-from-orders.xlsx (tabs: Materials, Orders, Recipe, How to use)

Layout
- Orders: A1 note. Header row 3, rows 4:203 (10 sample lines). A For (order or event), B Product, C Quantity (inputs); D Status (OK, No recipe or Check; red for No recipe and Check).
- Recipe: A1 note. Header row 3, rows 4:203 (18 sample lines). A Product, B Material, C Amount per item (inputs); D Items ordered, E Material needed, F Status (OK, Not on Materials tab or Check).
- Materials (first tab): A1 note. Header row 4, rows 5:104 (10 sample supplies). A Material, B Unit, C On hand, D Pack size, E Pack price (inputs); F Needed for the orders, G Short, H Packs to buy, I Cost to buy, J Left after buying, K Status (Buy orange, Enough green, Not used, Check red). Summary M4:N9: Materials to buy (N4), Total cost to buy (N5), Materials showing Check (N6), Order lines not counted (N7), Recipe lines with a problem (N8), CHECK (N9). Landscape.

Key formulas
- Orders D4 `=IF(AND(TRIM(A4)="",TRIM(B4)="",C4=""),"",IF(OR(TRIM(B4)="",NOT(ISNUMBER(C4))),"Check",IF(OR(C4<0,C4<>INT(C4)),"Check",IF(COUNTIF(Recipe!$A$4:$A$203,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B4,"~","~~"),"*","~*"),"?","~?"))=0,"No recipe","OK"))))`
- Recipe D4 `=IF(AND(TRIM(A4)="",TRIM(B4)="",C4=""),"",IF(OR(TRIM(A4)="",TRIM(B4)="",NOT(ISNUMBER(C4))),"Check",IF(C4<0,"Check",SUMIFS(Orders!$C$4:$C$203,Orders!$B$4:$B$203,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A4,"~","~~"),"*","~*"),"?","~?"),Orders!$D$4:$D$203,"OK"))))` (core `=SUMIFS(Orders!$C$4:$C$203,Orders!$B$4:$B$203,A4)`)
- Recipe E4 `=IF(ISNUMBER(D4),C4*D4,"")` (core `=C4*D4`)
- Recipe F4 `=IF(D4="","",IF(D4="Check","Check",IF(COUNTIF(Materials!$A$5:$A$104,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B4,"~","~~"),"*","~*"),"?","~?"))=0,"Not on Materials tab","OK")))`
- Materials F5 `=IF(TRIM(A5)="","",SUMIFS(Recipe!$E$4:$E$203,Recipe!$B$4:$B$203,"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A5,"~","~~"),"*","~*"),"?","~?")))` (core `=SUMIFS(Recipe!$E$4:$E$203,Recipe!$B$4:$B$203,A5)`)
- G5 `=IF(F5="","",IF(NOT(ISNUMBER(C5)),"Check",IF(C5<0,"Check",MAX(0,ROUND(F5-C5,6)))))` (core `=MAX(0,F5-C5)`)
- H5 `=IF(G5="","",IF(G5="Check","Check",IF(G5=0,0,IF(NOT(ISNUMBER(D5)),"Check",IF(D5<=0,"Check",ROUNDUP(ROUND(G5/D5,9),0))))))` (core `=ROUNDUP(G5/D5,0)`; the inner ROUND is a precaution, see pitfalls)
- I5 `=IF(NOT(ISNUMBER(H5)),"",IF(H5=0,0,IF(E5="","",IF(NOT(ISNUMBER(E5)),"Check",IF(E5<0,"Check",ROUND(H5*E5,2))))))`
- J5 `=IF(ISNUMBER(H5),ROUND(C5+H5*IF(H5=0,0,D5)-F5,6),"")`
- K5 `=IF(F5="","",IF(OR(G5="Check",H5="Check",I5="Check"),"Check",IF(F5=0,"Not used",IF(H5>0,"Buy","Enough"))))`
- N4 `=COUNTIF(K5:K104,"Buy")`; N5 `=SUM(I5:I104)`; N6 `=COUNTIF(K5:K104,"Check")`; N7 `=COUNTIF(Orders!D4:D203,"Check")+COUNTIF(Orders!D4:D203,"No recipe")`; N8 `=COUNTIF(Recipe!F4:F203,"Check")+COUNTIF(Recipe!F4:F203,"Not on Materials tab")`; N9 `=IF(N6+N7+N8=0,"OK","Check")`

Sample values (fictional; recalculated)
Orders: Main St. Gift Shop wholesale: Lavender soy candle 8 oz 24, Cedar soy candle 8 oz 18, Oatmeal soap bar 30, Holiday gift set 10 (No recipe); Saturday farmers market: Lavender 12, Mini soy candle 4 oz 30, Oatmeal soap bar 20; Etsy orders this week: Cedar 6, Mini 12, Lavender 4. Items ordered: Lavender 40, Cedar 24, Mini 42, Oatmeal soap 50.
Recipe per item: 8 oz candles: soy wax 7 oz, fragrance oil 0.6 oz (lavender or cedar), cotton wick 1, 8 oz amber jar 1, jar label 1. Mini 4 oz: soy wax 3.5 oz, lavender fragrance oil 0.3 oz, cotton wick 1, 4 oz tin 1, jar label 1. Oatmeal soap bar: soap base 4.5 oz, colloidal oatmeal 0.1 oz, soap band 1.
| Material | Unit | On hand | Pack size | Pack price | Needed | Short | Packs | Cost | Left after | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Soy wax | oz | 400 | 160 | $32.00 | 595 | 195 | 2 | $64.00 | 125 | Buy |
| Lavender fragrance oil | oz | 20 | 16 | $28.50 | 36.6 | 16.6 | 2 | $57.00 | 15.4 | Buy |
| Cedar fragrance oil | oz | 12 | 16 | $26.00 | 14.4 | 2.4 | 1 | $26.00 | 13.6 | Buy |
| Cotton wick | each | 150 | 100 | $9.00 | 106 | 0 | 0 | $0.00 | 44 | Enough |
| 8 oz amber jar | each | 36 | 24 | $30.00 | 64 | 28 | 2 | $60.00 | 20 | Buy |
| 4 oz tin | each | 50 | 24 | $14.40 | 42 | 0 | 0 | $0.00 | 8 | Enough |
| Jar label | each | 300 | 250 | $18.00 | 106 | 0 | 0 | $0.00 | 194 | Enough |
| Soap base | oz | 64 | 32 | $15.50 | 225 | 161 | 6 | $93.00 | 31 | Buy |
| Colloidal oatmeal | oz | 8 | 16 | $9.25 | 5 | 0 | 0 | $0.00 | 3 | Enough |
| Soap band | each | 40 | 100 | $7.00 | 50 | 10 | 1 | $7.00 | 90 | Buy |
Summary N4:N9: 6, $307.00, 0, 1, 0, Check. CHECK shows Check on purpose: the Holiday gift set line has no recipe, so its 10 sets are not in the totals. The post should show this and say that adding the gift set's lines to the Recipe tab (or deleting the order line) turns CHECK to OK. Soy wax worked example: 40 x 7 + 24 x 7 + 42 x 3.5 = 280 + 168 + 147 = 595 oz; 595 - 400 = 195 short; 195 / 160 = 1.22, rounded up to 2 packs; 400 + 320 - 595 = 125 left.

Edge cases tested (all matched the exact-fraction model)
- Names: "Candle*" vs "Candle 8" (separate), "candle 8" (same as "Candle 8"), "Candle 8 " with a trailing space (No recipe), "<Odd>", "Wax" and "WAX" both on the Materials tab (both rows show the full need), a material missing from the Materials tab (Not on Materials tab), a recipe line with no product or no material (Check), a product nobody ordered (its lines count 0), a material no product uses (Not used).
- Quantities 2.5, -1, "10" as text and blank (Check), 0 (OK, adds nothing); recipe amounts 0, "1" as text and -1 (Check for text and negative).
- On hand blank, "5" as text, -1 (Check); pack size 0, blank, -5 when short (Check), pack size blank when not short (fine, 0 packs); pack price blank (cost blank), "abc" and -2 (Check); need exactly equal to on hand (0 packs, Enough).
- Float cases: 30 x 0.1 oz against packs of 0.1 (30 packs) and 0.9 against packs of 0.3 (3 packs).
- 200 order lines, no orders at all (every material Not used), 8 random sets of up to 200 orders and recipe lines.

Pitfalls the post must state
- Units must agree: Amount per item uses the unit on the Materials tab (oz with oz, each with each). The file does not convert pounds to ounces; say 10 lb = 160 oz in the example.
- Names must match between tabs, upper and lower case aside; an extra space makes a different name (Orders shows No recipe, Recipe shows Not on Materials tab). Fix these before trusting the totals; CHECK says Check until then.
- Lines marked Check or No recipe are left out of the totals, they are not guessed.
- Waste, spills and test pours are not included; tell readers to raise Amount per item if they always lose some.
- The inner ROUND in Packs to buy guards against a decimal leftover such as 3.0000000001 adding a pack. In LibreOffice the probes ROUNDUP(0.9/0.3,0) and ROUNDUP(0.7/0.1,0) already gave 3 and 7 without it; do not claim the problem was seen. Excel and Sheets were not tested.
- This is a shopping list, not inventory accounting or cost of goods for taxes.
- Not tested: Excel, Google Sheets (SUMIFS across tabs with the "="& criterion is the open question).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- An order is not counted: the product name had a trailing space ("Lavender soy candle 8 oz " showed No recipe in a probe), or the product has no Recipe lines (Holiday gift set in the sample).
- A material shows 0 needed or Not used although products use it: the name on the Recipe tab differs (Recipe shows Not on Materials tab).
- Packs to buy shows Check: pack size is 0, blank or negative while you are short (tested 0, blank, -5); or on hand is blank or text.
- Cost to buy is blank: Pack price is empty (on purpose).
- CHECK shows Check: N7 or N8 is above 0 (sample: 1, the gift set).

Help links (checked 2026-11-12)
- https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b (200)
- https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7 (200)
- https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34 (200)
- https://support.google.com/docs/answer/3238496 (SUMIFS; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3093443 (ROUNDUP; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3093480 (COUNTIF; HEAD 404, GET 200)

---
## 3. convert-unix-timestamp-to-date-and-time-seconds-milliseconds-excel-google-sheets
Title: How to Convert a Unix Timestamp to a Date and Time in Excel and Google Sheets
Template: templates/tidy-tabs-unix-timestamp-to-date.xlsx (tabs: Timestamps, How to use)

Layout
- Timestamps: A1 "Hours from UTC", B1 offset (input, -5); C1 note (Eastern -5/-4, Central -6/-5, Mountain -7/-6, Pacific -8/-7). Header row 3, data rows 4:203 (12 sample payments). A Payment ID, B Timestamp as exported (inputs); C Unit, D Date and time (local, mm/dd/yyyy h:mm:ss AM/PM), E Date, F Time, G Status (OK; Before 2000 and Check in red). Summary I3:J9: Timestamps listed (J3), Converted (OK) (J4), Before 2000 (digits missing?) (J5), Rows showing Check (J6), Earliest OK (local) (J7), Latest OK (local) (J8), Hours from UTC (B1) (J9). Landscape.

Key formulas (row 4)
- C4 `=IF(TRIM(B4)="","",IF(ISNUMBER(B4),IF(B4<0,"Check",IF(B4<100000000000,IF(B4<=4102444800,"Seconds","Check"),IF(B4<=4102444800000,"Milliseconds","Check"))),IF(SUMPRODUCT(LEN(TRIM(B4))-LEN(SUBSTITUTE(TRIM(B4),{"0","1","2","3","4","5","6","7","8","9"},"")))=LEN(TRIM(B4)),IF(--TRIM(B4)<0,"Check",IF(--TRIM(B4)<100000000000,IF(--TRIM(B4)<=4102444800,"Seconds","Check"),IF(--TRIM(B4)<=4102444800000,"Milliseconds","Check"))),"Check")))` (below 100,000,000,000 = seconds; 4102444800 is 01/01/2100 UTC; digits-only text is read as the number)
- D4 `=IF(OR(C4="Seconds",C4="Milliseconds"),IF(ISNUMBER($B$1),IF(AND($B$1>=-12,$B$1<=14),(IF(C4="Seconds",IF(ISNUMBER(B4),B4,--TRIM(B4)),IF(ISNUMBER(B4),B4,--TRIM(B4))/1000)+$B$1*3600)/86400+DATE(1970,1,1),"Check"),"Check"),IF(C4="","","Check"))`
- Cores the post teaches: UTC `=B4/86400+DATE(1970,1,1)`; with the offset `=(B4+$B$1*3600)/86400+DATE(1970,1,1)` (Eastern standard written out: `=(B4-5*3600)/86400+DATE(1970,1,1)`); milliseconds `=B4/86400000+DATE(1970,1,1)`. Format: mm/dd/yyyy h:mm AM/PM (the file shows seconds too).
- E4 `=IF(ISNUMBER(D4),INT(D4),"")`; F4 `=IF(ISNUMBER(D4),D4-INT(D4),"")`; G4 `=IF(C4="","",IF(ISNUMBER(D4),IF(D4<DATE(2000,1,1),"Before 2000","OK"),"Check"))`
- J3 `=SUMPRODUCT(--(G4:G203<>""))`; J4 `=COUNTIF(G4:G203,"OK")`; J5 `=COUNTIF(G4:G203,"Before 2000")`; J6 `=COUNTIF(G4:G203,"Check")`; J7 `=IF(J4=0,"",SMALL(D4:D203,COUNTIF(D4:D203,"<"&DATE(2000,1,1))+1))`; J8 `=IF(J4=0,"",MAX(D4:D203))`; J9 `=IF(ISNUMBER(B1),IF(AND(B1>=-12,B1<=14),"OK","Check"),"Check")`

Sample values (fictional; B1 = -5; recalculated)
| Payment ID | Timestamp | Unit | Date and time (local) | Date | Time | Status |
|---|---|---|---|---|---|---|
| pay_0001 | 1793630100 | Seconds | 11/02/2026 9:35:00 AM | 11/02/2026 | 9:35:00 AM | OK |
| pay_0002 | 1793650500 | Seconds | 11/02/2026 3:15:00 PM | 11/02/2026 | 3:15:00 PM | OK |
| pay_0003 | 1793725200000 | Milliseconds | 11/03/2026 12:00:00 PM | 11/03/2026 | 12:00:00 PM | OK |
| pay_0004 | 1793768399 | Seconds | 11/03/2026 11:59:59 PM | 11/03/2026 | 11:59:59 PM | OK |
| pay_0005 | 1793768400 | Seconds | 11/04/2026 12:00:00 AM | 11/04/2026 | 12:00:00 AM | OK |
| pay_0006 | 1793901045123 | Milliseconds | 11/05/2026 12:50:45 PM (.123 s) | 11/05/2026 | 12:50:45 PM | OK |
| pay_0007 | 1794087000 | Seconds | 11/07/2026 4:30:00 PM | 11/07/2026 | 4:30:00 PM | OK |
| pay_0008 | 1794157200 | Seconds | 11/08/2026 12:00:00 PM | 11/08/2026 | 12:00:00 PM | OK |
| pay_0009 | 1794520800 | Seconds | 11/12/2026 5:00:00 PM | 11/12/2026 | 5:00:00 PM | OK |
| pay_0010 | "1794575700" (text) | Seconds | 11/13/2026 8:15:00 AM | 11/13/2026 | 8:15:00 AM | OK |
| pay_0011 | "2026-11-13T13:15:00Z" (text) | Check | Check | | | Check |
| pay_0012 | 17945208 (cut off) | Seconds | 07/27/1970 11:46:48 AM | 07/27/1970 | 11:46:48 AM | Before 2000 |
Summary J3:J9: 12, 10, 1, 1, 11/02/2026 9:35:00 AM, 11/13/2026 8:15:00 AM, OK. pay_0004 and pay_0005 show the local midnight boundary (one second apart, different dates).
Worked example for the post: 1793630100 / 86400 = 20759.6076...; + DATE(1970,1,1) (serial 25569) = 46328.6076, which is 11/02/2026 2:35 PM UTC; minus 5 hours = 9:35 AM Eastern.

Edge cases tested (all matched the Python datetime model)
- 45 edge values at 12 offsets (-5, -4, 0, 5.5, 14, -12, 0.25, -9.5, and the invalid 14.01, -12.5, "abc", blank): 0 (12/31/1969 7:00 PM at -5), 1, -1 (Check), 86399, 86400, 17999, 18000, 4102444800 (Seconds), 4102444801 and 99999999999 (Check), 100000000000 (Milliseconds, 03/03/1973), 4102444800000 (Milliseconds), 4102444800001 (Check), decimals 1793630100.5 and 1793725200123.4, local midnight and 11:59:59 PM, 946684800 (01/01/2000 UTC), 17945208 (Before 2000), 2^31-1, 2^31, 1E10 (Check), 1E12.
- Text: "1793630100", " 1793630100 ", "0001793630100", "1793725200000" (read as numbers), "1793630100.5", "-1793630100", "1 793 630 100", "1,793,630,100", "99999999999999999", Arabic-Indic digits, ISO text, "abc" (Check), spaces only and blank (empty row).
- 8 random sets of 200 (seconds, milliseconds, digits as text, too-large values, local midnights) at offsets -8 to 9 including -3.5.

Pitfalls the post must state
- B1 is one fixed offset. It does not follow daylight saving time: in 2026 US daylight saving time ended on 11/01/2026, so -5 is right for Eastern from then until it starts again in March; summer dates need -4. For a list that spans the change, convert in UTC (0) or split the rows. Link the existing time-zone post if helpful.
- Seconds vs milliseconds: 10 digits for dates in our era are seconds, 13 digits are milliseconds. The file uses 100,000,000,000 as the line, so a milliseconds value before 03/03/1973 would be read as seconds (not a real-world case for payment exports).
- Using the seconds formula on a 13-digit value gives a date far past the year 9999 (probe: serial about 20,786,277); divide by 86400000 instead.
- A date in 1970 usually means digits were cut off (Before 2000 in the file).
- Text made only of digits works in the file and, in a LibreOffice probe, in the short formula too (pay_0010). ISO date text such as 2026-11-13T13:15:00Z is not a Unix timestamp; it shows Check (the short formula gave an error in the probe). Excel and Sheets behavior with text was not tested.
- Excel and Sheets store dates as days since 12/30/1899 (DATE(1970,1,1) is 25569 in LibreOffice); this was not checked in Excel or Sheets, and Excel's 1904 date system option was not tested.
- Not tested: Excel, Google Sheets.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Hours are off by 5 (or 4): the short formula gives UTC; pay_0001 is 2:35 PM UTC and 9:35 AM at -5 (probe).
- A strange date thousands of years ahead or an error: a 13-digit milliseconds value divided by 86400 (probe on pay_0003: serial about 20,786,277; openpyxl could not read it as a date, LibreOffice's YEAR() returned 32767).
- The date is in 1970: the value was cut to 8 digits (17945208 gave 07/27/1970 11:46:48 AM, Before 2000).
- Check on a row: ISO text, a minus sign, commas or spaces inside the number, or a value past 2100.
- Every row shows Check: B1 is text, blank or outside -12 to 14 (tested "abc", blank, 14.01, -12.5).

Help links (checked 2026-11-12)
- https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349 (200)
- https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309 (200)
- https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef (200)
- https://support.google.com/docs/answer/3092969 (DATE; HEAD 404, GET 200)
- https://support.google.com/docs/answer/56470 (Format numbers in a spreadsheet; HEAD 404, GET 200)

---
## 4. write-check-amount-in-words-dollars-and-cents-index-mod-excel-google-sheets
Title: How to Write a Check Amount in Words in Excel and Google Sheets
Template: templates/tidy-tabs-check-amount-in-words.xlsx (tabs: Checks, Words, How to use)

Layout
- Checks: A1 note (works for $0.00 to $999,999.99). Header row 3, data rows 4:203 (12 sample checks). A Check no., B Date, C Pay to, D Amount (inputs); E Dollars, F Cents, G Amount in words, H Status (formulas, H red for Check). Summary J3:K6: Checks listed (K3), Written out (K4), Rows showing Check (K5), Total of written-out checks (K6). Landscape.
- Words: header row 3; A4:A23 numbers 0 to 19 with B4:B23 words (Zero ... Nineteen, yellow, editable); D4:D13 tens 0 to 90 with E4:E13 words ("", Ten, Twenty ... Ninety, yellow).

Key formulas (row 4)
- E4 `=IF(AND(TRIM(A4)="",B4="",TRIM(C4)="",D4=""),"",IF(NOT(ISNUMBER(D4)),"Check",IF(OR(D4<0,D4>=1000000,ROUND(D4,2)<>D4),"Check",(ROUND(D4*100,0)-MOD(ROUND(D4*100,0),100))/100)))`
- F4 `=IF(ISNUMBER(E4),MOD(ROUND(D4*100,0),100),"")`
- G4 `=IF(NOT(ISNUMBER(E4)),"",IF(E4=0,"Zero",TRIM(IF(INT(E4/1000)>0,TRIM(IF(INT(INT(E4/1000)/100)>0,INDEX(Words!$B$4:$B$23,INT(INT(E4/1000)/100)+1)&" Hundred","")&" "&IF(MOD(INT(E4/1000),100)=0,"",IF(MOD(INT(E4/1000),100)<20,INDEX(Words!$B$4:$B$23,MOD(INT(E4/1000),100)+1),INDEX(Words!$E$4:$E$13,INT(MOD(INT(E4/1000),100)/10)+1)&IF(MOD(INT(E4/1000),10)>0,"-"&INDEX(Words!$B$4:$B$23,MOD(INT(E4/1000),10)+1),""))))&" Thousand","")&" "&TRIM(IF(INT(MOD(E4,1000)/100)>0,INDEX(Words!$B$4:$B$23,INT(MOD(E4,1000)/100)+1)&" Hundred","")&" "&IF(MOD(MOD(E4,1000),100)=0,"",IF(MOD(MOD(E4,1000),100)<20,INDEX(Words!$B$4:$B$23,MOD(MOD(E4,1000),100)+1),INDEX(Words!$E$4:$E$13,INT(MOD(MOD(E4,1000),100)/10)+1)&IF(MOD(MOD(E4,1000),10)>0,"-"&INDEX(Words!$B$4:$B$23,MOD(MOD(E4,1000),10)+1),""))))))&" and "&TEXT(F4,"00")&"/100")`
- How to explain G4 in the post (it is one formula built from one repeated piece): for a number n from 0 to 999, hundreds = INT(n/100) → INDEX(Words, hundreds+1)&" Hundred"; last two digits t = MOD(n,100): 1 to 19 → INDEX(Words, t+1); 20 to 99 → tens word INDEX(Tens, INT(t/10)+1) plus "-" and the ones word when MOD(n,10)>0. G4 applies the piece to INT(E4/1000) (adds " Thousand") and to MOD(E4,1000), then adds " and "&TEXT(F4,"00")&"/100". The post should show the piece on its own and point to the file for the full formula.
- H4 `=IF(E4="","",IF(ISNUMBER(E4),"OK","Check"))`; K3 `=SUMPRODUCT(--(H4:H203<>""))`; K4 `=COUNTIF(H4:H203,"OK")`; K5 `=COUNTIF(H4:H203,"Check")`; K6 `=SUMIFS(D4:D203,H4:H203,"OK")`

Sample values (fictional; recalculated)
| Check no. | Date | Pay to | Amount | Dollars | Cents | Amount in words | Status |
|---|---|---|---|---|---|---|---|
| 1041 | 11/02/2026 | Cedar Print Shop | $1,250.00 | 1,250 | 00 | One Thousand Two Hundred Fifty and 00/100 | OK |
| 1042 | 11/03/2026 | Harbor Packaging Supply | $86.50 | 86 | 50 | Eighty-Six and 50/100 | OK |
| 1043 | 11/05/2026 | Lakeside Studio Rentals | $2,019.07 | 2,019 | 07 | Two Thousand Nineteen and 07/100 | OK |
| 1044 | 11/06/2026 | Pine Street Market | $45.00 | 45 | 00 | Forty-Five and 00/100 | OK |
| 1045 | 11/09/2026 | Maria Lopez | $312.75 | 312 | 75 | Three Hundred Twelve and 75/100 | OK |
| 1046 | 11/10/2026 | Jordan Reed | $0.99 | 0 | 99 | Zero and 99/100 | OK |
| 1047 | 11/12/2026 | Main St. Gift Shop | $15,600.00 | 15,600 | 00 | Fifteen Thousand Six Hundred and 00/100 | OK |
| 1048 | 11/13/2026 | Blue Ridge Wax Supply | $700.10 | 700 | 10 | Seven Hundred and 10/100 | OK |
| 1049 | 11/16/2026 | Sunrise Sign Co. | $4,080.40 | 4,080 | 40 | Four Thousand Eighty and 40/100 | OK |
| 1050 | 11/17/2026 | Harbor Dental | 12.345 (shows $12.35) | Check | | | Check |
| 1051 | 11/18/2026 | Lopez family | -$25.00 | Check | | | Check |
| 1052 | 11/19/2026 | Lakeside Yoga | "$45.00" (text) | Check | | | Check |
Summary K3:K6: 12, 9, 3, $24,094.81.

Edge cases tested (all matched an independent recursive Python speller)
- Every dollar value 0 to 999 (with cents), every thousands group 0 to 999 combined with the rest 999 down to 0; 0.01, 0.1 (10/100), 0.29 (29/100), 0.99, 19, 20, 21, 99, 100, 101, 110, 111, 119, 120, 1,000, 1,001, 1,010, 1,100, 10,000, 10,001, 100,000, 100,100, 101,010.10, 7E5, 999,999.99.
- Check: 1.005, 9.995, 12.345, 1E-9 (fractions of a cent), 999,999.995, 1,000,000, 1,000,000.01, -0.01, -25, "abc", "$45.00" as text, spaces, a payee or a date with no amount; blank rows stay empty.
- 6 random sets of 200 amounts up to $1,000,000.

Pitfalls the post must state
- Fractions of a cent show Check even though the cell displays $12.35 (12.345 in the sample). Round the amount first if it comes from a calculation.
- Style choices are the file's, not rules: capitalized words, a hyphen in 21 to 99 (Forty-Five), "and" only before the cents, cents as 00/100, Zero for amounts under $1.00, and no "Dollars" at the end (add &" Dollars" if wanted; probe: "Two Thousand Nineteen and 07/100 Dollars"). The Words tab can be edited (for example lower case).
- Up to $999,999.99 only.
- Read the words against the number before signing. The post makes no claim about what a bank accepts; it does not print checks or handle MICR lines.
- Not tested: Excel, Google Sheets (INDEX on another tab and TEXT with "00" are the open questions). Microsoft's own approach uses a VBA macro (SpellNumber); the post may say ours uses no macro, but must not say Excel has no other way.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Check on an amount that looks fine: it has a third decimal (12.345 shows as $12.35 but is Check), or it was pasted as text ("$45.00").
- Check on a negative amount (-25.00) or $1,000,000.00 and above.
- Cents look wrong: 7 cents must show as 07/100; TEXT(7,"00") gives 07 (probe).
- 0.29 gives 29/100, not 28/100: ROUND(0.29*100,0) gives 29 (probe).

Help links (checked 2026-11-12)
- https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd (200)
- https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3 (200)
- https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef (200)
- https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c (200)
- https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c (200)
- https://support.google.com/docs/answer/3098242 (INDEX; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3093497 (MOD; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3094139 (TEXT; HEAD 404, GET 200)

---
## 5. pull-size-and-scent-from-variations-column-by-label-find-mid-excel-google-sheets
Title: How to Pull Size and Scent Out of an Order Variations Column in Excel and Google Sheets
Template: templates/tidy-tabs-variation-values-by-label.xlsx (tabs: Orders, How to use)

Layout
- Orders: A1 note. Header row 3: A Order no., B Variations as exported, C Cleaned (helper), D3:F3 option names (yellow inputs: Scent, Size, Gift wrap), G Options not found. Data rows 4:203 (12 sample orders): A, B inputs; C to G formulas. G green when 0, red when above 0 or No variations. Summary I3:J6: Orders listed (J3), Every option found (J4), Some options not found (J5), No variations (J6). Landscape.

Key formulas (row 4)
- C4 `=IF(TRIM(B4)="","",","&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B4,CHAR(10),","),CHAR(13),","),";",",")),", ",",")," ,",",")," :",":")&",")` (line breaks and semicolons become commas, extra spaces go, spaces next to commas and before a colon go, a comma at each end)
- D4 (copied to F4 and down) `=IF(OR($C4="",TRIM(D$3)=""),"",IFERROR(TRIM(MID($C4,FIND(","&LOWER(TRIM(D$3))&":",LOWER($C4))+LEN(TRIM(D$3))+2,FIND(",",$C4,FIND(","&LOWER(TRIM(D$3))&":",LOWER($C4))+1)-FIND(","&LOWER(TRIM(D$3))&":",LOWER($C4))-LEN(TRIM(D$3))-2)),""))`
- Core the post teaches (label in D3, cleaned text in C4): `=IFERROR(TRIM(MID($C4,FIND(","&LOWER(D$3)&":",LOWER($C4))+LEN(D$3)+2,FIND(",",$C4,FIND(","&LOWER(D$3)&":",LOWER($C4))+1)-FIND(","&LOWER(D$3)&":",LOWER($C4))-LEN(D$3)-2)),"")`. Explain the parts: start = position of ",size:" + length of the label + 2 (the comma and the colon); length = position of the next comma minus that start.
- G4 `=IF(C4="",IF(TRIM(A4)="","","No variations"),SUMPRODUCT((TRIM($D$3:$F$3)<>"")*(D4:F4="")))`
- J3 `=SUMPRODUCT(--(G4:G203<>""))`; J4 `=COUNTIF(G4:G203,0)`; J5 `=COUNTIF(G4:G203,">0")`; J6 `=COUNTIF(G4:G203,"No variations")`

Sample values (fictional; recalculated; labels Scent, Size, Gift wrap)
| Order no. | Variations as exported | Scent | Size | Gift wrap | Not found |
|---|---|---|---|---|---|
| #2101 | Scent: Lavender, Size: 8 oz | Lavender | 8 oz | | 1 |
| #2102 | Size: 12 oz, Scent: Cedar & Sage, Gift wrap: Yes | Cedar & Sage | 12 oz | Yes | 0 |
| #2103 | scent:Vanilla bean,size:4 oz,gift wrap:No | Vanilla bean | 4 oz | No | 0 |
| #2104 | Scent: Pumpkin Spice / Size: 8 oz / Gift wrap: Yes (three lines in one cell) | Pumpkin Spice | 8 oz | Yes | 0 |
| #2105 | Wick size: Double, Scent: Lavender | Lavender | | | 2 |
| #2106 | Size : 8 oz , Scent : Cedar & Sage | Cedar & Sage | 8 oz | | 1 |
| #2107 | Gift wrap: No; Scent: Fresh Linen; Size: 12 oz | Fresh Linen | 12 oz | No | 0 |
| #2108 | (blank) | | | | No variations |
| #2109 | Scent: Lavender, Scent: Cedar & Sage, Size: 4 oz | Lavender | 4 oz | | 1 |
| #2110 | Size: 8 oz, Scent: | | 8 oz | | 2 |
| #2111 | Scent: Bay Rum, Size: 8 oz, Pickup time: 2:30 PM | Bay Rum | 8 oz | | 1 |
| #2112 | Color: Gold lid, Size: 8 oz | | 8 oz | | 2 |
Summary J3:J6: 12, 4, 7, 1. Cleaned C4 for #2101: ,Scent: Lavender,Size: 8 oz,

Edge cases tested (all matched a split-on-separators Python model)
- A value with a comma ("Note: Happy birthday, Mom" gives Happy birthday), a value with a colon (Pickup time label gives 2:30 PM), double spaces (collapsed), upper case labels and values (labels matched, values kept as typed), no space after the colon, a bare label or a label with no colon (not found), ":Size: 8" (not found), "Size:: 8" (gives ": 8", the text after the first colon), CR LF, semicolons with a trailing one, ",,Size: 4 oz,,", leading and trailing spaces, Wick size and Sizes vs Size (not matched), a label twice (first wins), & < > % * ~ ? inside values, an empty value, " , , ", a number in B, spaces only, blank B with and without an order number, "Gift wrap :No", a tab before the colon (not matched), a blank line between options, a row on row 203.
- Header labels "size", " Scent ", "GIFT WRAP", "Gift  wrap" (two spaces), "Pickup time", "Wick size", "Note", blank (column stays empty and is not counted in G).
- 8 random sets of 200 with random labels, case, separators (", ", ",", "; ", line break, " , ") and spacing.

Pitfalls the post must state
- A value is cut at the next comma, so a free-text option such as a gift note with a comma is cut (Happy birthday, Mom gives Happy birthday). Commas, semicolons and line breaks are all treated as separators.
- Labels match whole and ignore upper and lower case; Size does not match Wick size. A label typed differently in the export (Sizes, Size (oz)) needs the same spelling in the header.
- If a label appears twice, the first one wins (#2109).
- The cleanup drops a space just before a colon inside a value too (Ratio: 1 :2 gives 1:2); tabs and non-breaking spaces are not cleaned (link the existing remove-non-breaking-spaces post).
- Do not say which marketplaces export which format; the sample strings are made up. Newer functions such as TEXTAFTER or REGEXEXTRACT can do this in some versions; the post may mention they exist but must say the team did not test them and the file does not use them.
- Not tested: Excel, Google Sheets (TRIM on a range inside SUMPRODUCT in G and CHAR(10)/CHAR(13) handling of pasted line breaks are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Size shows Double or a wrong value: a simple formula without the leading comma, `=TRIM(MID(B8,SEARCH("Size:",B8)+5,...))`, returned " Double" from "Wick size: Double, Scent: Lavender" (probe); the file returns nothing for Size there.
- A value is cut short: it contained a comma (probe: "Note: Happy birthday, Mom, Size: 8 oz" gave Happy birthday).
- An option comes back empty although it is in the text: the header label is spelled differently, or the text has a tab before the colon (tested; not matched).
- Options in a different order or on separate lines in one cell still work (#2102, #2104, #2107).

Help links (checked 2026-11-12)
- https://support.microsoft.com/en-us/office/find-function-06213f91-b5be-4544-8b0b-2fd5a775436f (200; the older find-findb URL now says "This article has been retired", do not use it)
- https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028 (200)
- https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332 (200)
- https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9 (200)
- https://support.microsoft.com/en-us/office/lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4 (200)
- https://support.google.com/docs/answer/3094126 (FIND; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3094129 (MID; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3094215 (SUBSTITUTE; HEAD 404, GET 200)
- https://support.google.com/docs/answer/3094083 (LOWER; HEAD 404, GET 200)

Note on Google links: every support.google.com/docs URL above returns 404 to `curl -sIL` (HEAD) in this sandbox and 200 to a plain GET with the matching page title, the same as in earlier runs (tasks/briefs.md). Microsoft URLs return 200 to `curl -sIL`.
