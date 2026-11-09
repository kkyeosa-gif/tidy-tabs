# Handoff 2026-11-10: five topics researched, hooks written, templates built and tested (LibreOffice Calc 24.2.7.2 only), for the script-writers

Stage 1 of the content team (researcher, hook-writer, builder) was done by one agent. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: tasks/team-topics-2026-11-10.md, tasks/hooks-2026-11-10.md, scripts/build-templates.py (`last_price_paid`, `order_quantity_case_pack`, `exact_tag_count`, `cents_to_dollars`, `billing_periods`), five new files in templates/, results in tasks/briefs.md ("2026-11-10 builder results"), titles appended to tasks/topics-seed.md.

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 150 workbooks, 23,720 comparisons, 0 mismatches against independent Python models, no error value in any recalculated workbook. All five files exported to PDF: every page US Letter, no error values in the PDF text. Bugs found and fixed are in tasks/briefs.md (a text date anywhere in the Date column broke the newest-date formula in the last-price file; fixed with an N() helper column).

Repo note: only the five new builder functions were run, so no older template was rewritten; `git status` shows the script, the topics, hooks, handoff, briefs and seed files and five new .xlsx files.

Post file names: `posts/ready/2026-11-10-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays. The billing-period post must not say when anyone must bill; the order-quantity post must not say what to buy.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-10-team-01-look-up-most-recent-price-paid-last-match-lookup-excel-google-sheets | tidy-tabs-last-price-paid.xlsx | 25 workbooks, 960 comparisons |
| 2 | 2026-11-10-team-02-round-order-quantity-up-to-case-size-minimum-order-roundup-excel-google-sheets | tidy-tabs-order-quantity-case-size.xlsx | 39 workbooks, 7,866 comparisons |
| 3 | 2026-11-10-team-03-count-orders-with-exact-tag-comma-separated-cell-find-sumproduct-excel-google-sheets | tidy-tabs-count-exact-tag.xlsx | 26 workbooks, 1,115 comparisons |
| 4 | 2026-11-10-team-04-convert-cents-to-dollars-payments-export-round-divide-100-excel-google-sheets | tidy-tabs-cents-to-dollars.xlsx | 27 workbooks, 8,067 comparisons |
| 5 | 2026-11-10-team-05-group-dates-into-14-day-billing-periods-int-excel-google-sheets | tidy-tabs-billing-periods.xlsx | 33 workbooks, 5,712 comparisons |

Every workbook has a "How to use" tab. Titles and openings: tasks/hooks-2026-11-10.md. Deviations from tasks/team-topics-2026-11-10.md: none in topic or slug; the template cell maps below are the built ones.

---
## 1. look-up-most-recent-price-paid-last-match-lookup-excel-google-sheets
Title: How to Look Up the Most Recent Price Paid for an Item in Excel and Google Sheets
Template: templates/tidy-tabs-last-price-paid.xlsx (tabs: Purchases, Lookup, How to use)

Layout
- Purchases: note A1, header row 3, data rows 4:203 (13 sample purchases). A Date, B Item, C Unit cost (inputs), D Item (trimmed) `=IF(TRIM(B4)="","",TRIM(B4))`, E Date as number `=N(A4)` (shown blank for 0).
- Lookup: header row 3, items typed in A4:A23 (6 sample items). B Purchases found, C Last price entered, D On Purchases row, E Date on that row, F Newest date for this item, G Status. Summary I3:J9: items typed, with a price, no purchase, newer date higher up, purchase rows with an item but no numeric cost, purchase rows with a date that is not a real date, CHECK.

Key formulas (Lookup row 4; `match` = `(Purchases!$D$4:$D$203=TRIM(A4))*ISNUMBER(Purchases!$C$4:$C$203)`)
- B4 `=IF(TRIM(A4)="","",SUMPRODUCT(match))`
- C4 `=IF(B4="","",IF(B4=0,"No purchase",LOOKUP(2,1/(match),Purchases!$C$4:$C$203)))` (the post teaches the core `=LOOKUP(2,1/((range=item)*ISNUMBER(cost)),cost)`)
- D4 same LOOKUP returning `ROW(Purchases!$C$4:$C$203)`; E4 `=IF(ISNUMBER(D4),IF(ISNUMBER(INDEX(Purchases!$A$4:$A$203,D4-3)),INDEX(Purchases!$A$4:$A$203,D4-3),"No date"),"")`
- F4 `=IF(ISNUMBER(D4),SUMPRODUCT(MAX(match*Purchases!$E$4:$E$203)),"")`
- G4 No purchase / No date on that row / Newer date higher up (F4>E4) / OK
- J9 `=IF(AND(J3=J4+J5,J7=0,J8=0),"OK","Check")`

Sample values (fictional)
| Item typed | Found | Last price | Row | Date on row | Newest date | Status |
|---|---|---|---|---|---|---|
| Soy wax, 10 lb | 3 | $61.00 | 12 | 11/03/2026 | 11/03/2026 | OK |
| Kraft mailer 6x6x4 | 3 | $0.71 | 14 | 11/09/2026 | 11/09/2026 | OK |
| "  tissue paper, 500 sheets " | 3 | $18.90 | 15 | 10/15/2026 | 10/27/2026 | Newer date higher up |
| Lavender oil, 4 oz | 2 | $15.10 | 13 | 11/05/2026 | 11/05/2026 | OK |
| 8 oz tin | 2 | $1.20 | 16 | 11/07/2026 | 11/07/2026 | OK |
| Wax melt molds | 0 | No purchase | | | | No purchase |
Summary J3:J9: 6, 5, 1, 1, 0, 0, OK.

Pitfalls the post must state
- LOOKUP(2,1/...) returns the LAST matching row of the list, not the newest date. The tissue paper row shows it: the last row is dated 10/15/2026 but 10/27/2026 is newer.
- Rows with a blank or text cost are skipped; a cost of 0 counts as a price.
- Names must match after TRIM and ignoring capitals; "Kraft mailer" does not match "Kraft mailer 6x6x4".
- A text date is treated as no date in the newest-date check and counted in the CHECK block. In LibreOffice a date-like text such as "11/05/2026" was read as a date inside arithmetic; do not promise the same elsewhere.
- Not tested: Excel, Google Sheets (the array behavior of LOOKUP(2,1/...) and SUMPRODUCT(MAX(...)) there). Link Microsoft's LOOKUP help page; do not claim Sheets behavior.

---
## 2. round-order-quantity-up-to-case-size-minimum-order-roundup-excel-google-sheets
Title: How to Round an Order Quantity Up to a Case Size and Minimum Order in Excel and Google Sheets
Template: templates/tidy-tabs-order-quantity-case-size.xlsx (tabs: Order, How to use)

Layout
- Order: note A1, header row 3, data rows 4:43 (8 sample items). A Item, B Need (units), C Case size, D Minimum order (units), H Unit cost (inputs); E Cases, F Units to order, G Extra units, I Order cost, J Row check (formulas). Summary L3:M9: rows with an order, units needed, units to order, extra units, order cost, rows showing Check, CHECK.

Key formulas (row 4)
- J4 `OK` or `Check` (need and minimum whole numbers 0 or more, case size a whole number 1 or more, unit cost blank or a number; all-empty rows stay blank)
- E4 `=IF(J4="OK",IF(B4=0,0,ROUNDUP(MAX(B4,D4)/C4,0)),"")` (the post teaches `=ROUNDUP(MAX(B4,D4)/C4,0)`)
- F4 `=IF(J4="OK",E4*C4,"")`; G4 `=IF(J4="OK",F4-B4,"")`; I4 `=IF(AND(J4="OK",ISNUMBER(H4)),ROUND(F4*H4,2),"")`
- M9 `=IF(AND(M8=0,M5=M4+M6),"OK","Check")`

Sample values (fictional)
| Item | Need | Case | Minimum | Cases | Units | Extra | Unit cost | Cost |
|---|---|---|---|---|---|---|---|---|
| Kraft mailer 6x6x4 | 130 | 25 | 100 | 6 | 150 | 20 | $0.62 | $93.00 |
| Tissue paper packs | 3 | 2 | 0 | 2 | 4 | 1 | $18.50 | $74.00 |
| 8 oz candle tin | 340 | 24 | 240 | 15 | 360 | 20 | $1.15 | $414.00 |
| Soy wax 10 lb box | 0 | 1 | 0 | 0 | 0 | 0 | $54.00 | $0.00 |
| Lavender oil 4 oz | 5 | 6 | 12 | 2 | 12 | 7 | $14.25 | $171.00 |
| Thank-you cards | 40 | 50 | 0 | 1 | 50 | 10 | $0.30 | $15.00 |
| Shipping label roll (500) | 900 | 500 | 1000 | 2 | 1000 | 100 | $0.03 | $30.00 |
| Glass jar 8 oz | 77 | 12 | 0 | 7 | 84 | 7 | $1.40 | $117.60 |
Summary M3:M9: 8, 1,495, 1,660, 165, $914.60, 0, OK.

Pitfalls the post must state
- A need of 0 orders nothing even when the supplier has a minimum.
- A minimum that is not a multiple of the case is rounded up to whole cases (minimum 100, case 24 gives 120).
- Decimals, negatives, text, case size 0 or 1.5 show Check and the row is left out of the totals.
- The case size and minimum come from the supplier; the sample values are made up. Not tested: Excel, Google Sheets (ROUNDUP and MAX are documented by Microsoft; link the help page).

---
## 3. count-orders-with-exact-tag-comma-separated-cell-find-sumproduct-excel-google-sheets
Title: How to Count Orders With an Exact Tag in a Comma Separated Cell in Excel and Google Sheets
Template: templates/tidy-tabs-count-exact-tag.xlsx (tabs: Orders, How to use)

Layout
- Orders: note A1, header row 3, data rows 4:203 (12 sample orders). A Order, B Tags (comma separated), C Amount (inputs), D Tags cleaned (formula). Tag table F3:J13: F Tag to count (inputs F4:F13, 6 sample tags), G Orders with this tag, H Amount of those orders, I COUNTIF with * (partial match), J Difference. Summary L3:M7: orders listed, orders with no tags, amounts that are not numbers, tags to count that contain a comma, CHECK.

Key formulas (row 4)
- D4 `=IF(TRIM(B4)="","",","&SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B4)),", ",",")," ,",",")&",")`
- G4 `=IF(TRIM(F4)="","",IF(ISNUMBER(FIND(",",F4)),"Check",SUMPRODUCT(--ISNUMBER(FIND(","&LOWER(TRIM(F4))&",",$D$4:$D$203)))))`
- H4 `=IF(ISNUMBER(G4),SUMPRODUCT(--ISNUMBER(FIND(","&LOWER(TRIM(F4))&",",$D$4:$D$203)),$C$4:$C$203),"")`
- I4 `=IF(ISNUMBER(G4),COUNTIF($B$4:$B$203,"*"&TRIM(F4)&"*"),"")`; J4 `=IF(ISNUMBER(G4),I4-G4,"")`
- M7 `=IF(AND(M5=0,M6=0),"OK","Check")`

Sample values (fictional orders 1001 to 1012; tags such as "gift, rush", "Gift Wrap", "gifted, repeat", "gift , personalized", "  GIFT  ", "gift, gift")
| Tag | Orders (exact) | Amount | COUNTIF * | Difference |
|---|---|---|---|---|
| gift | 4 | $109.50 | 7 | 3 |
| gift wrap | 2 | $80.00 | 2 | 0 |
| rush | 3 | $122.00 | 3 | 0 |
| repeat | 3 | $274.50 | 3 | 0 |
| wholesale | 2 | $390.00 | 2 | 0 |
| personalized | 2 | $73.00 | 2 | 0 |
Summary M3:M7: 12, 1, 0, 0, OK.

Pitfalls the post must state
- Exact means the whole tag between commas: gift does not match gifted, a gift or gift wrap. An order that lists a tag twice counts once.
- Capitals and spaces around commas are ignored; spaces INSIDE a tag are kept (gift wrap is one tag). Tabs and non-breaking spaces were not tested.
- A tag to count cannot contain a comma (Check). * ? ~ in a tag are searched as plain characters in G and H but act as wildcards in the COUNTIF column; the COUNTIF column was not compared for those.
- Text amounts are skipped in H and counted in the CHECK block.
- Not tested: Excel, Google Sheets (FIND, SUMPRODUCT and COUNTIF wildcards are documented by Microsoft and Google; link them but do not claim a test).

---
## 4. convert-cents-to-dollars-payments-export-round-divide-100-excel-google-sheets
Title: How to Convert Cents to Dollars in a Payments Export in Excel and Google Sheets
Template: templates/tidy-tabs-cents-to-dollars.xlsx (tabs: Payments, How to use)

Layout
- Payments: A1 label, B1 payout your processor reported (input, 1160.94, fictional), note C1, header row 3, data rows 4:203 (8 sample payments). A Date, B Order, C Amount (cents), D Fee (cents) (inputs); E Amount ($), F Fee ($), G Net ($), H Row check. Dollars to cents block: J4:J13 typed dollars (4 samples), K Cents. Summary M3:N10: payments converted, amount total, fee total, net total, payout in B1, net total minus payout, rows showing Check, CHECK.

Key formulas (row 4)
- H4 `OK` when C4 is a whole number and D4 is blank or a whole number, else `Check` (an all-empty row stays blank; a number stored as text is Check)
- E4 `=IF(H4="OK",ROUND(C4/100,2),"")`; F4 `=IF(H4="OK",IF(D4="",0,ROUND(D4/100,2)),"")`; G4 `=IF(H4="OK",ROUND(E4-F4,2),"")`
- K4 `=IF(J4="","",IF(ISNUMBER(J4),ROUND(J4*100,0),"Check"))`
- N8 `=IF(ISNUMBER(B1),ROUND(N6-B1,2),"")`; N10 `=IF(N9<>0,"Check",IF(ISNUMBER(B1),IF(N8=0,"OK","Check"),"No payout in B1"))`

Sample values (fictional; fees are made up and say nothing about any processor)
| Cents | Fee cents | Amount | Fee | Net |
|---|---|---|---|---|
| 1850 | 84 | $18.50 | $0.84 | $17.66 |
| 12500 | 405 | $125.00 | $4.05 | $120.95 |
| 5 | blank | $0.05 | $0.00 | $0.05 |
| -2400 | blank | -$24.00 | $0.00 | -$24.00 |
| 4999 | 175 | $49.99 | $1.75 | $48.24 |
| 100000 | 3120 | $1,000.00 | $31.20 | $968.80 |
| 2175 | 94 | $21.75 | $0.94 | $20.81 |
| 899 | 56 | $8.99 | $0.56 | $8.43 |
Summary N3:N10: 8, $1,200.28, $39.34, $1,160.94, $1,160.94, $0.00, 0, OK. Dollars to cents: 19.99 gives 1999, 0.07 gives 7, 1250 gives 125000, 4.35 gives 435.

Pitfalls the post must state
- Amounts with a fraction of a cent (1850.5), text, or a number stored as text show Check and are left out of the totals; convert text to numbers first (link the existing convert-text-to-numbers post).
- A blank fee counts as 0. A refund stays negative.
- Going back, always ROUND: 19.99*100 is 1998.9999999999998 before rounding. Dollars typed with 3 decimals are not handled exactly (1.005 gave 100 cents in LibreOffice because of floating point); type at most 2 decimals.
- The payout check only compares the net total with the number you type; it does not say the export is right.
- Not tested: Excel, Google Sheets.

---
## 5. group-dates-into-14-day-billing-periods-int-excel-google-sheets
Title: How to Group Dates Into 14-Day Billing Periods in Excel and Google Sheets
Template: templates/tidy-tabs-billing-periods.xlsx (tabs: Entries, How to use)

Layout
- Entries: B1 first period starts (input, 11/02/2026), B2 days in each period (input, 14), note C1, header row 4, data rows 5:204 (13 sample entries). A Date, B What, C Amount (inputs); D Period no., E Period starts, F Period ends (formulas). Period table H4:L16: H Period no. (1 to 12), I Starts, J Ends, K Entries, L Amount. Summary rows: H17 Periods 1 to 12 (K17, L17), H18 Before the first period (K18, L18), H19 After period 12 (K19, L19), H20 Rows showing Check (K20), H21 All amounts (L21), H22 CHECK (L22).

Key formulas (row 5)
- D5 `=IF(A5="","",IF(NOT(ISNUMBER(A5)),"Check",IF(NOT(ISNUMBER($B$1)),"Check",IF($B$1<>INT($B$1),"Check",IF(NOT(ISNUMBER($B$2)),"Check",IF(OR($B$2<1,$B$2<>INT($B$2)),"Check",IF(A5<$B$1,"Before start",INT((A5-$B$1)/$B$2)+1)))))))` (the post teaches the core `=INT((A5-$B$1)/$B$2)+1`)
- E5 `=IF(ISNUMBER(D5),$B$1+(D5-1)*$B$2,"")`; F5 `=IF(ISNUMBER(D5),E5+$B$2-1,"")`
- K5 `=COUNTIF($D$5:$D$204,H5)`; L5 `=SUMIFS($C$5:$C$204,$D$5:$D$204,H5)`
- L22 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),K20=0,ROUND(L17+L18+L19-L21,2)=0),"OK","Check")`

Sample values (fictional freelance entries)
| Period | Starts | Ends | Entries | Amount |
|---|---|---|---|---|
| 1 | 11/02/2026 | 11/15/2026 | 4 | $655.00 |
| 2 | 11/16/2026 | 11/29/2026 | 3 | $435.00 |
| 3 | 11/30/2026 | 12/13/2026 | 3 | $585.00 |
| 4 | 12/14/2026 | 12/27/2026 | 2 | $145.00 |
Periods 5 to 12 are empty (starts 12/28/2026, 01/11/2027, ... 04/05/2027). Before the first period: 1 entry, $120.00 (10/30/2026 Logo revision). After period 12: 0. All amounts $1,940.00, CHECK OK. 11/15/2026 is period 1, 11/16/2026 is period 2, 12/20/2026 is period 4.

Pitfalls the post must state
- Period 1 runs from B1 through B1 + B2 - 1, so the last day of period 1 is 11/15/2026 for an 11/02/2026 start and 14 days.
- Dates before B1 show Before start; the table covers 12 periods and later dates go to After period 12 (extend the table by copying the last row down and adjusting the three total formulas).
- A date typed as text (even "11/03/2026"), a start date with a time, a period length that is 0, negative, fractional, text or blank, show Check. A date with a time of day counts by its date.
- This only groups dates; it does not say when to bill. Not tested: Excel, Google Sheets (INT, COUNTIF, SUMIFS are documented by Microsoft and Google; link them but do not claim a test).
