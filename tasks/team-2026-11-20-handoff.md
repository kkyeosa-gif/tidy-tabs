# Handoff 2026-11-20: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-20.md (researcher) and tasks/hooks-2026-11-20.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`subscription_churn_rate`, `discount_break_even_sales`, `ordinal_dates` with the helper `_ordinal_suffix`, `recipe_unit_converter`, `project_time_elapsed_tracker`, registered after `average_without_highest_lowest()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-20 build and test"). Heroes are not made yet (designer, see the hero section below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 182 edited workbooks recalculated and compared (87,235 cell comparisons, 0 mismatches against independent Python models; a mutation run confirmed the harness catches wrong results), plus 1 probe workbook of the traps read by hand; no error value in any recalculated template workbook. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text. No formula bug turned up; the PDF renders showed 5 cut-off labels, now fixed (tasks/briefs.md).

Repo note: only the five new builder functions were run against templates/ (imported and run with /usr/bin/python3 -B), so no older template was rewritten and __pycache__ was not touched. The whole script was also run from a scratch directory, /tmp/claude-0/build-1120/full (exit 0, 226 files written there, including the five new ones), so a full re-run still works. `git status` shows scripts/build-templates.py, the five new .xlsx files, this handoff and tasks/briefs.md.

Post file names: `posts/ready/2026-11-20-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks Check (every status starting "Check"), a CHECK cell that is not OK, End = next start? Check (churn), Check units (converter) and Overdue (projects); orange marks Discount wipes out profit and the Table's No (discount) and Behind (projects); green marks Done and Ahead (projects). The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, MINIFS/MAXIFS, dynamic arrays (none used; COUNTIFS, AVERAGEIFS, CONVERT, CHOOSE and IFERROR are plain, no _xlfn in any file). Help links were not checked by the builder; the script-writer checks every link as usual. Hooks: every number in the chosen openings of tasks/hooks-2026-11-20.md matches the built files; opening 5 says "(Check dates)", but the template's label is **Check: due before start** (the hooks file was not edited; the script-writer fixes the wording).

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-20-team-01-subscription-box-churn-rate-by-month-active-subscribers-countifs-excel-google-sheets | tidy-tabs-subscription-churn-rate.xlsx | 37 workbooks, 27,898 comparisons |
| 2 | 2026-11-20-team-02-how-many-more-sales-to-make-up-for-a-discount-break-even-volume-excel-google-sheets | tidy-tabs-discount-break-even-sales.xlsx | 38 workbooks, 20,330 comparisons |
| 3 | 2026-11-20-team-03-add-st-nd-rd-th-to-dates-and-numbers-ordinal-suffix-mod-choose-excel-google-sheets | tidy-tabs-ordinal-dates.xlsx | 36 workbooks, 21,888 comparisons |
| 4 | 2026-11-20-team-04-convert-cups-tablespoons-teaspoons-recipe-units-convert-function-excel-google-sheets | tidy-tabs-recipe-unit-converter.xlsx | 30 workbooks, 8,550 comparisons |
| 5 | 2026-11-20-team-05-percent-of-time-elapsed-vs-percent-complete-project-status-excel-google-sheets | tidy-tabs-project-time-elapsed-tracker.xlsx | 41 workbooks, 8,569 comparisons |

tested_in text pattern (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets."

Hero files: for the designer (not made by the builder). Key columns for a hero (real LibreOffice PDF render of the template with its sample data, 1600x900):
1. Churn: Monthly tab rows 4 to 10 (June to November 2026), A Month start, B Active at start, C New, D Cancelled, E Cancelled from the start group, F Active at end, G Churn rate, H All cancels / active at start (July 25.00% beside 50.00% is the point), plus B1 First month. Hide rows 11 to 28 (December 2026 on: no sample events, every month 7 / 0 / 0.00%) and the summary L:M; I End = next start? and J Blank-only start count are optional (J shows the blank-only trap, 2 instead of 4 in July). Subscribers A4:E18 (14 rows, Active / Cancelled) is an optional inset.
2. Discount: Products tab A Product to J More units than now plus M Status, rows 4 to 12 (8 rows; the 25% row 171.43 / 172 / 71.43% and the two orange Discount wipes out profit rows). K, L and the summary O4:P8 can be hidden. The Table tab A4:K17 (the 60% row: 9.09% to 500.00%, No cells orange) works as a second image or inset.
3. Ordinal: Dates tab A Date to E Sign line, rows 4 to 16 (12 dates, 1st / 2nd / 3rd / 11th to 13th / 21st to 23rd / 31st), B1 "Pickup on" visible. F Status and H4:I7 can be hidden. Numbers A4:B18 (1st to 113th) is an optional inset.
4. Converter: Converter tab A Ingredient, B Amount, C From unit, D To unit, G Result, rows 4 to 10 (6 rows); hide E and F (code helpers), H Status and J4:K9. The Units tab grid E4:L11 (1 cup = 16 tbs = 48 tsp = 8 fl oz) is a good inset.
5. Projects: Projects tab B1:B2 (as-of 11/20/2026, tolerance 10%), A Project to I Status, rows 4 to 12 (Behind orange, Overdue red, Done green). The summary K4:L12 can be hidden: L12 shows Check on purpose because of the Brochure row (dates typed backwards, from the topic); if the hero should not show red, end at row 11.

Every workbook has a "How to use" tab and a US Letter page setup with the header row (row 4) repeated (data tabs landscape, fit to one page wide). Deviations from tasks/team-topics-2026-11-20.md are listed per section.

---
## 1. subscription-box-churn-rate-by-month-active-subscribers-countifs-excel-google-sheets
Title (from topics): How to Calculate Monthly Churn Rate for a Subscription Box in Excel and Google Sheets
Template: templates/tidy-tabs-subscription-churn-rate.xlsx (tabs: Subscribers, Monthly, How to use)

Layout
- Subscribers: A1:A2 notes (cancel date = the day it stopped; clear with Delete, a space is not empty; a pause and restart is a new row). Header row 4, rows 5:504 (14 sample rows): A Subscriber or ID, B Plan, C Start date, D Cancel date (blank while active) (inputs, MM/DD/YYYY); E Status (Active, Cancelled, Check: start date is text, Check: cancel date is text, Check: no start date, Check: cancel, no start, Check: cancel before start). Summary G4:H8: Subscribers listed (H4), Active (no cancel date) (H5), Cancelled (H6), Rows showing Check (H7), Typed below row 504 (H8). Landscape.
- Monthly: A1 First month, B1 06/01/2026 (input); A2 Average churn through, B2 10/01/2026 (input; blank = the month of the latest date on Subscribers). Header row 4, rows 5:28 (24 months): A Month start, B Active at start, C New in month, D Cancelled in month, E Cancelled from the start group, F Active at end, G Churn rate (start group), H All cancels / active at start (comparison), I End = next start?, J Blank-only start count (wrong, comparison). Summary L4:M12: Average churn through month (M4), Months in the average (M5), Average monthly churn rate (M6), New subscribers (M7), Cancellations (M8), Months where end is not next start (M9), Subscribers tab rows showing Check (M10), Typed below row 504 (M11), CHECK (M12). Landscape.
- Deviations: B1 is normalized to the 1st of its month (topic A5 `=$B$1`); B2 and M4 added so the average stops at a chosen month (the 24 rows run into future months that show 0.00% and would pull a plain average down); J blank-only comparison column added; I28 recomputes the next month's start inline (no row 29); every Monthly cell is blank when B1 is blank or text; Status words as above.

Key formulas (row 5; S = Subscribers!$C$5:$C$504, C = Subscribers!$D$5:$D$504)
- Subscribers E5 `=IF(AND(TRIM(A5)="",TRIM(B5)="",C5="",D5=""),"",IF(AND(C5<>"",NOT(ISNUMBER(C5))),"Check: start date is text",IF(AND(D5<>"",NOT(ISNUMBER(D5))),"Check: cancel date is text",IF(C5="",IF(D5="","Check: no start date","Check: cancel, no start"),IF(D5="","Active",IF(D5<C5,"Check: cancel before start","Cancelled"))))))`
- A5 `=IF(ISNUMBER($B$1),DATE(YEAR($B$1),MONTH($B$1),1),"")`; A6 `=IF(A5="","",DATE(YEAR(A5),MONTH(A5)+1,1))`
- B5 `=IF(A5="","",COUNTIFS(S,"<"&A5,C,"")+COUNTIFS(S,"<"&A5,C,">="&A5))` (core `=COUNTIFS(start,"<"&m,cancel,"")+COUNTIFS(start,"<"&m,cancel,">="&m)`)
- C5 `=IF(A5="","",COUNTIFS(S,">="&A5,S,"<"&A6))`; D5 `=IF(A5="","",COUNTIFS(C,">="&A5,C,"<"&A6))`; E5 `=IF(A5="","",COUNTIFS(S,"<"&A5,C,">="&A5,C,"<"&A6))` (row 28 uses `DATE(YEAR(A28),MONTH(A28)+1,1)` in place of A29)
- F5 `=IF(A5="","",B5+C5-D5)`; G5 `=IF(N(B5)=0,"",E5/B5)`; H5 `=IF(N(B5)=0,"",D5/B5)`; I5 `=IF(A5="","",IF(F5=B6,"OK","Check"))`; J5 `=IF(A5="","",COUNTIFS(S,"<"&A5,C,""))`
- M4 `=IF(ISNUMBER($B$2),DATE(YEAR($B$2),MONTH($B$2),1),IF(COUNT(Subscribers!$C$5:$D$504)=0,"",DATE(YEAR(MAX(Subscribers!$C$5:$D$504)),MONTH(MAX(Subscribers!$C$5:$D$504)),1)))`; M5 `=IF(M4="",0,COUNTIFS(A5:A28,"<="&M4,B5:B28,">0"))`; M6 `=IF(M5=0,"",AVERAGEIFS(G5:G28,A5:A28,"<="&M4,B5:B28,">0"))`; M7 `=IF(M4="","",SUMIF(A5:A28,"<="&M4,C5:C28))` (M8 on D); M9 `=COUNTIF(I5:I28,"Check")`; M10 `=Subscribers!H7`; M11 `=Subscribers!H8`; M12 `=IF(AND(ISNUMBER(A5),Subscribers!H4>0,M9=0,M10=0,M11=0),"OK","Check")`

Sample values (fictional; 14 subscribers Ava to Nora, starts 06/01 to 10/14/2026; recalculated)
| Month | Active at start | New | Cancelled | From start group | Active at end | Churn rate | All cancels / start | Blank-only count |
|---|---|---|---|---|---|---|---|---|
| 06/01/2026 | 0 | 4 | 0 | 0 | 4 | (blank) | (blank) | 0 |
| 07/01/2026 | 4 | 3 | 2 | 1 | 5 | 25.00% | 50.00% | 2 |
| 08/01/2026 | 5 | 3 | 1 | 1 | 7 | 20.00% | 20.00% | 3 |
| 09/01/2026 | 7 | 2 | 2 | 1 | 7 | 14.29% | 28.57% | 5 |
| 10/01/2026 | 7 | 2 | 1 | 1 | 8 | 14.29% | 14.29% | 6 |
| 11/01/2026 | 8 | 0 | 1 | 1 | 7 | 12.50% | 12.50% | 7 |
December 2026 to May 2028: 7, 0, 0, 0, 7, 0.00%. Every End = next start? OK. M4:M12: 10/01/2026, 4, 18.39% (0.183928571428571), 14, 6, 0, 0, 0, OK. With B2 blank: 11/01/2026, 5 months, 17.21%, 14 new, 7 cancellations. Subscribers H4:H8: 14, 7, 7, 0, 0. Quotable: Nora's cancel date 07/01 counts her as active on July 1 and cancelled in July; July's 25% vs 50% comes from Dev (joined 07/01, left 07/25), who was never in the start count.

Edge cases tested (all matched a Python model)
- Cancel date with no start date (the topic's unverified row, 09/10/2026 added to the sample): the template's start counts stay 0, 4, 5, 7, 7, 8; the short form `=COUNTIF(start,"<"&m)-COUNTIF(cancel,"<"&m)` gave 0, 4, 5, 7, 6, 7 (one low from October on). In the template, September Cancelled in month is 3, end 6 vs October start 7, I Check; Status Check: cancel, no start; CHECK Check. On the clean sample the short form gave the same starts as the template.
- Cancel typed as text "11/02/2026" and a space in the cancel cell (both on a 07/02 start): COUNTIFS skips the person after July (July new 4, end 6, August start 5: I Check), Status Check: cancel date is text, CHECK Check. Cancel before start (09/15 start, 08/20 cancel): August and September I Check. Start typed as text and name only: not counted anywhere, I stays OK, only Status and CHECK show it.
- Start and cancel both 07/01: July start 4, new 4, cancelled 3, from start group 1, end 5, churn 25.00%, all-cancels 75.00%. Cancel on 10/01 (Aug start): active at the start of October and cancelled in October (churn 25.00%). Pause and restart as two rows: both counted, every I OK. Date-times (start 08/01 3:30 PM, cancel 09/01 9:00 AM): August new, September cancelled.
- B1 blank or text (every month blank, CHECK Check), B1 06/15/2026 (same as 06/01), B1 07/01/2026, B2 text (uses the latest month), B2 before B1 (0 months, average blank, CHECK OK), no subscribers (CHECK Check), a row on 505 (CHECK Check), 500 random subscribers, cancels inside row 28's month, 10 random sets of 20 to 300 subscribers and 3 random sets with bad rows.

Pitfalls the post must state
- Churn has several definitions; this sheet uses cancels from the start group / active at start and shows all cancels / active at start as a comparison only. No claim about which one tools or investors use, no benchmark, no revenue (MRR).
- The cancel date is the day it stopped; clear a cancel cell with Delete (a space shows Check); a restart is a new row.
- Months after the last real data show 0.00%: set B2 to the last full month before averaging.
- Not tested: Excel, Google Sheets (COUNTIFS "" against an empty cell and a space, AVERAGEIFS skipping "" results are the open questions). Billing platforms and exports are not named.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (template copies and probe)
- Too few active subscribers: only blank cancel dates were counted (J: July 2 instead of 4).
- Churn looks too high in a month with sign-ups: all of the month's cancels were divided by the start count (July 50% instead of 25%).
- End = next start? shows Check: a cancel date typed as text, a space in the cancel cell, a cancel with no start date, or a cancel before the start.
- The average churn is lower than any real month: B2 is blank or later than your data, so 0.00% future months are included (B2 blank with sample: 17.21% through November; with 10/01/2026: 18.39%).

---
## 2. how-many-more-sales-to-make-up-for-a-discount-break-even-volume-excel-google-sheets
Title (from topics): How to Calculate How Many More Sales You Need to Make Up for a Discount in Excel and Google Sheets
Template: templates/tidy-tabs-discount-break-even-sales.xlsx (tabs: Products, Table, How to use)

Layout
- Products: A1:A2 notes. Header row 4, rows 5:54 (8 sample rows): A Product, B Regular price, C Unit cost, D Discount %, E Units at regular price (inputs); F Margin, G Profit per unit on sale, H Units needed on sale, I Whole units, J More units than now (%), K Profit now, L Profit at whole units, M Status (OK, Check: price, Check: cost, Check: cost at or above price, Check: discount, Check: units). Summary O4:P8: Rows listed (P4), Rows showing Check (P5), Discount wipes out profit (P6), Typed below row 54 (P7), CHECK (P8). Landscape.
- Table: A1:A2 notes. A4 "Margin \ Discount"; B4:K4 discounts 5% to 50% (inputs), A5:A17 margins 20% to 80% (inputs); B5:K17 percent more units or No (orange). Landscape.
- Deviations: G is ROUND(...,10) and I is ROUNDUP(ROUND(H,6),0) (safety for float leftovers); cost at or above price and units 0 or less are Check; F, G, K are blank on a Check row; Table inputs are editable.

Key formulas (row 5)
- M5 `=IF(AND(TRIM(A5)="",COUNTA(B5:E5)=0),"",IF(NOT(ISNUMBER(B5)),"Check: price",IF(B5<=0,"Check: price",IF(NOT(ISNUMBER(C5)),"Check: cost",IF(C5<0,"Check: cost",IF(C5>=B5,"Check: cost at or above price",IF(NOT(ISNUMBER(D5)),"Check: discount",IF(OR(D5<0,D5>1),"Check: discount",IF(NOT(ISNUMBER(E5)),"Check: units",IF(E5<=0,"Check: units","OK"))))))))))`
- F5 `=IF(M5<>"OK","",(B5-C5)/B5)`; G5 `=IF(M5<>"OK","",ROUND(B5*(1-D5)-C5,10))`
- H5 `=IF(M5<>"OK","",IF(G5<=0,"Discount wipes out profit",E5*(B5-C5)/G5))` (core `=units*(price-cost)/(price*(1-discount)-cost)`)
- I5 `=IF(ISNUMBER(H5),ROUNDUP(ROUND(H5,6),0),"")`; J5 `=IF(ISNUMBER(H5),H5/E5-1,"")`; K5 `=IF(M5<>"OK","",E5*(B5-C5))`; L5 `=IF(ISNUMBER(I5),I5*G5,"")`
- P4 `=SUMPRODUCT(--(M5:M54<>""))`; P5 `=P4-COUNTIF(M5:M54,"OK")`; P6 `=COUNTIF(H5:H54,"Discount wipes out profit")`; P7 `=COUNTA(A55:E300)`; P8 `=IF(AND(P4>0,P5=0,P7=0),"OK","Check")`
- Table B5 `=IF(AND(ISNUMBER($A5),ISNUMBER(B$4)),IF($A5<=B$4,"No",$A5/($A5-B$4)-1),"")` (core `=margin/(margin-discount)-1`)

Sample values (fictional; recalculated)
| Product | Price | Cost | Discount | Units now | Margin | Profit per unit on sale | Units needed | Whole | More units | Profit now | Profit at whole |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Candle | $40.00 | $16.00 | 20% | 100 | 60.0% | $16.00 | 150.00 | 150 | 50.00% | $2,400.00 | $2,400.00 |
| Candle | $40.00 | $16.00 | 25% | 100 | 60.0% | $14.00 | 171.43 | 172 | 71.43% | $2,400.00 | $2,408.00 |
| Candle | $40.00 | $16.00 | 10% | 100 | 60.0% | $20.00 | 120.00 | 120 | 20.00% | $2,400.00 | $2,400.00 |
| Mug | $25.00 | $15.00 | 10% | 80 | 40.0% | $7.50 | 106.67 | 107 | 33.33% | $800.00 | $802.50 |
| Mug | $25.00 | $15.00 | 30% | 80 | 40.0% | $2.50 | 320.00 | 320 | 300.00% | $800.00 | $800.00 |
| Sticker pack | $12.00 | $3.00 | 15% | 200 | 75.0% | $7.20 | 250.00 | 250 | 25.00% | $1,800.00 | $1,800.00 |
| Candle | $40.00 | $16.00 | 60% | 100 | 60.0% | $0.00 | Discount wipes out profit | | | $2,400.00 | |
| Candle | $40.00 | $16.00 | 70% | 100 | 60.0% | -$4.00 | Discount wipes out profit | | | $2,400.00 | |
P4:P8: 8, 0, 2, 0, OK. Table (recalculated): 60% margin row 9.09%, 20.00%, 33.33%, 50.00%, 71.43%, 100.00%, 140.00%, 200.00%, 300.00%, 500.00% for 5% to 50% off; 40% row 14.29%, 33.33%, 60.00%, 100.00%, 166.67%, 300.00%, 700.00%, No, No, No; 20% row 33.33%, 100.00%, 300.00%, then No; 75% row 7.14% to 200.00%.

Edge cases tested (all matched a Python model)
- Discount = margin (60%: $0.00, wipes out), above it (65%: -$2.00), 0% (100 units, 0.00%), 100% (-$16.00, wipes out), 100.5%, -5%, 25 typed, "25%" text and blank (Check: discount). $29.99 with cost $18.8937 at 37% (discount equal to the margin with odd cents): profit per unit exactly $0.00, wipes out.
- Cost 0 (100% margin, 25% off: 133.33, 134, 33.33%, $4,020.00 vs $4,000.00), cost above and equal to price (Check: cost at or above price), negative and blank cost (Check: cost), "$40.00" text and price 0 (Check: price), units 0, -3 and text "100" (Check: units), 12.5 units (21.43, 22, $308.00), a name only (Check: price), numbers with no name (counted).
- A row on 55, no rows (CHECK Check), 50 random rows, 8 random sets (4 with bad rows), the Table with random margins and discounts, a text "15%" header (that column blank), a blank margin (that row blank), 0% and 100% headers.

Pitfalls the post must state
- Gross profit on unit cost only: no fees, shipping or fixed costs (link break-even-calculator, marketplace-fee-net-payout-calculator, sale-price-discount-calculator).
- No advice on running a sale or how deep, and no claim about how much a discount lifts sales.
- Type the discount as a percent (25%), not 25.
- Not tested: Excel, Google Sheets (float leftovers when the discount equals the margin are the open question; the sheet rounds G to 10 decimals).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (template copies)
- Status Check: discount: 25 typed instead of 25%, "25%" typed as text, a blank, or above 100%.
- Discount wipes out profit: the discount is the margin or more (60% and 70% on the 60% candle).
- Status Check: price: "$40.00" typed as text, or 0.
- Whole units profit is a little more than now ($2,408.00 vs $2,400.00): rounding up to 172 units.

---
## 3. add-st-nd-rd-th-to-dates-and-numbers-ordinal-suffix-mod-choose-excel-google-sheets
Title (from topics): How to Add st, nd, rd and th to Dates and Numbers in Excel and Google Sheets
Template: templates/tidy-tabs-ordinal-dates.xlsx (tabs: Dates, Numbers, How to use)

Layout
- Dates: A1 Sign text, B1 "Pickup on" (input); C1 note. Header row 4, rows 5:104 (12 sample dates): A Date (input, MM/DD/YYYY); B Long, C Short, D Day only, E Sign line (B1 + date), F Status (OK, Check: not a date). Summary H4:I7: Dates listed (I4), Rows showing Check (I5), Typed below row 104 (I6), CHECK (I7). Landscape.
- Numbers: A1 note. Header row 4, rows 5:54 (14 sample numbers): A Number (input); B With suffix, C Status (OK, Check: text, Check: 0 or less, Check: decimal). Summary E4:F7 (same four). Landscape.
- Deviations: B1 sign text input (topic hard-coded "Pickup on"); C and E build on D (Day only) instead of repeating the suffix formula; E is wrapped in TRIM (a blank B1 gives no leading space).

Key formulas (row 5)
- Dates F5 `=IF(A5="","",IF(NOT(ISNUMBER(A5)),"Check: not a date",IF(A5<1,"Check: not a date","OK")))`
- B5 `=IF(F5<>"OK","",TEXT(A5,"dddd, mmmm ")&DAY(A5)&IF(AND(DAY(A5)>=11,DAY(A5)<=13),"th",CHOOSE(MOD(DAY(A5),10)+1,"th","st","nd","rd","th","th","th","th","th","th"))&TEXT(A5,", yyyy"))`
- D5 `=IF(F5<>"OK","",DAY(A5)&IF(AND(DAY(A5)>=11,DAY(A5)<=13),"th",CHOOSE(MOD(DAY(A5),10)+1,"th","st","nd","rd","th","th","th","th","th","th")))`; C5 `=IF(F5<>"OK","",TEXT(A5,"mmm ")&D5)`; E5 `=IF(F5<>"OK","",TRIM($B$1&" "&TEXT(A5,"dddd, mmmm ")&D5))`
- Numbers C5 `=IF(A5="","",IF(NOT(ISNUMBER(A5)),"Check: text",IF(A5<1,"Check: 0 or less",IF(A5<>INT(A5),"Check: decimal","OK"))))`; B5 `=IF(C5<>"OK","",A5&IF(AND(MOD(A5,100)>=11,MOD(A5,100)<=13),"th",CHOOSE(MOD(A5,10)+1,"th","st","nd","rd","th","th","th","th","th","th")))` (core `=n&IF(AND(MOD(n,100)>=11,MOD(n,100)<=13),"th",CHOOSE(MOD(n,10)+1,...))`)
- I4 `=SUMPRODUCT(--(F5:F104<>""))`; I5 `=SUMPRODUCT(--(LEFT(F5:F104,5)="Check"))`; I6 `=COUNTA(A105:A1000)`; I7 `=IF(AND(I4>0,I5=0,I6=0),"OK","Check")` (Numbers F4:F7 the same on C5:C54)

Sample values (fictional market and pickup dates; recalculated)
| Date | Long | Short | Day only | Sign line |
|---|---|---|---|---|
| 11/01/2026 | Sunday, November 1st, 2026 | Nov 1st | 1st | Pickup on Sunday, November 1st |
| 11/02/2026 | Monday, November 2nd, 2026 | Nov 2nd | 2nd | Pickup on Monday, November 2nd |
| 11/03/2026 | Tuesday, November 3rd, 2026 | Nov 3rd | 3rd | Pickup on Tuesday, November 3rd |
| 11/11/2026 | Wednesday, November 11th, 2026 | Nov 11th | 11th | Pickup on Wednesday, November 11th |
| 11/12/2026 | Thursday, November 12th, 2026 | Nov 12th | 12th | Pickup on Thursday, November 12th |
| 11/13/2026 | Friday, November 13th, 2026 | Nov 13th | 13th | Pickup on Friday, November 13th |
| 11/20/2026 | Friday, November 20th, 2026 | Nov 20th | 20th | Pickup on Friday, November 20th |
| 11/21/2026 | Saturday, November 21st, 2026 | Nov 21st | 21st | Pickup on Saturday, November 21st |
| 11/22/2026 | Sunday, November 22nd, 2026 | Nov 22nd | 22nd | Pickup on Sunday, November 22nd |
| 11/23/2026 | Monday, November 23rd, 2026 | Nov 23rd | 23rd | Pickup on Monday, November 23rd |
| 12/31/2026 | Thursday, December 31st, 2026 | Dec 31st | 31st | Pickup on Thursday, December 31st |
| 01/01/2027 | Friday, January 1st, 2027 | Jan 1st | 1st | Pickup on Friday, January 1st |
I4:I7: 12, 0, 0, OK. Numbers: 1st, 2nd, 3rd, 4th, 11th, 12th, 13th, 21st, 22nd, 23rd, 101st, 111th, 112th, 113th; F4:F7: 14, 0, 0, OK.

Edge cases tested (all matched a Python model)
- Every day of 11/2026, 12/2026, 01/2027, 02/2027 and 02/2028 (Tuesday, February 29th, 2028); numbers 1 to 1000 (20 workbooks of 50), 1011th, 1012th, 1013th, 10001st, 1000000th, 99th, 100th.
- 0 and -1 (Check: 0 or less), 2.5 and 3.000000001 (Check: decimal), text "21" (Check: text); a date-time 11/20/2026 6:45 PM (Friday, November 20th, 2026) and 11/21/2026 11:59:59 PM (21st); text "11/20/2026", 0, -5, 0.5 (Check: not a date).
- B1 "Market day:" (Market day: Friday, November 20th), B1 blank (starts with the weekday), B1 with a double space (TRIM collapses it); rows on 105 and 55, both tabs empty (CHECK Check), 4 random sets of 100 dates (1950 to 2097) and 50 numbers.

Pitfalls the post must state
- The result is text: keep the real date in A for sorting and math, use B to E for printing.
- Day and month names follow the app's language; checked in English only (non-English not tested).
- Link Microsoft's and Google's TEXT help pages for the format codes rather than claiming them from our run.
- Not tested: Excel, Google Sheets (TEXT "dddd"/"mmmm" output and MOD on a text "21" are the open questions; the template's guard shows Check either way).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe and template copies)
- "Market on 46346": a date joined straight to text (`="Market on "&A1`); use TEXT.
- 0th, -1th, 2.5nd and 21st from a text "21": the unguarded number formula (probe); the template shows Check for all four.
- Status Check: not a date: the date was typed or pasted as text.

---
## 4. convert-cups-tablespoons-teaspoons-recipe-units-convert-function-excel-google-sheets
Title (from topics): How to Convert Cups, Tablespoons and Teaspoons in Excel and Google Sheets
Template: templates/tidy-tabs-recipe-unit-converter.xlsx (tabs: Converter, Units, How to use)

Layout
- Converter: A1:A2 notes (type 0.5 not 1/2; volume only). Header row 4, rows 5:54 (6 sample rows): A Ingredient, B Amount, C From unit, D To unit (inputs; C and D dropdowns from Units!A5:A20, warning style so a typed name is still checked); E From code (helper), F To code (helper), G Result, H Status (OK, Check: amount, Check: from unit, Check: to unit). Summary J4:K9: Rows listed (K4), Rows showing Check (K5), Results showing Check units (K6), Units tab codes showing Check (K7), Typed below row 54 (K8), CHECK (K9). Landscape.
- Units: A1:A2 notes. Header row 4, rows 5:20: A Friendly name, B CONVERT code (inputs; 7 filled: Teaspoon tsp, Tablespoon tbs, Fluid ounce oz, Cup cup, Pint pt, Quart qt, Gallon gal; rows 12 to 20 free for the reader's own names), C Teaspoons in one (Check code red). Grid E4:L11: 1 of the unit down = how many of the unit across. Landscape.
- Deviations: Units rows to 20 and codes TRIMmed; the grid; sixth sample row Lemon juice 0.0625 Cup -> Teaspoon (3); G is blank on a Check row; blank amount is Check: amount.

Key formulas (row 5)
- E5 `=IF(TRIM(C5)="","",IFERROR(TRIM(INDEX(Units!$B$5:$B$20,MATCH(TRIM(C5),Units!$A$5:$A$20,0))),""))` (F5 the same on D5)
- H5 `=IF(AND(TRIM(A5)="",B5="",TRIM(C5)="",TRIM(D5)=""),"",IF(NOT(ISNUMBER(B5)),"Check: amount",IF(B5<0,"Check: amount",IF(E5="","Check: from unit",IF(F5="","Check: to unit","OK")))))`
- G5 `=IF(H5<>"OK","",IFERROR(CONVERT(B5,E5,F5),"Check units"))` (core `=CONVERT(amount,"cup","tbs")`)
- Units C5 `=IF(TRIM(B5)="","",IFERROR(CONVERT(1,TRIM(B5),"tsp"),"Check code"))`; grid F5 `=IFERROR(CONVERT(1,TRIM($B5),TRIM($B$5)),"Check")` (column G uses $B$6, ... L uses $B$11); E5 `=A5`, F4 `=A5` ... L4 `=A11`
- K4 `=SUMPRODUCT(--(H5:H54<>""))`; K5 `=SUMPRODUCT(--(LEFT(H5:H54,5)="Check"))`; K6 `=COUNTIF(G5:G54,"Check units")`; K7 `=COUNTIF(Units!C5:C20,"Check code")`; K8 `=COUNTA(A55:D300)`; K9 `=IF(AND(K4>0,K5=0,K6=0,K7=0,K8=0),"OK","Check")`

Sample values (fictional batch; recalculated)
| Ingredient | Amount | From | To | Codes | Result |
|---|---|---|---|---|---|
| Honey | 2.5 | Cup | Fluid ounce | cup -> oz | 20 |
| Vanilla extract | 3 | Teaspoon | Tablespoon | tsp -> tbs | 1 |
| Water | 1 | Gallon | Cup | gal -> cup | 16 |
| Fragrance oil | 24 | Teaspoon | Cup | tsp -> cup | 0.5 |
| Milk | 1 | Quart | Cup | qt -> cup | 4 |
| Lemon juice | 0.0625 | Cup | Teaspoon | cup -> tsp | 3 |
K4:K9: 6, 0, 0, 0, 0, OK. Units C (teaspoons in one): 1, 3, 6, 48, 96, 192, 768. Grid: Cup row 48, 16, 8, 1, 0.5, 0.25, 0.0625; Tablespoon row 3, 1, 0.5, 0.0625, 0.0313, 0.0156, 0.0039 (shown to 4 decimals); Gallon row 768, 256, 128, 16, 8, 4, 1.

Edge cases tested (all matched a Python model of exact US ratios)
- All 49 pairs, amount 1 and random amounts; 0 (0), 0.0001 tsp -> gal, 1000 gal -> 768,000 tsp.
- -1, text "1/2" (written as text by openpyxl, not typed by hand), "2 cups", blank amount, ingredient only, units only (Check: amount); "Cup " and " Cup", "Tablespoon " (TRIM: 16); "cup", "teaspoon" (lowercase names match: 48); "Tbsp" not on the list (Check: from unit); blank To / From (Check: to unit / from unit).
- Units rows added: Tbsp -> tbs (2 Tbsp = 6 tsp, CHECK OK); code "tbsp", "ozm", "Cup" (Check code; any row using them Check units; CONVERT codes are case-sensitive); code " cup " (48 after TRIM). A row on 55, no rows (CHECK Check), 6 random sets of 50 rows (3 with negatives, text amounts, unknown and upper-case names).

Pitfalls the post must state
- Volume only; weight from volume needs a density (not given, not a CONVERT job). No recipe or food safety advice.
- Exact codes: tsp, tbs, oz (fluid ounce), cup, pt, qt, gal; "tbsp", "Cup" and "ozm" fail. Codes for Excel and Sheets come from Microsoft's and Google's CONVERT help pages (read and link them; do not quote other apps' error names).
- Type 0.5, not 1/2 (typing 1/2 by hand was not tested; a text "1/2" shows Check: amount).
- Not tested: Excel, Google Sheets. US units only; the liter check stays out of the post.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe and template copies)
- #VALUE! from CONVERT: "tbsp", "Cup" with a capital C, "ozm" from a cup, or " cup" with a leading space (probe); the template shows Check units / Check code, and TRIM fixes the space.
- Check: from unit / to unit: a name not on the Units tab (Tbsp); add it in Units rows 12 to 20 with its code.
- Check: amount: text such as "2 cups" or "1/2", a blank, or a negative.
- Weight rows are not exact: 1 lbm -> 16.0000023429409 ozm in LibreOffice (probe), and this sheet is volume only.

---
## 5. percent-of-time-elapsed-vs-percent-complete-project-status-excel-google-sheets
Title (from topics): How to Compare Percent of Time Elapsed With Percent Complete in Excel and Google Sheets
Template: templates/tidy-tabs-project-time-elapsed-tracker.xlsx (tabs: Projects, How to use)

Layout
- Projects: A1 As-of date, B1 11/20/2026 (input); A2 Tolerance (points), B2 10% (input); C1:C2 notes. Header row 4, rows 5:54 (8 sample projects): A Project, B Client, C Start date, D Due date, E Percent done (inputs); F Time elapsed %, G Done minus elapsed, H Days left, I Status (Done, Overdue, Behind, On track, Ahead, Check: as-of date, Check: tolerance, Check: dates, Check: due before start, Check: percent done). Summary K4:L12: Projects listed (L4), Done (L5), Overdue (L6), Behind (L7), On track (L8), Ahead (L9), Rows showing Check (L10), Typed below row 54 (L11), CHECK (L12). Landscape.
- Deviations: no optional bar column J (not built); status words with "Check: ..." (topic "Check dates"; hook 5 must say Check: due before start); G is ROUND(E-F,10); tolerance must be 0% to under 100% (10 typed is Check); blank percent done is Check; the sample's Brochure row (dates backwards, from the topic) makes L12 show Check by design (How to use line 8).

Key formulas (row 5)
- F5 `=IF(OR(NOT(ISNUMBER(C5)),NOT(ISNUMBER(D5)),NOT(ISNUMBER($B$1))),"",IF(D5<C5,"",IF(D5=C5,IF($B$1>=D5,1,0),MIN(1,MAX(0,($B$1-C5)/(D5-C5))))))` (core `=MIN(1,MAX(0,(asof-start)/(due-start)))`)
- G5 `=IF(AND(ISNUMBER(F5),ISNUMBER(E5)),IF(AND(E5>=0,E5<=1),ROUND(E5-F5,10),""),"")`; H5 `=IF(ISNUMBER(F5),MAX(0,D5-$B$1),"")`
- I5 `=IF(AND(TRIM(A5)="",TRIM(B5)="",C5="",D5="",E5=""),"",IF(NOT(ISNUMBER($B$1)),"Check: as-of date",IF(NOT(AND(ISNUMBER($B$2),$B$2>=0,$B$2<1)),"Check: tolerance",IF(NOT(AND(ISNUMBER(C5),ISNUMBER(D5))),"Check: dates",IF(D5<C5,"Check: due before start",IF(NOT(ISNUMBER(G5)),"Check: percent done",IF(E5>=1,"Done",IF($B$1>D5,"Overdue",IF(G5<-$B$2,"Behind",IF(G5>$B$2,"Ahead","On track"))))))))))`
- L4 `=SUMPRODUCT(--(I5:I54<>""))`; L5:L9 COUNTIF of each status; L10 `=SUMPRODUCT(--(LEFT(I5:I54,5)="Check"))`; L11 `=COUNTA(A55:E300)`; L12 `=IF(AND(L4>0,L10=0,L11=0),"OK","Check")`

Sample values (fictional; as-of 11/20/2026, tolerance 10%; recalculated)
| Project | Client | Start | Due | Done | Elapsed | Done minus elapsed | Days left | Status |
|---|---|---|---|---|---|---|---|---|
| Logo refresh | Maple Street Bakery | 11/02/2026 | 11/30/2026 | 50% | 64.29% | -14.29% | 10 | Behind |
| Website rebuild | Harbor Yoga | 11/10/2026 | 12/10/2026 | 20% | 33.33% | -13.33% | 20 | Behind |
| Holiday catalog | Pine & Co. Gifts | 11/01/2026 | 11/15/2026 | 80% | 100.00% | -20.00% | 0 | Overdue |
| Menu boards | Lakeside Cafe | 11/01/2026 | 11/15/2026 | 100% | 100.00% | 0.00% | 0 | Done |
| Spring lookbook | Fern & Thread | 12/01/2026 | 12/20/2026 | 0% | 0.00% | 0.00% | 30 | On track |
| Flyer edit | Main St. Hardware | 11/20/2026 | 11/20/2026 | 0% | 100.00% | -100.00% | 0 | Behind |
| Social posts | Bloom Florist | 11/18/2026 | 11/22/2026 | 50% | 50.00% | 0.00% | 2 | On track |
| Brochure (dates typed backwards) | Cedar Dental | 11/30/2026 | 11/10/2026 | 30% | | | | Check: due before start |
L4:L12: 8, 1, 1, 3, 2, 0, 1, 0, Check.

Edge cases tested (all matched a Python model)
- As-of before start (0%, On track, 15 days), on start (0%), on due (100%; 95% done is -5 points, On track), after due and 100% (Done), start = due after the as-of (0%, 5 days) and before it (100%; 50% done -> Overdue), due before start, text start and blank due (Check: dates).
- 50 typed, blank, -10%, 120% and "50%" text (Check: percent done; elapsed and days left still show), exactly -10 and +10 points (On track), 0.3 done vs 0.4 elapsed (On track), 90% at 50% elapsed (Ahead), date-times 11/10 12:00 PM to 11/30 6:00 PM (46.91%, days left 10.75 shown as 11), name only (Check: dates).
- Tolerance 0 (same statuses as the sample), 10 typed and blank (every row Check: tolerance), 25%; as-of blank and text (Check: as-of date); as-of 12/31/2026 (6 Overdue, 1 Done); as-of 10/01/2026 (4 Ahead, 2 On track, 1 Done); a row on 55, no rows, 50 random rows, 8 random sets.

Pitfalls the post must state
- Calendar days only; no business-day version was built or tested (no NETWORKDAYS claim). The as-of date is typed (no TODAY).
- Percent done is the reader's estimate; the tolerance is their choice; no project management claims.
- Pick one day-count rule: counting both end days (+1) gives 65.52% instead of 64.29%.
- Not tested: Excel, Google Sheets (comparison exactly at the tolerance edge is the open question; the sheet rounds G to 10 decimals).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe and template copies)
- Elapsed over 100% or below 0%: the plain (as-of - start)/(due - start) gave 135.71% (overdue) and -57.89% (not started); MIN/MAX fixes it.
- #DIV/0!: a one-day job (start = due) in the plain formula; the template shows 0% before that day and 100% on it.
- Every row shows Check: tolerance: 10 typed instead of 10%, or B2 blank.
- Check: percent done: 50 typed instead of 50%, a blank, or a value above 100%.

---
What was NOT run (all five): Microsoft Excel, Google Sheets, any non-US language setting, more rows than the slots (504 subscribers / 24 months, 54 products, 104 dates / 54 numbers, 54 conversions / 20 unit names, 54 projects), a printer (PDF export only), typing 1/2 by hand, revenue or MRR, churn benchmarks, a business-day elapsed version, the optional bar column, weight or density conversion. Nothing is claimed about Excel or Sheets.
