# Handoff 2026-11-21: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-21.md (researcher) and tasks/hooks-2026-11-21.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`shared_expense_split`, `week_number_from_date` with the helper `_week_cells`, `economic_order_quantity`, `shelf_bin_location_finder`, `url_handle_from_title` with `URL_HANDLE_SPACE_CHARS` and `_url_handle_formula`, registered after `project_time_elapsed_tracker()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-21 build and test"). Heroes are not made yet (designer, see each section).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 168 edited workbooks recalculated and compared (207,831 cell comparisons, 0 mismatches against independent Python models; a mutation run with one wrong formula per template produced mismatches in all five, so the harness catches errors), plus 6 probe workbooks of the traps read by hand; no error value in any recalculated template workbook. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text. Two formula problems and five layout problems were found and fixed (tasks/briefs.md).

Repo note: `python3 scripts/build-templates.py` ran (exit 0) but rewrote every older template byte-wise; those were restored with `git checkout -- templates/`, and the five new files were rebuilt by running only the five new functions. `git status`: scripts/build-templates.py, tasks/briefs.md, this handoff, five new .xlsx.

Post file names: `posts/ready/2026-11-21-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks every status starting "Check" (and "Fix:" on the handle sheet) and any CHECK cell that is not OK; orange marks repeated SKUs on the Map, Duplicate (numbered) on the handle sheet and Differ on the Calendar check tab; green marks Lowest on the Cost table. Functions used in the five files: AND, COLUMN, COLUMNS, COUNT, COUNTA, COUNTIF, DATE, DAY, FIND, IF, IFERROR, INDEX, INDIRECT, INT, ISERROR, ISNUMBER, LEFT, LEN, LOWER, MATCH, MAX, MID, MIN, MONTH, NOT, OR, ROUND, ROUNDUP, ROW, ROWS, SQRT, SUBSTITUTE, SUM, SUMIF, SUMPRODUCT, TEXT, TRIM, WEEKDAY, WEEKNUM, YEAR. No TODAY, FILTER, UNIQUE, LET, LAMBDA, XLOOKUP, TEXTSPLIT, SORT, MINIFS/MAXIFS, ISOWEEKNUM; 0 `_xlfn` in the files. Help links were not checked by the builder.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-21-team-01-split-shared-expenses-between-partners-who-owes-whom-sumif-excel-google-sheets | tidy-tabs-shared-expense-split.xlsx | 54 workbooks, 13,338 comparisons |
| 2 | 2026-11-21-team-02-week-number-from-date-us-and-iso-week-week-of-month-weeknum-excel-google-sheets | tidy-tabs-week-number-from-date.xlsx | 26 workbooks, 130,052 comparisons |
| 3 | 2026-11-21-team-03-economic-order-quantity-how-much-to-order-at-once-eoq-sqrt-excel-google-sheets | tidy-tabs-economic-order-quantity.xlsx | 44 workbooks, 24,948 comparisons |
| 4 | 2026-11-21-team-04-find-value-in-grid-return-row-and-column-label-shelf-bin-map-sumproduct-excel-google-sheets | tidy-tabs-shelf-bin-location-finder.xlsx | 24 workbooks, 7,512 comparisons |
| 5 | 2026-11-21-team-05-turn-product-titles-into-url-handles-slugs-lower-substitute-excel-google-sheets | tidy-tabs-url-handle-from-title.xlsx | 20 workbooks, 31,981 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets."

Hooks check (tasks/hooks-2026-11-21.md, chosen openings): every number matches the built files: 1 ($340.82, Alex $209.00, "Pay Alex $34.29", "Pay Alex $61.11", $340.83 without the fix), 2 (12/27/2026 US 53 / ISO 52, 01/01/2027 US 1 / ISO 53 of 2026, "2026-W47"), 3 (707.11, 708, every 107.7 days, $169.71, 600 -> $172.00, 800 -> $171.00; trap 70.71 and 204.12 are probe values, the template shows Check for 20 typed and cannot detect monthly use), 4 (SKU-147 -> F-Bin 8, MUG-BLU-12 -> A-Bin 6, "Not on map", "In 2 bins"; the #REF! trap is a probe value, the template never shows it), 5 (lavender-and-honey-soy-candle-8-oz, quick formula lavender-&-honey-soy-candle---8-oz). Formulas in openings 1 and 3 are the core forms; the cells hold guarded versions (below).

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=...)` fits the whole sheet on one page): the gray formula fill runs to the last input row (Expenses 204, Dates 504, Titles 204, Supplies and Find 54), so a whole-sheet render of Expenses, Dates or Titles will be tiny. Render People, Find, Supplies or Cost table as they are, and for Dates or Titles use a copy with the empty rows deleted (or a print area) before rendering.

---
## 1. split-shared-expenses-between-partners-who-owes-whom-sumif-excel-google-sheets
Title (from topics): How to Split Shared Expenses and See Who Owes Whom in Excel and Google Sheets
Template: templates/tidy-tabs-shared-expense-split.xlsx (tabs: Expenses, People, How to use)

Layout
- Expenses: A1:A2 notes. Header row 4, rows 5:204 (6 sample rows): A Date, B Description, C Paid by (dropdown from People!$A$5:$A$10, warning style so a typed name is still checked), D Amount (inputs); E Status. Summary G4:H9: Expense rows listed (H4), Rows showing Check (H5), Rows whose payer is not on People (H6), Refunds (negative, counted) (H7), Typed below row 204 (H8), Total spent (H9). Landscape.
- People: A1 Total spent, B1 formula; C1 and A2 notes. Header row 4, rows 5:10 (up to 6 people; row 5 is the treasurer): A Name, B Share % (inputs; sample =1/3 in each), C Paid, D Fair share, E Balance (+ is owed, - owes), F Settle line, G Status. Summary I4:J13: Shares ready (J4), Shares add to (J5), People rows showing Check (J6), Total spent (J7), Paid by people on this list (J8), Spent under names not on this list (J9), Fair shares add to (J10), Balances add to (J11), Expense rows showing Check (J12), CHECK (J13). Landscape.
- Deviations from the topic: refunds (negative amounts) are allowed and counted, Status "Refund (counted)"; D5 is its own formula with the leftover cent and D6:D10 are plain ROUND (not the ROW()=5 form); C:F are blank on blank-name rows and D:F stay blank until J4 is OK; F5 reads "Collects $95.40, pays out $0.00"; payer check H6 has the blank-row guard `(C5:C204<>"")`.

Key formulas (exact text in the file)
- People B1 `=SUM(Expenses!D5:D204)`
- C5 `=IF(TRIM(A5)="","",SUMIF(Expenses!$C$5:$C$204,A5,Expenses!$D$5:$D$204))` (core `=SUMIF(paid_by,name,amount)`)
- D5 (treasurer) `=IF(OR($J$4<>"OK",TRIM(A5)=""),"",ROUND(ROUND($B$1*B5,2)+$B$1-SUMPRODUCT(ROUND($B$1*$B$5:$B$10,2)),2))`
- D6 `=IF(OR($J$4<>"OK",TRIM(A6)=""),"",ROUND($B$1*B6,2))` (same down to D10; core `=ROUND(total*share,2)`)
- E5 `=IF(OR(C5="",D5=""),"",ROUND(C5-D5,2))`
- F5 `=IF(E5="","","Collects "&TEXT(-SUMIF(E6:E10,"<0"),"$#,##0.00")&", pays out "&TEXT(SUMIF(E6:E10,">0"),"$#,##0.00"))`
- F6 `=IF(E6="","",IF(E6<0,"Pay "&$A$5&" "&TEXT(-E6,"$#,##0.00"),IF(E6>0,$A$5&" pays you "&TEXT(E6,"$#,##0.00"),"Even")))`
- G6 `=IF(AND(TRIM(A6)="",B6=""),"",IF(TRIM(A6)="","Check: share with no name",IF(A6<>TRIM(A6),"Check: extra space in name",IF(COUNTIF($A$5:$A$10,A6)>1,"Check: name twice",IF(NOT(ISNUMBER(B6)),"Check: share",IF(B6>1,"Check: share (50% not 50)",IF(B6<0,"Check: share","OK")))))))`; G5 starts with `IF(AND(TRIM(A5)="",COUNTA(A6:A10)>0),"Check: treasurer goes in row 5",` then the same test.
- J4 `=IF(AND(TRIM(A5)<>"",ROUND(SUM(B5:B10),6)=1,J6=0),"OK","Check")`; J9 `=ROUND(J7-J8,2)`; J10 `=ROUND(SUM(D5:D10),2)`; J11 `=ROUND(SUM(E5:E10),2)`; J13 `=IF(AND(J4="OK",J9=0,ROUND(J10-J7,2)=0,ROUND(J11,2)=0,J12=0,Expenses!H8=0),"OK","Check")`
- Expenses E5 `=IF(AND(A5="",TRIM(B5)="",TRIM(C5)="",D5=""),"",IF(AND(A5<>"",NOT(ISNUMBER(A5))),"Check: date is text",IF(TRIM(C5)="","Check: no payer",IF(COUNTIF(People!$A$5:$A$10,C5)=0,"Check: payer not on People",IF(D5="","Check: no amount",IF(NOT(ISNUMBER(D5)),"Check: amount is text",IF(D5<0,"Refund (counted)","OK")))))))`
- Expenses H6 (payer check) `=SUMPRODUCT((C5:C204<>"")*(COUNTIF(People!$A$5:$A$10,C5:C204)=0))`

Sample values (fictional; recalculated)
Expenses: 09/05/2026 Booth fee, fall season, Alex, $150.00; 09/05/2026 Tent weights, Bea, $45.99; 09/12/2026 Ice and water, Cal, $12.50; 09/19/2026 Card reader, Alex, $59.00; 10/03/2026 Signs and price tags, Bea, $33.33; 10/17/2026 Gas for the van, Cal, $40.00. Total $340.82.
| Name | Share | Paid | Fair share | Balance | Settle line |
|---|---|---|---|---|---|
| Alex | 33.33% | $209.00 | $113.60 | $95.40 | Collects $95.40, pays out $0.00 |
| Bea | 33.33% | $79.32 | $113.61 | -$34.29 | Pay Alex $34.29 |
| Cal | 33.33% | $52.50 | $113.61 | -$61.11 | Pay Alex $61.11 |
J4:J13: OK, 100.00%, 0, $340.82, $340.82, $0.00, $340.82, $0.00, 0, OK. Expenses H4:H9: 6, 0, 0, 0, 0, $340.82. Shares 50%/30%/20%: fair $170.41, $102.25, $68.16; balances +$38.59, -$22.93, -$15.66; "Collects $38.59, pays out $0.00", "Pay Alex $22.93", "Pay Alex $15.66". Probe: `=3*ROUND(B1/3,2)` gave $340.83 (one cent more than spent).

Edge cases tested (all matched the Python model)
- Trailing space payer "Bea " $10.00 napkins (the trap): Total $350.82, Bea paid stays $79.32, fair $116.94 each, balances +$92.06, -$37.62, -$64.44, Balances add to -$10.00, J9 $10.00, Expenses H6 1, Status Check: payer not on People, CHECK Check. Same numbers for a name not on the list ("Dana") and a payer of only spaces (Check: no payer).
- Payer typed "bea": counted for Bea (LibreOffice SUMIF and COUNTIF ignore case), CHECK OK. People "Bea" and "bea": both Check: name twice. "Bea " on People: Check: extra space in name.
- Refund -$12.50 by Bea: total $328.32, fair $109.44 each, +$99.56, -$42.62, -$56.94, Refund (counted), H7 1, CHECK OK.
- One person paid everything: +$227.22, -$113.61, -$113.61. Cal paid nothing (his rows paid by Bea): Bea +$18.21 "Alex pays you $18.21", Cal -$113.61, F5 "Collects $113.61, pays out $18.21". No expenses: all $0.00, "Even". Cal at 0%: +$52.50 "Alex pays you $52.50". Two people 50/50: +$91.09 / -$91.09.
- Shares: 50 typed (Check: share (50% not 50), Shares add to 5050.00%), 99% and 101% (J4 Check, D:F blank), "33%" as text, -10% (Check: share), a share with no name, a blank share, treasurer row blank (Check: treasurer goes in row 5), a blank People row in the middle (same results as the sample).
- Amount "$10.00" as text (Check: amount is text, not counted), date as text (Check: date is text; amount still counted), no amount, a row on 205 (H8, CHECK Check), 6 people at =1/6 and at 30/25/20/10/10/5 with 200 expense rows, 24 random groups of 2 to 6 people (every clean group: balances $0.00, fair shares = total).

Pitfalls the post must state
- Arithmetic only: one line that the sheet does not decide how partners should split costs or how to record shared costs for a partnership or for taxes. No payment apps.
- Settling through one person (row 5) gives at most (people - 1) payments; it is not always the fewest payments.
- Names must match exactly except for case in LibreOffice; say the case result is a LibreOffice Calc result.
- Not tested: Excel, Google Sheets (case-insensitive SUMIF/COUNTIF, TEXT with "$#,##0.00").

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Balances do not add to $0.00: a payer typed with a space or not on the People list ($10.00 -> -$10.00, J9 $10.00).
- The shares add to $340.83: each share rounded on its own; the treasurer row takes the leftover cent.
- Fair share is blank: shares do not add to 100% (99%, 101%, 50 typed) or a People row shows Check.
- A refund: type it as a negative amount; it lowers the total.

Hero: sheet "People", hide_cols "HIJ" (summary) and optionally "G" (Status, all OK), rows 1 to 7: B1 $340.82, the three balances and the settle lines. Expenses A4:E10 is an optional inset (hide "FGH").

---
## 2. week-number-from-date-us-and-iso-week-week-of-month-weeknum-excel-google-sheets
Title (from topics): How to Get the Week Number From a Date (US and ISO Weeks) in Excel and Google Sheets
Template: templates/tidy-tabs-week-number-from-date.xlsx (tabs: Dates, Calendar check, How to use)

Layout
- Dates: A1:A2 notes. Header row 4, rows 5:504 (15 sample dates): A Date (input), B US week (Sunday start), C ISO week (Monday start), D ISO year, E Year-week label (ISO), F Week of month (Sunday start), G ISO week starts (Monday), H Status (OK, Check: date is text, Check: not a date), I ISO week, long formula (cross-check), J Label with YEAR() (wrong, comparison). Summary L4:M10: Dates listed (M4), Rows showing Check (M5), ISO week differs from the long formula (M6), Year-week to count (M7, yellow input, sample "2026-W47"), Rows with that year-week label (M8), Typed below row 504 (M9), CHECK (M10). Landscape.
- Calendar check: A1 First day, B1 12/01/2026 (input). Header row 4, rows 5:66 (62 days): A Date, B Day, C US week, D ISO week, E ISO year, F Year-week label, G Week of month, H US and ISO week (Differ orange / Same). Landscape.
- Deviations: G uses INT(A5) so a time part is dropped; the long ISO formula is visible as column I with M6 counting differences; J and the Calendar check Day and Differ columns added; the sample has 3 dates in 2026-W47, so M8 is 3 (the researcher's probe had 2).

Key formulas (exact text; every Dates formula is wrapped in `IF($H5<>"OK","",...)`)
- B5 `=IF($H5<>"OK","",INT((A5-DATE(YEAR(A5),1,1)+WEEKDAY(DATE(YEAR(A5),1,1))-1)/7)+1)`
- C5 `=IF($H5<>"OK","",WEEKNUM(A5,21))`
- D5 `=IF($H5<>"OK","",YEAR(A5-WEEKDAY(A5,2)+4))`
- E5 `=IF(C5="","",D5&"-W"&TEXT(C5,"00"))`
- F5 `=IF($H5<>"OK","",INT((DAY(A5)+WEEKDAY(DATE(YEAR(A5),MONTH(A5),1))-2)/7)+1)`
- G5 `=IF($H5<>"OK","",INT(A5)-WEEKDAY(A5,2)+1)`
- H5 `=IF(A5="","",IF(NOT(ISNUMBER(A5)),"Check: date is text",IF(A5<1,"Check: not a date","OK")))`
- I5 `=IF($H5<>"OK","",INT((A5-DATE(YEAR(A5-WEEKDAY(A5-1)+4),1,3)+WEEKDAY(DATE(YEAR(A5-WEEKDAY(A5-1)+4),1,3))+5)/7))`
- J5 `=IF($H5<>"OK","",YEAR(A5)&"-W"&TEXT(C5,"00"))`
- M6 `=SUMPRODUCT(--(C5:C504<>I5:I504))`; M8 `=COUNTIF(E5:E504,M7)`; M10 `=IF(AND(M4>0,M5=0,M6=0,M9=0),"OK","Check")`
- Calendar check A5 `=IF(ISNUMBER($B$1),INT($B$1),"")`, A6 `=IF(A5="","",A5+1)`, B5 `=IF(A5="","",TEXT(A5,"ddd"))`, H5 `=IF(A5="","",IF(C5<>D5,"Differ","Same"))`; C:G use the same formulas as Dates B, C, D, E, F on column A.

Sample values (fictional dates; recalculated)
| Date | US week | ISO week | Label | Week of month | ISO Monday | YEAR() label |
|---|---|---|---|---|---|---|
| 12/29/2025 | 53 | 1 | 2026-W01 | 5 | 12/29/2025 | 2025-W01 |
| 01/04/2026 | 2 | 1 | 2026-W01 | 2 | 12/29/2025 | 2026-W01 |
| 01/05/2026 | 2 | 2 | 2026-W02 | 2 | 01/05/2026 | 2026-W02 |
| 11/01/2026 | 45 | 44 | 2026-W44 | 1 | 10/26/2026 | 2026-W44 |
| 11/16/2026 | 47 | 47 | 2026-W47 | 3 | 11/16/2026 | 2026-W47 |
| 11/21/2026 | 47 | 47 | 2026-W47 | 3 | 11/16/2026 | 2026-W47 |
| 11/22/2026 | 48 | 47 | 2026-W47 | 4 | 11/16/2026 | 2026-W47 |
| 11/30/2026 | 49 | 49 | 2026-W49 | 5 | 11/30/2026 | 2026-W49 |
| 12/27/2026 | 53 | 52 | 2026-W52 | 5 | 12/21/2026 | 2026-W52 |
| 12/28/2026 | 53 | 53 | 2026-W53 | 5 | 12/28/2026 | 2026-W53 |
| 12/31/2026 | 53 | 53 | 2026-W53 | 5 | 12/28/2026 | 2026-W53 |
| 01/01/2027 | 1 | 53 | 2026-W53 | 1 | 12/28/2026 | 2027-W53 |
| 01/03/2027 | 2 | 53 | 2026-W53 | 2 | 12/28/2026 | 2027-W53 |
| 01/04/2027 | 2 | 1 | 2027-W01 | 2 | 01/04/2027 | 2027-W01 |
| 05/01/2027 | 18 | 17 | 2027-W17 | 1 | 04/26/2027 | 2027-W17 |
ISO year (D) is 2026 for every 2025-12-29 to 2027-01-03 row and 2027 for the last two. I matches C on every row. M4:M10: 15, 0, 0, 2026-W47, 3, 0, OK. Calendar check (12/01/2026): 12/01 Tue 49/49 Same; every Sunday Differ (12/06 50 vs 49); 12/27 53 vs 52; 01/01/2027 1 vs 53.

Edge cases tested (all matched Python's isocalendar and plain day counts)
- Every day 01/01/2020 to 12/31/2029 (3,653 days), Dec 20 to Jan 8 for 2015 to 2035, 500 random dates 1901 to 2097 with times, Dec 31 and Feb 29 of every leap year 1904 to 2096: 12/31/2028 is US week 54, ISO 52, 2028-W52, week of month 6, Monday 12/25/2028. Months starting on Saturday (05/01/2027) and Sunday (11/01/2026): week of month 1.
- Date with a time: 01/01/2027 12:00:01 AM gives US 1, ISO 53, 2026-W53, Monday 12/28/2026 (no time). 12/27/2026 6:30 PM gives 53 / 52.
- Text "11/21/2026", "2026-11-21" and a cell of spaces: Check: date is text (all columns blank); 0.5, 0, -3: Check: not a date. Probe: unguarded WEEKNUM of the text "11/21/2026" gave 47 in LibreOffice (converted silently), which is why Status checks for text.
- Count by label: "2026-W53" over 500 days gives 7; "2027-w01" lowercase still counts (COUNTIF ignores case). Row 505 (CHECK Check), empty tab (CHECK Check), Calendar B1 12/01/2028 (12/31/2028 54 vs 52), B1 as text (whole tab blank).

LibreOffice finding (probe, 3,653 days 01/01/2020 to 12/31/2029; not checked in Excel or Sheets)
- WEEKNUM(date,1) returned 1 where the counting formula gives 53 (54 on 12/31/2028) on 33 days, all between Dec 26 and Dec 31 (12/27/2026 to 12/31/2026; none in 2022). WEEKNUM(date,2) differed from a Monday-start count on 30 days. WEEKNUM(date,21) and ISOWEEKNUM (stored as `_xlfn.ISOWEEKNUM`) had 0 mismatches. The topic file says 37 and 33 days; use 33 and 30 (this run) or say "on several late-December days". Link the Microsoft and Google WEEKNUM help pages; do not claim what Excel or Sheets return.

Pitfalls the post must state
- US week here means Sunday start, the week of January 1 is week 1 (so week 53 or 54 exists). Week of month: Sunday start, the 1st is in week 1; there is no standard. Fiscal or 4-5-4 weeks: one line and a link. No payroll talk.
- Use the ISO year (D), not YEAR(date), in a year-week label.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The label says 2027-W53 for 01/01/2027: built with YEAR() (column J); use the ISO year.
- Week 1 on 12/28/2026: WEEKNUM(date,1) in LibreOffice (probe); the counting formula gives 53.
- Status Check: date is text: a date typed or pasted as text.
- Sunday dates show different US and ISO weeks (Calendar check Differ): expected, the weeks start on different days.

Hero: sheet "Dates", rows 4 to 19, hide_cols "HIKLM" (Status, cross-check, summary) to show A to G plus J (the YEAR() label beside the correct one). Without the trap column use "HIJKLM". Delete rows 20 to 504 in the render copy first (see note at the top). The Calendar check A4:H35 is an optional inset (Differ in orange).

---
## 3. economic-order-quantity-how-much-to-order-at-once-eoq-sqrt-excel-google-sheets
Title (from topics): How to Find How Much to Order at Once (Economic Order Quantity) in Excel and Google Sheets
Template: templates/tidy-tabs-economic-order-quantity.xlsx (tabs: Supplies, Cost table, How to use)

Layout
- Supplies: A1:A2 notes. Header row 4, rows 5:54 (5 sample rows): A Item, B Units used per year, C Cost per order, D Unit cost, E Holding cost % per year, F Case size (inputs); G Holding cost per unit per year, H EOQ, I Order quantity (rounded up to the case), J Orders per year, K Days between orders, L Yearly order cost, M Yearly holding cost, N Yearly total, O Status. Summary Q4:R8: Items listed (R4), Rows showing Check (R5), Yearly total, all OK items (R6), Typed below row 54 (R7), CHECK (R8). Landscape.
- Cost table: A1 Item, B1 (input with dropdown, sample Glass jars, 8 oz), A2 First quantity, B2 100, A3 Step, B3 100. Header row 4, rows 5:24: A Order quantity, B Yearly order cost, C Yearly holding cost, D Yearly total, E Lowest in table (green). Summary G4:H15: Row on Supplies (H4), That row's Status (H5), Units used per year (H6), Cost per order (H7), Holding cost per unit per year (H8), EOQ (H9), Order quantity (H10), Yearly total at that quantity (H11), Yearly total at the exact EOQ (H12), Lowest-cost quantity in the table (H13), Lowest yearly total in the table (H14), CHECK (H15). Landscape.
- Deviations: I5 has an inner ROUND; the Cost table is rows 5:24 with the total in D (topic: rows 12:31, total in B); 20 typed as the holding % is Check (no EOQ shown); two extra sample rows.

Key formulas (exact text; Supplies G:N are wrapped in `IF($O5<>"OK","",...)`)
- G5 `=IF($O5<>"OK","",D5*E5)`; H5 `=IF($O5<>"OK","",SQRT(2*B5*C5/G5))` (core `=SQRT(2*units_per_year*cost_per_order/(unit_cost*holding_pct))`)
- I5 `=IF($O5<>"OK","",ROUNDUP(ROUND(H5/F5,9),0)*F5)` (core `=ROUNDUP(eoq/case,0)*case`)
- J5 `=IF($O5<>"OK","",B5/I5)`; K5 `=IF($O5<>"OK","",365/J5)`; L5 `=IF($O5<>"OK","",J5*C5)`; M5 `=IF($O5<>"OK","",I5/2*G5)`; N5 `=IF($O5<>"OK","",L5+M5)`
- O5 `=IF(AND(TRIM(A5)="",COUNTA(B5:F5)=0),"",IF(NOT(ISNUMBER(B5)),"Check: units per year",IF(B5<=0,"Check: units per year",IF(NOT(ISNUMBER(C5)),"Check: cost per order",IF(C5<=0,"Check: cost per order",IF(NOT(ISNUMBER(D5)),"Check: unit cost",IF(D5<=0,"Check: unit cost",IF(NOT(ISNUMBER(E5)),"Check: holding %",IF(E5<=0,"Check: holding %",IF(E5>1,"Check: holding % (20% not 20)",IF(NOT(ISNUMBER(F5)),"Check: case size",IF(F5<1,"Check: case size",IF(F5<>INT(F5),"Check: case size","OK")))))))))))))`
- Cost table A5 `=IF(AND(ISNUMBER($B$2),ISNUMBER($B$3)),$B$2,"")`, A6 `=IF(A5="","",A5+$B$3)`; B5 `=IF(OR(A5="",$H$15<>"OK"),"",$H$6/A5*$H$7)`; C5 `=IF(OR(A5="",$H$15<>"OK"),"",A5/2*$H$8)`; D5 `=IF(B5="","",B5+C5)`; E5 `=IF(D5="","",IF(D5=MIN($D$5:$D$24),"Lowest",""))`
- H4 `=IFERROR(MATCH(B1,Supplies!$A$5:$A$54,0),"")`; H5 `=IF(H4="","Not on Supplies",INDEX(Supplies!$O$5:$O$54,H4))`; H6 `=IF(H5="OK",INDEX(Supplies!$B$5:$B$54,H4),"")` (H7 C, H8 G, H9 H, H10 I, H11 N the same way); H12 `=IF(H5="OK",H6/H9*H7+H9/2*H8,"")`; H13 `=IF(COUNT(D5:D24)=0,"",INDEX(A5:A24,MATCH(MIN(D5:D24),D5:D24,0)))`; H14 `=IF(COUNT(D5:D24)=0,"",MIN(D5:D24))`; H15 `=IF(AND(H5="OK",ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B2>0,B3>0),"OK","Check"),"Check")`

Sample values (fictional; recalculated)
| Item | Per year | Per order | Unit | Holding | Case | Holding/unit | EOQ | Order qty | Orders/yr | Days | Order cost | Holding cost | Total |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Glass jars, 8 oz | 2,400 | $25.00 | $1.20 | 20.0% | 12 | $0.2400 | 707.11 | 708 | 3.4 | 107.7 | $84.75 | $84.96 | $169.71 |
| Candle wicks | 5,000 | $8.00 | $0.10 | 25.0% | 100 | $0.0250 | 1,788.85 | 1,800 | 2.8 | 131.4 | $22.22 | $22.50 | $44.72 |
| Mailer boxes | 1,800 | $15.00 | $0.85 | 20.0% | 50 | $0.1700 | 563.60 | 600 | 3.0 | 121.7 | $45.00 | $51.00 | $96.00 |
| Soy wax, lb | 1,200 | $30.00 | $2.40 | 20.0% | 50 | $0.4800 | 387.30 | 400 | 3.0 | 121.7 | $90.00 | $96.00 | $186.00 |
| Fragrance oil, 16 oz | 300 | $12.00 | $9.50 | 20.0% | 6 | $1.9000 | 61.56 | 66 | 4.5 | 80.3 | $54.55 | $62.70 | $117.25 |
Unrounded: jars orders 3.38983, days 107.675, total 169.7058; R4:R8 5, 0, $613.67, 0, OK. Cost table (jars, 100 by 100): 100 $612.00, 200 $324.00, 300 $236.00, 400 $198.00, 500 $180.00, 600 $172.00, 700 $169.71 Lowest, 800 $171.00, 900 $174.67, 1,000 $180.00, 1,500 $220.00, 2,000 $270.00. H4:H15: 1, OK, 2,400, $25.00, $0.2400, 707.11, 708, $169.71, $169.71 (exact EOQ 169.7056; probe: $84.85 + $84.85), 700, $169.71, OK. Other items in the table: wicks lowest 1,800 at $44.72; boxes 600 at $96.00; wax 400 at $186.00; fragrance oil lowest is the first row, 100 at $131.00 (its EOQ 61.56 is below 100: set B2 and B3 smaller for small items). B2 600, B3 12 (jars): lowest 708 at $169.71.

Edge cases tested (all matched the Python model)
- 20 typed instead of 20% and 100.5%: Check: holding % (20% not 20); 100%: EOQ 316.23, 324. Holding 0: Check: holding %. Cost per order 0 or "$25.00" text: Check: cost per order. Case 1: 708; case 0 and 12.5: Check: case size. Units -2,400, blank unit cost, name only: Check.
- EOQ exactly 400 at case 100: 400. Units 1,250 / $8.00 / $0.10 / 20%: EOQ 1,000, order 1,000 at case 100 (the ROUND guard). EOQ 15.81 below one case of 50: 50, total $11.00. 10,000,000 units: EOQ 252,982.21, 253,000. Monthly use 200 typed in the yearly column: EOQ 204.12, 216, Status OK (the sheet cannot tell).
- Blank rows between items, no item name (still computed), a row on 55 (CHECK Check), empty tab, 50 random rows, 10 random sets; Cost table for each item, item typed "candle WICKS" (found), "Glass jar" (Not on Supplies, Check), an item on a Check row (Check), first quantity "100" as text and step -100 (Check).
- Probe (unguarded): holding 20 -> EOQ 70.71; monthly 200 -> 204.12; cost per order 0 -> EOQ 0 and #DIV/0! for orders per year; CEILING(EOQ,12) = ROUNDUP(EOQ/12,0)*12 = 708. LibreOffice-specific: ROUNDUP(H/F,0) without ROUND also gave 1,000 on the 1,250-unit row; the guard stays for other apps (not checked there).

Pitfalls the post must state
- EOQ is a starting point: steady use, the same cost per order; no quantity discounts (link quantity-price-tiers-lookup), lead time or safety stock (link reorder-point-calculator). Holding % is the reader's estimate; no typical rate, no accounting or tax talk. Rounding to a case: link round-order-quantity-up-to-case-size.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Status Check: holding % (20% not 20): 20 typed; unguarded it gives 70.71 instead of 707.11.
- EOQ looks about 3.5 times too small: monthly use typed (204.12 instead of 707.11); the sheet cannot flag it.
- Check: cost per order: 0, blank or "$25.00" text (unguarded: EOQ 0 and #DIV/0!).
- Lowest in table is the first row: the EOQ is below the first quantity; lower B2 and B3.

Hero: sheet "Supplies", rows 4 to 9, hide_cols "OPQR" (Status and summary); optionally also "JK". Second image or inset: sheet "Cost table", hide_cols "FGH", rows 1 to 24 (700 Lowest in green; 600 $172.00 and 800 $171.00 beside it).

---
## 4. find-value-in-grid-return-row-and-column-label-shelf-bin-map-sumproduct-excel-google-sheets
Title (from topics): How to Find a Value in a Grid and Return Its Row and Column Label in Excel and Google Sheets
Template: templates/tidy-tabs-shelf-bin-location-finder.xlsx (tabs: Map, Find, How to use)

Layout
- Map: A1:A2 notes. A4 "Shelf" (header), B4:I4 Bin 1 to Bin 8 (yellow inputs), A5:A10 A to F (yellow), B5:I10 SKUs (yellow; 48 bins). Summary K4:L8: Filled bins (L4), Empty bins (L5), Bins holding a SKU that is in another bin too (L6, those bins orange), Bins with a space before or after the SKU (L7), Map CHECK (L8). Landscape.
- Find: A1:A2 notes. Header row 4, rows 5:54 (11 sample lookups): A SKU (input), B Times on map, C Row number, D Column number, E Location, F Walk order, G Status (OK, Check: in N bins, Check: not on map, Check: space, or number vs text). Summary I4:J11: SKUs listed (J4), Found in one bin (J5), Not on map (J6), In more than one bin (J7), Rows showing Check (J8), Typed below row 54 (J9), Map CHECK (J10), CHECK (J11). Landscape.
- Deviations: C and D are blank unless B is exactly 1 (topic: always computed); G Status added with the TRIM test; the sample includes the traps, so Map CHECK and CHECK show Check by design (How to use line 9 says so).

Key formulas (exact text)
- B5 `=IF(TRIM(A5)="","",SUMPRODUCT(--(Map!$B$5:$I$10=A5)))`
- C5 `=IF(B5<>1,"",SUMPRODUCT((Map!$B$5:$I$10=A5)*(ROW(Map!$B$5:$I$10)-ROW(Map!$B$5)+1)))` (core `=SUMPRODUCT((grid=x)*(ROW(grid)-ROW(first_cell)+1))`)
- D5 `=IF(B5<>1,"",SUMPRODUCT((Map!$B$5:$I$10=A5)*(COLUMN(Map!$B$5:$I$10)-COLUMN(Map!$B$5)+1)))`
- E5 `=IF(B5="","",IF(B5=1,INDEX(Map!$A$5:$A$10,C5)&"-"&INDEX(Map!$B$4:$I$4,D5),IF(B5=0,"Not on map","In "&B5&" bins")))`
- F5 `=IF(B5=1,C5*100+D5,"")`
- G5 `=IF(B5="","",IF(B5=1,"OK",IF(B5>1,"Check: in "&B5&" bins",IF(SUMPRODUCT(--(TRIM(Map!$B$5:$I$10)=TRIM(A5)))>0,"Check: space, or number vs text","Check: not on map"))))`
- Map L4 `=COUNTA(B5:I10)`; L5 `=ROWS(B5:I10)*COLUMNS(B5:I10)-L4`; L6 `=SUMPRODUCT((B5:I10<>"")*(COUNTIF(B5:I10,B5:I10)>1))`; L7 `=SUMPRODUCT((B5:I10<>"")*(LEN(B5:I10)<>LEN(TRIM(B5:I10))))`; L8 `=IF(AND(L4>0,L6=0,L7=0),"OK","Check")`
- Find J11 `=IF(AND(J4>0,J8=0,J9=0,J10="OK"),"OK","Check")`

Sample values (fictional; recalculated)
Map: SKU-100 to SKU-147 across then down, with MUG-BLU-12 at A-Bin 6, CAN-LAV-08 at B-Bin 2 and E-Bin 5, the number 1001 at C-Bin 3, D-Bin 4 empty. L4:L8: 47, 1, 2, 0, Check.
| SKU typed | Times | Row | Col | Location | Walk | Status |
|---|---|---|---|---|---|---|
| SKU-100 | 1 | 1 | 1 | A-Bin 1 | 101 | OK |
| SKU-147 | 1 | 6 | 8 | F-Bin 8 | 608 | OK |
| MUG-BLU-12 | 1 | 1 | 6 | A-Bin 6 | 106 | OK |
| SKU-123 | 1 | 3 | 8 | C-Bin 8 | 308 | OK |
| SKU-131 | 1 | 4 | 8 | D-Bin 8 | 408 | OK |
| mug-blu-12 | 1 | 1 | 6 | A-Bin 6 | 106 | OK |
| "MUG-BLU-12 " (space) | 0 | | | Not on map | | Check: space, or number vs text |
| CAN-LAV-08 | 2 | | | In 2 bins | | Check: in 2 bins |
| SKU-200 | 0 | | | Not on map | | Check: not on map |
| 1001 (number) | 1 | 3 | 3 | C-Bin 3 | 303 | OK |
| "1001" (text) | 0 | | | Not on map | | Check: space, or number vs text |
J4:J11: 11, 7, 3, 1, 4, 0, Check, Check.

Edge cases tested (all matched a Python dictionary model with LibreOffice equality rules)
- All 48 SKUs on a clean map (Map CHECK OK) plus a missing SKU and a blank; 12 random 6 x 8 maps with 0 to 6 empty bins (4 with a duplicate SKU), 50 lookups each; a pre-sorted pick list (walk order 101 up to 608).
- Blank and "  " lookups: blank row (without the guard the probe counted 1 match at D-Bin 4, the empty bin).
- Trailing space on the map ("SKU-120 "): L7 1; "SKU-120" Check: space, or number vs text; "SKU-120 " found at C-Bin 5.
- "SKU-1*" and "SKU-14?" on the map: SUMPRODUCT finds them exactly (A-Bin 1, F-Bin 8), "SKU-1~*" not found; but Map L6 counted those 2 bins as repeated (COUNTIF reads * and ? as wildcards) and Map CHECK shows Check.
- The number 0 typed as a lookup matched the empty bin (Times 1, D-Bin 4, OK: an empty cell equals 0 in LibreOffice); the text "0" is Not on map. Advise SKUs as text.
- Numbers vs text both ways (2002 number on the map: 2002 found, "2002" Check; "2003" text on the map: "2003" found, 2003 Check; 1.5 found). Word and number labels: Top-1, Closet-8. Empty map (all Not on map, L5 48). Row on 55 (CHECK Check).
- Probe (unguarded, sample map): CAN-LAV-08 row 7 and column 7, INDEX #REF!; SKU-200 0 and 0, #VALUE!; text "1001": SUMPRODUCT 0 but COUNTIF 1; "MUG-BLU-12 ": TRIM version 1; "SKU-1*": COUNTIF 43, SUMPRODUCT 0.

Pitfalls the post must state
- The map must be one block, one SKU per cell; keep SKUs as text (link keep-leading-zeros-zip-codes for Plain text). Case-insensitive matching and the COUNTIF wildcard result are LibreOffice results; link the SUMPRODUCT help pages for Excel and Sheets. No warehouse or inventory system claims.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- In 2 bins: the SKU is on the map twice (unguarded: row 7, column 7 and #REF!).
- Not on map with Check: space, or number vs text: a trailing space or a number stored as text on one side.
- A blank lookup cell or the number 0 points at the empty bin (unguarded blank; 0 even with the guard).
- Map CHECK says Check for a SKU with * or ?: the COUNTIF repeat check reads them as wildcards; the lookup itself is exact.

Hero: sheet "Find", rows 4 to 15, hide_cols "HIJ" (summary); G Status can stay (shows OK and the red Check rows) or be hidden with "GHIJ". Map A4:I10 is a good inset (sheet "Map", hide_cols "JKL"; CAN-LAV-08 bins orange).

---
## 5. turn-product-titles-into-url-handles-slugs-lower-substitute-excel-google-sheets
Title (from topics): How to Turn Product Titles Into URL Handles (Slugs) in Excel and Google Sheets
Template: templates/tidy-tabs-url-handle-from-title.xlsx (tabs: Titles, How to use)

Layout
- Titles: A1:A2 notes. Header row 4, rows 5:204 (12 sample titles): A Product title (input), B Handle, C Leftover characters, D Times this handle appears, E Unique handle (adds -2, -3), F Length, G Status (OK, Duplicate (numbered) orange, Fix: special letters, Check: no letters or digits, Check: numbered handle repeats), H Unique handle repeats, I Quick formula (wrong, comparison). Summary K4:L11: Titles listed (L4), OK (L5), Duplicate (numbered) (L6), Fix: special letters (L7), Rows showing Check (L8), Longest handle (L9), Typed below row 204 (L10), CHECK (L11). Freeze panes at B5. Landscape. The How to use tab lists the replaced characters.
- Deviations: C is blank (not 0) on a blank title; F is blank on a blank title; H and I added; two extra Status words; CHECK shows Check on the sample by design (2 Fix rows).

Key formulas (exact text)
- B5 `=IF(TRIM(A5)="","",SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(LOWER(A5),"&"," and "),"'",""),"’",""),"-"," "),"/"," "),","," "),"."," "),":"," "),";"," "),"!"," "),"?"," "),"("," "),")"," "),""""," "),"#"," "),"%"," "),"+"," "),"_"," "),"|"," "))," ","-"))` (19 SUBSTITUTE inside TRIM: & -> " and ", ' and ’ removed, - / , . : ; ! ? ( ) " # % + _ | -> space; core `=SUBSTITUTE(TRIM(cleaned)," ","-")`)
- C5 `=IF(B5="","",SUMPRODUCT(--ISERROR(FIND(MID(B5,ROW(INDIRECT("1:"&LEN(B5))),1),"abcdefghijklmnopqrstuvwxyz0123456789-"))))`
- D5 `=IF(B5="","",COUNTIF($B$5:$B$204,B5))`
- E5 `=IF(B5="","",IF(COUNTIF($B$5:B5,B5)>1,B5&"-"&COUNTIF($B$5:B5,B5),B5))`
- F5 `=IF(E5="","",LEN(E5))`; H5 `=IF(E5="","",COUNTIF($E$5:$E$204,E5))`
- G5 `=IF(TRIM(A5)="","",IF(B5="","Check: no letters or digits",IF(C5>0,"Fix: special letters",IF(H5>1,"Check: numbered handle repeats",IF(D5>1,"Duplicate (numbered)","OK")))))`
- I5 `=IF(TRIM(A5)="","",SUBSTITUTE(LOWER(TRIM(A5))," ","-"))`
- L11 `=IF(AND(L4>0,L7=0,L8=0,L10=0),"OK","Check")`

Sample values (fictional; recalculated)
| Title | Unique handle (E) | Leftover | Len | Status | Quick formula (I) |
|---|---|---|---|---|---|
| Lavender & Honey Soy Candle - 8 oz | lavender-and-honey-soy-candle-8-oz | 0 | 34 | Duplicate (numbered) | lavender-&-honey-soy-candle---8-oz |
| Grandma's Apple Butter (16 oz.) | grandmas-apple-butter-16-oz | 0 | 27 | OK | grandma's-apple-butter-(16-oz.) |
| "  Hand-Painted  Mug, Blue/White " | hand-painted-mug-blue-white | 0 | 27 | OK | hand-painted-mug,-blue/white |
| Holiday Gift Box: 3 Jams! | holiday-gift-box-3-jams | 0 | 23 | OK | holiday-gift-box:-3-jams! |
| 100% Cotton Tote #2 | 100-cotton-tote-2 | 0 | 17 | OK | 100%-cotton-tote-#2 |
| Grandma’s Fig Jam (curly) | grandmas-fig-jam | 0 | 16 | OK | grandma’s-fig-jam |
| Pumpkin_Spice+Clove \| Wax Melt | pumpkin-spice-clove-wax-melt | 0 | 28 | OK | pumpkin_spice+clove-\|-wax-melt |
| "Best Seller" Tea Towel? | best-seller-tea-towel | 0 | 21 | OK | "best-seller"-tea-towel? |
| Crème Brûlée Candle | crème-brûlée-candle | 3 | 19 | Fix: special letters | crème-brûlée-candle |
| Lavender & Honey Soy Candle – 8 oz (en dash) | lavender-and-honey-soy-candle-–-8-oz | 1 | 36 | Fix: special letters | lavender-&-honey-soy-candle-–-8-oz |
| Lavender and Honey Soy Candle 8 oz | lavender-and-honey-soy-candle-8-oz-2 | 0 | 36 | Duplicate (numbered) | lavender-and-honey-soy-candle-8-oz |
| Lavender and Honey Soy Candle 8 oz | lavender-and-honey-soy-candle-8-oz-3 | 0 | 36 | Duplicate (numbered) | lavender-and-honey-soy-candle-8-oz |
D is 3 on the three lavender-and-honey rows. L4:L11: 12, 7, 3, 2, 0, 36, 0, Check.

Edge cases tested (all matched a character-by-character Python slug model)
- Blank and "   ": blank row. "!!!" and "'''": Check: no letters or digits. "- Sale -" -> sale. "&" -> and. "2026" (text) and 2026 (number) -> 2026 and 2026-2. 12.5 (number) -> 12-5. 200 x's -> 200 characters; a 224-character title. "A--B" and "a - b" -> a-b, a-b-2. "Mug (Blue)", "MUG blue", "mug-blue" -> mug-blue, mug-blue-2, mug-blue-3. "1/2 Pint Jar" and "1-2 pint jar" -> 1-2-pint-jar, -2. "50% Off!" -> 50-off. "Price $5" -> price-$5 (Fix).
- "-2" collision: "Tote", "Tote", "Tote 2" -> tote, tote-2, tote-2: both Check: numbered handle repeats; a later second "Tote 2" -> tote-2-2.
- Every printable ASCII character between letters and at the start of a title: 14 are left over (Fix): $ * < = > @ [ \ ] ^ ` { } ~. é, É (LOWER makes it é), ü, ñ, en dash, em dash, ©, curly double quotes and a non-breaking space are all Fix (TRIM does not remove a non-breaking space); the curly apostrophe ’ is removed. 200 random titles, 10 random sets (4 with every character), a row on 205, empty tab.
- LibreOffice COUNTIF on Fix rows only: "a~b" counted 0 (~ is an escape) and "<3-mug" counted 195 (read as "less than"); these rows already show Fix: special letters.
- Probe: the ROW($A$1:INDEX($A:$A,LEN(B5))) version gave the same leftover counts as ROW(INDIRECT()) on the sample.

Pitfalls the post must state
- No claim about any store platform's handle rules, length limits or the SEO effect of URLs; "check your platform's import help". Accented letters are flagged, not converted. No CSV export claims (none tested). The full B5 formula appears once in a code block.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Handles with & or triple hyphens: the quick formula (column I: lavender-&-honey-soy-candle---8-oz).
- Fix: special letters: accented letters, an en dash, $, @, a non-breaking space; retype them in the title.
- Two products with the same handle: E numbers them -2, -3; Check: numbered handle repeats when a numbered handle equals another title's handle (tote-2).
- Empty handle: a title with no letters or digits (Check: no letters or digits).

Hero: sheet "Titles", rows 4 to 16, hide_cols "DFHJKL" (show A title, B handle, C leftover, E unique handle, G status, I quick formula for the before/after); for a cleaner image also hide "I" ("DFHIJKL"). Delete rows 17 to 204 in the render copy first (see note at the top).
