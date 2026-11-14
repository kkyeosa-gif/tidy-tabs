# Handoff 2026-11-15: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-15.md (researcher) and tasks/hooks-2026-11-15.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`CHECK_ORANGE`, `reverse_list_order`, `size_color_combinations`, `next_date_on_schedule`, `plan_backward_from_event`, `reviews_needed_for_rating`, registered after `payment_plan_interest_rate()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-15 build and test"). Heroes are not made yet (designer, see the hero section below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 129 edited workbooks recalculated and compared (68,423 cell comparisons, 0 mismatches against independent Python models), plus 1 probe workbook of the short (core) formulas and failure cases read by hand; no error value in any recalculated workbook. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text. Bugs found and fixed are in tasks/briefs.md.

Repo note: only the five new builder functions were run against templates/ (imported and run with /usr/bin/python3 -B), so no older template was rewritten and __pycache__ was not touched. The whole script was also run once from a scratch directory (exit 0, all files written there, including the five new ones), so a full re-run still works. `git status` shows scripts/build-templates.py, the five new .xlsx files, this handoff and tasks/briefs.md.

Post file names: `posts/ready/2026-11-15-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks Check and the other flagged results named below, green marks Upcoming / At or above target, orange marks Blank row. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, MINIFS/MAXIFS, dynamic arrays (none used; WORKDAY and EOMONTH are plain, no _xlfn in any file). Help links were not checked by the builder; the script-writer checks every link as usual. The hook openings in tasks/hooks-2026-11-15.md were rechecked against the built files: every cell address and value they quote matches.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-15-team-01-reverse-list-order-newest-to-oldest-first-index-rows-excel-google-sheets | tidy-tabs-reverse-list-order.xlsx | 25 workbooks, 25,275 comparisons |
| 2 | 2026-11-15-team-02-list-all-size-and-color-variant-combinations-index-int-mod-excel-google-sheets | tidy-tabs-size-color-combinations.xlsx | 27 workbooks, 32,616 comparisons |
| 3 | 2026-11-15-team-03-find-next-scheduled-date-after-a-date-small-countif-excel-google-sheets | tidy-tabs-next-date-on-schedule.xlsx | 25 workbooks, 5,650 comparisons |
| 4 | 2026-11-15-team-04-plan-backward-from-event-date-task-start-dates-workday-excel-google-sheets | tidy-tabs-plan-backward-from-event.xlsx | 34 workbooks, 3,298 comparisons |
| 5 | 2026-11-15-team-05-five-star-reviews-needed-to-reach-target-rating-roundup-excel-google-sheets | tidy-tabs-reviews-needed-for-rating.xlsx | 18 workbooks, 1,584 comparisons |

tested_in text pattern (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets."

Hero files: for the designer (not made by the builder). Key columns for a hero (real LibreOffice PDF render of the template with its sample data, 1600x900):
1. Reverse: Flipped tab A Date, B Description, C Amount, D Running balance, E Status (rows 4 to 12) with G1:H2 (Rows found vs COUNTA); a side-by-side with Export A:C (newest first) shows the flip best. Summary G4:H12 can be hidden or cropped.
2. Size-color: Variants tab A No., B Size, C Color, D Variant name, E Price, G Status (rows 5 to 17, all 12 rows) with B1:B3; F On hand and the summary I5:J9 can be hidden. Options tab A:D is a good small inset.
3. Next date: Schedule tab A Date, B Event, E Status (unsorted list) beside G No., I Date, J Event, K Days until (rows 4 to 9) and B1 As of date; C (ZIP), D Sort key and H Key can be hidden; N4:N5 Last one before is optional.
4. Plan: Plan tab B1 Event date, A Task, B Work days needed, C Start, D Finish, E Status (rows 4 to 9) and H4 Start by. The other summary cells can be hidden.
5. Reviews: Calculator tab A Shop or product, B Current average, C Number of reviews, D Target average, E 5-star reviews needed, G 1-star reviews before it drops below target, H Status (rows 4 to 11); F Average after those is optional; the summary J5:K9 can be hidden.

Every workbook has a "How to use" tab and a US Letter page setup with the header row repeated (data tabs landscape, fit to one page wide). Deviations from tasks/team-topics-2026-11-15.md are listed per section.

---
## 1. reverse-list-order-newest-to-oldest-first-index-rows-excel-google-sheets
Title (from topics): How to Reverse a List So the Oldest Row Comes First in Excel and Google Sheets
Template: templates/tidy-tabs-reverse-list-order.xlsx (tabs: Export, Flipped, How to use)

Layout
- Export: A1 note. Header row 4, rows 5:204 (8 sample rows, newest first). A Date, B Description, C Amount (inputs, yellow). Landscape.
- Flipped: A1 note; G1 Rows found (last row with a date) H1; G2 Filled Date cells (COUNTA, for comparison) H2. Header row 4, rows 5:204 (all formulas): A Date, B Description, C Amount, D Running balance, E Status (OK, Blank row orange, Check red). Summary G4:H12: Rows in Export (any cell filled) (H4), Rows flipped (H5), Blank rows inside (H6), Text dates in Export (H7), Rows showing Check (H8), Rows below row 204 (not read) (H9), Total in Export (H10), Total flipped (H11), CHECK (H12). Landscape.
- Deviations: no opening balance cell (the sample's Opening deposit row is the start); the CHECK also needs every filled export row flipped, no Check rows and nothing below row 204 (a text amount is skipped by SUM on both tabs, so the totals alone would not catch it); E also says Check for a text amount; summary has Rows below row 204.

Key formulas (row 5)
- H1 `=IFERROR(MATCH(9.99E+307,Export!$A$5:$A$204),0)` (core `=MATCH(9.99E+307,Export!$A$5:$A$204)`); H2 `=COUNTA(Export!$A$5:$A$204)`
- A5 (copied to B:C and down) `=IF(ROWS($A$5:A5)>$H$1,"",IF(INDEX(Export!A$5:A$204,$H$1-ROWS($A$5:A5)+1)="","",INDEX(Export!A$5:A$204,$H$1-ROWS($A$5:A5)+1)))` (core `=INDEX(Export!A$5:A$204,$H$1-ROWS($A$5:A5)+1)`)
- D5 `=IF(OR(E5="",E5="Blank row"),"",SUM($C$5:C5))` (core `=SUM($C$5:C5)`)
- E5 `=IF(ROWS($A$5:A5)>$H$1,"",IF(AND(A5="",B5="",C5=""),"Blank row",IF(OR(NOT(ISNUMBER(A5)),NOT(ISNUMBER(C5))),"Check","OK")))`
- H4 `=SUMPRODUCT(--((Export!A5:A204<>"")+(Export!B5:B204<>"")+(Export!C5:C204<>"")>0))`; H5 `=COUNTIF(E5:E204,"OK")+COUNTIF(E5:E204,"Check")`; H6 `=COUNTIF(E5:E204,"Blank row")`; H7 `=SUMPRODUCT(--ISTEXT(Export!A5:A204))`; H8 `=COUNTIF(E5:E204,"Check")`; H9 `=COUNTA(Export!A205:C1000)`; H10 `=SUM(Export!C5:C204)`; H11 `=SUM(C5:C204)`; H12 `=IF(AND(H5=H4,H8=0,H9=0,ROUND(H10-H11,2)=0),"OK","Check")`

Sample values (fictional; recalculated). Export newest first: 11/12 Etsy payout $412.50; 11/10 Supplies -$86.40; 11/10 Square payout $98.75; 11/06 Shipping labels -$42.18; 11/05 Stripe payout $1,150.00; 11/03 Booth fee -$75.00; 11/02 Etsy payout $199.99; 11/01 Opening deposit $500.00.
| Date | Description | Amount | Running balance | Status |
|---|---|---|---|---|
| 11/01/2026 | Opening deposit | $500.00 | $500.00 | OK |
| 11/02/2026 | Etsy payout | $199.99 | $699.99 | OK |
| 11/03/2026 | Booth fee | -$75.00 | $624.99 | OK |
| 11/05/2026 | Stripe payout | $1,150.00 | $1,774.99 | OK |
| 11/06/2026 | Shipping labels | -$42.18 | $1,732.81 | OK |
| 11/10/2026 | Square payout | $98.75 | $1,831.56 | OK |
| 11/10/2026 | Supplies | -$86.40 | $1,745.16 | OK |
| 11/12/2026 | Etsy payout | $412.50 | $2,157.66 | OK |
H1 8, H2 8. Summary H4:H12: 8, 8, 0, 0, 0, 0, $2,157.66, $2,157.66, OK. Quotable: the two 11/10/2026 rows come out in exactly the reverse of the export (Supplies above Square payout in the export, Square payout first when flipped), which is the right order for a newest-first export; a Data > Sort by date would put them in whatever order the sort leaves. The post must not say same-day rows "keep" the export order.

Edge cases tested (all matched a Python model: reverse the rows up to the last row with a numeric date)
- Blank row inside (one, and three in a row): H1 9 vs H2 8, every row still flipped, the 11/12/2026 Etsy payout still last, the blank shows Blank row (orange), CHECK OK.
- Last export row with no date (description and amount only): H1 7, that row is not flipped, Total flipped $1,657.66 vs $2,157.66, CHECK Check.
- Text date "11/01/2026" in the last (oldest) row: MATCH skips it, H1 7, Text dates 1, CHECK Check. Text date in the newest row: it is flipped (last row) and shows Check. Text dates in the middle (one and two): flipped with Check.
- Three rows on one day (exactly reversed as a block, never shuffled), 1 row, 0 rows (H1 0, all blank, CHECK OK), a description only (H1 0, Rows in Export 1, CHECK Check), 200 rows, a row on 205 (Rows below row 204 = 3, CHECK Check), "$412.50" as text (Check; both totals $1,745.16 because SUM skips it on both tabs, so CHECK uses the Check count), all negative, a date with a time, $0.00 and a number as description.
- 8 random lists of 1 to 200 rows (same-day runs, blanks, text dates, text and missing amounts).

Pitfalls the post must state
- MATCH(9.99E+307,...) finds the last number in the column; it was checked only in LibreOffice (on a newest-first, descending list). Do not say it works the same in Excel and Sheets; link Microsoft's MATCH help for approximate match if the script-writer checks it.
- COUNTA counts filled cells, not the last row: with a blank row inside, the core INDEX with COUNTA drops the oldest (bottom) export row. Correction to the topic file, which says the newest row is dropped: in the probe the COUNTA version listed 11/02 Etsy payout first and 11/12 Etsy payout last and left out 11/01 Opening deposit $500.00.
- A date typed as text in the bottom row is not found by MATCH, so that row is not flipped; the file catches it (Rows in Export vs Rows flipped, CHECK).
- Without the blank guard a blank row comes through as 0 (12/30/1899 under a date format).
- Do not offer SORT or SORTBY (off the list). Data > Sort can be one line with its warning (it changes the export and can reorder same-day rows). Bank data is fictional.
- Not tested: Excel, Google Sheets (MATCH(9.99E+307) on an unsorted or descending column, INDEX with a ROWS counter, SUMPRODUCT with ISTEXT are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- The oldest row is missing: COUNTA used for the row count with a blank row inside (probe: COUNTA 8, MATCH 9; the 11/01 Opening deposit was left out).
- A 0 or 12/30/1899 row appears: a blank export row read without the blank guard (probe: INDEX returned 0).
- The bottom row is missing: its date is text (probe: MATCH(9.99E+307) on 11/12, 11/10, "11/01/2026" returned 2); the template shows Text dates 1 and CHECK Check.
- The totals agree but CHECK says Check: an amount typed as text ("$412.50"), skipped by SUM on both tabs; the row shows Check.

---
## 2. list-all-size-and-color-variant-combinations-index-int-mod-excel-google-sheets
Title (from topics): How to List Every Size and Color Combination in Excel and Google Sheets
Template: templates/tidy-tabs-size-color-combinations.xlsx (tabs: Options, Variants, How to use)

Layout
- Options: A1 note. Header row 4: A Size (A5:A14), B Price for this size (B5:B14), D Color (D5:D24), inputs. Landscape.
- Variants: A1:B3 Sizes (B1), Colors (B2), Rows (sizes x colors) (B3). Header row 5, rows 6:205 (room for 10 x 20): A No., B Size, C Color, D Variant name, E Price (formulas), F On hand (yellow input), G Status (OK, Blank option, Extra space, Duplicate, No price; not OK in red). Summary I5:J9: Combinations (J5), Rows listed (J6), Rows not OK (J7), Options typed past the yellow rows (J8), CHECK (J9). Landscape.
- Deviations: the header is "No." (not "#"); Price uses the same INDEX position as the size instead of MATCH (so two sizes with the same name keep their own prices and a MATCH name mismatch cannot happen); Status has more cases than Check; the third option (scent) was not built or tested, so the post has no "Going further" formula for it.

Key formulas (row 6)
- B1 `=COUNTA(Options!$A$5:$A$14)`; B2 `=COUNTA(Options!$D$5:$D$24)`; B3 `=B1*B2`
- A6 `=IF(ROWS($B$6:B6)>$B$3,"",ROWS($B$6:B6))` (the row number n)
- B6 `=IF(A6="","",IF(INDEX(Options!$A$5:$A$14,INT((ROWS($B$6:B6)-1)/$B$2)+1)="","",INDEX(Options!$A$5:$A$14,INT((ROWS($B$6:B6)-1)/$B$2)+1)))` (core `=INDEX(sizes,INT((n-1)/colors)+1)`)
- C6 `=IF(A6="","",IF(INDEX(Options!$D$5:$D$24,MOD(ROWS($B$6:B6)-1,$B$2)+1)="","",INDEX(Options!$D$5:$D$24,MOD(ROWS($B$6:B6)-1,$B$2)+1)))` (core `=INDEX(colors,MOD(n-1,colors)+1)`)
- D6 `=IF(A6="","",B6&" / "&C6)`; E6 same as B6 on Options!$B$5:$B$14 (price of that size)
- G6 `=IF(A6="","",IF(OR(B6="",C6=""),"Blank option",IF(OR(TRIM(B6&"")<>B6&"",TRIM(C6&"")<>C6&""),"Extra space",IF(SUMPRODUCT(--($D$6:$D$205=D6))>1,"Duplicate",IF(NOT(ISNUMBER(E6)),"No price","OK")))))`
- J6 `=COUNT(A6:A205)`; J7 `=COUNT(A6:A205)-COUNTIF(G6:G205,"OK")`; J8 `=COUNTA(Options!A15:B40)+COUNTA(Options!D25:D60)`; J9 `=IF(AND(J6=J5,J7=0,J8=0),"OK","Check")`

Sample values (fictional candle shop; recalculated). Sizes 8 oz $14.00, 12 oz $18.00, 16 oz $24.00; colors Natural, Black, Amber, Sage. B1 3, B2 4, B3 12.
| No. | Size | Color | Variant name | Price | Status |
|---|---|---|---|---|---|
| 1 | 8 oz | Natural | 8 oz / Natural | $14.00 | OK |
| 2 | 8 oz | Black | 8 oz / Black | $14.00 | OK |
| 3 | 8 oz | Amber | 8 oz / Amber | $14.00 | OK |
| 4 | 8 oz | Sage | 8 oz / Sage | $14.00 | OK |
| 5 | 12 oz | Natural | 12 oz / Natural | $18.00 | OK |
| 6 to 8 | 12 oz | Black, Amber, Sage | 12 oz / ... | $18.00 | OK |
| 9 to 12 | 16 oz | Natural, Black, Amber, Sage | 16 oz / ... | $24.00 | OK |
Row 13 onward blank. Summary J5:J9: 12, 12, 0, 0, OK.

Edge cases tested (all matched a Python model pairing sizes and colors by position)
- Blank cell inside the colors (Natural, blank, Amber, Sage): B2 3, 9 rows, the blank shows as an empty color with Blank option, Sage is dropped, CHECK Check. Blank inside the sizes: same on the size side.
- "Black " with a trailing space and "8  oz" with a double space: Extra space. "Black" twice and "Black" vs "black": Duplicate (LibreOffice's = ignores case). 12 oz with no price: No price on its 4 rows; "$18.00" as text: No price.
- 1 x 1, 0 colors (B3 0, all blank, no error), 0 sizes, 10 x 20 = 200 rows exactly (last row 22 oz / Color 20), 11 sizes (the 11th on row 15 is not read: J8 2, CHECK Check), a color on row 25 (J8 1, Check).
- Numbers as sizes 8, 12, 16 (names "8 / Black"), decimals 8.5 and 0.75, "8 oz" vs "8oz" (two different sizes), the same size twice with two prices (each keeps its own price, Duplicate status).
- 8 random option sets (1 to 10 sizes, 1 to 20 colors, number sizes, missing prices, trailing spaces, duplicates, blanks).

Pitfalls the post must state
- Keep each list together: a blank cell inside makes COUNTA count one less, so the last option drops off and the blank shows up as an empty variant.
- On hand is typed, so it does not move if the list changes: add new colors at the bottom, or paste values before counting.
- A trailing space makes a variant that looks like a duplicate; the file shows Extra space.
- No marketplace variant limits or import formats; SKU building stays out (dropped 11-13).
- Not tested: Excel, Google Sheets (INDEX with INT and MOD on a ROWS counter, COUNTA on typed lists, = on text ignoring case are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- The last color is missing and one variant has no color: a blank cell inside the color list (probe: COUNTA 3, the blank showed as 0, Sage dropped).
- #REF! at the bottom of the list: the core formula copied past sizes x colors rows (probe: rows 10 to 12 returned #REF! with 3 sizes and a count of 3 colors); the file blanks rows past B3.
- Two variants look the same: "Black " with a trailing space, or the same color typed twice.
- A size typed below row 14 never shows: it is past the yellow rows (J8, CHECK Check).

---
## 3. find-next-scheduled-date-after-a-date-small-countif-excel-google-sheets
Title (from topics): How to Find the Next Date on a Schedule in Excel and Google Sheets
Template: templates/tidy-tabs-next-date-on-schedule.xlsx (tabs: Schedule, How to use)

Layout
- Schedule: A1 As of date, B1 11/14/2026 (input), C1 note. Header row 4, rows 5:104 (7 sample events, unsorted): A Date, B Event, C Location (ZIP, text) (inputs); D Sort key (helper), E Status (Past, Upcoming green, Check red). Next block header G4:K4, rows 5:9: G No. (1 to 5), H Key, I Date, J Event, K Days until (I green when 0 days). Summary M4:N9: Last one before (N4), Its event (N5), Events this month (N6), Events on or after the as-of date (N7), Text dates (not counted) (N8), Rows showing Check (N9). Landscape.
- Deviations: the helper is `INT(A5)+ROW()/100000`, not `A5+ROW()/100000` (with the topic's version an 11:59:59 PM event on row 104 moved to the next day in the probe); a blank or text as-of date gives None scheduled and Status Check (the topic expected blank to act as 0); the summary has "Events on or after the as-of date" and "Its event"; header "No." instead of "#".

Key formulas (row 5)
- D5 `=IF(ISNUMBER(A5),INT(A5)+ROW()/100000,"")`
- E5 `=IF(AND(A5="",TRIM(B5)="",C5=""),"",IF(OR(NOT(ISNUMBER(A5)),NOT(ISNUMBER($B$1))),"Check",IF(INT(A5)<INT($B$1),"Past","Upcoming")))`
- H5 `=IF(ISNUMBER($B$1),IFERROR(SMALL($D$5:$D$104,COUNTIF($D$5:$D$104,"<"&INT($B$1))+G5),""),"")` (core `=SMALL(A5:A104,COUNTIF(A5:A104,"<"&B1)+1)`)
- I5 `=IF(H5="","None scheduled",INT(H5))`; K5 `=IF(H5="","",I5-INT($B$1))`
- J5 `=IF(H5="","",IF(INDEX($B$5:$B$104,MATCH(H5,$D$5:$D$104,0))="","",INDEX($B$5:$B$104,MATCH(H5,$D$5:$D$104,0))))` (core `=INDEX(B5:B104,MATCH(date,A5:A104,0))`, which repeats the first name on a tie)
- N4 `=IF(ISNUMBER($B$1),IFERROR(LARGE($A$5:$A$104,COUNTIF($A$5:$A$104,">="&INT($B$1))+1),"None"),"")` (core `=LARGE(A5:A104,COUNTIF(A5:A104,">="&B1)+1)`)
- N5 `=IF(ISNUMBER($B$1),IFERROR(INDEX($B$5:$B$104,MATCH(LARGE($D$5:$D$104,COUNTIF($D$5:$D$104,">="&INT($B$1))+1),$D$5:$D$104,0))&"",""),"")`
- N6 `=IF(ISNUMBER($B$1),COUNTIF($A$5:$A$104,">="&DATE(YEAR($B$1),MONTH($B$1),1))-COUNTIF($A$5:$A$104,">="&(EOMONTH($B$1,0)+1)),"")`; N7 `=IF(ISNUMBER($B$1),COUNTIF($D$5:$D$104,">="&INT($B$1)),"")`; N8 `=SUMPRODUCT(--ISTEXT($A$5:$A$104))`; N9 `=COUNTIF($E$5:$E$104,"Check")`

Sample values (fictional; as of 11/14/2026; recalculated). List in row order: 11/28 Small Biz Saturday 02108, 11/07 Fall craft show 07030, 12/05 Craft fair 01060, 11/14 Farmers market 02108, 10/24 Harvest market 06510, 11/28 Pop-up 07030, 11/21 Holiday market 02139.
| No. | Date | Event | Days until |
|---|---|---|---|
| 1 | 11/14/2026 | Farmers market | 0 |
| 2 | 11/21/2026 | Holiday market | 7 |
| 3 | 11/28/2026 | Small Biz Saturday | 14 |
| 4 | 11/28/2026 | Pop-up | 14 |
| 5 | 12/05/2026 | Craft fair | 21 |
Summary N4:N9: 11/07/2026, Fall craft show, 5, 5, 0, 0. Status: 5 Upcoming, 2 Past. (A 6th next date would show None scheduled; the file has 5 slots.)

Edge cases tested (all matched a Python model: sorted day + row keys)
- As-of 11/14/2026 9:00 AM: same result as the sample (11/14 Farmers market, 0 days). An event at 11/14/2026 6:00 PM: listed first, I shows 11/14/2026, 0 days. An 11:59:59 PM event on 11/14 on row 104: listed on 11/14 (second, after the Farmers market on row 8).
- Three events on 11/28: all three names show in row order. All past (as of 01/15/2027): five None scheduled, last one before 12/05/2026 Craft fair. Before all (01/01/2026). As-of on the last event day (12/05): Craft fair 0 days.
- As-of blank, "abc" and "11/14/2026" as text: None scheduled, every Status Check (N9 7). A text date in the list: skipped, Text dates 1, its Status Check. Blank rows and a name with no date (Check). An event with no name: J blank. December to January (as of 12/28/2026): 12/31/2026 3 days, 01/02/2027 5, 01/09/2027 12. 0 events, 100 events, 8 random schedules (times, text dates, blanks, as-of with times).

Pitfalls the post must state
- No TODAY in the template, so the results stay fixed and checkable; a reader may type =TODAY() in B1 in their own copy and the results then change every day.
- INT on the as-of date: with a time in B1 and no INT, today's event is skipped (probe: 11/21/2026 instead of 11/14/2026).
- "<" includes events on the as-of date; "<=" lists only later ones (probe: 11/21/2026).
- Same-day events need the helper; without it MATCH repeats the first name (probe).
- The ROW()/100000 tie-break works up to 99,999 rows (the file has 100).
- A date typed as text is skipped by SMALL and COUNTIF; the file counts it. LibreOffice converted a text as-of date "11/14/2026" inside INT and COUNTIF in the probe, so the file now requires a real date in B1.
- No MINIFS for a next date per market name (off the list); say a per-market next date needs a separate list. Fictional events and ZIPs.
- Not tested: Excel, Google Sheets (COUNTIF with "<"&INT(), SMALL on a helper with decimals, text-to-date conversion in INT are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook)
- Two events on one day show the same name: core INDEX/MATCH on the date returned Small Biz Saturday for both 11/28 rows.
- Today's market is skipped: B1 = 11/14/2026 9:00 AM with no INT returned 11/21/2026.
- The next date is one day late for a late-evening event: helper `A+ROW()/100000` on 11/14/2026 11:59:59 PM in row 104 gave a key whose INT is 11/15/2026; INT(A) first keeps 11/14.
- An event never shows: its date is text (template copy: Text dates 1, Status Check).

---
## 4. plan-backward-from-event-date-task-start-dates-workday-excel-google-sheets
Title (from topics): How to Plan Backward From an Event Date With WORKDAY in Excel and Google Sheets
Template: templates/tidy-tabs-plan-backward-from-event.xlsx (tabs: Plan, Holidays, How to use)

Layout
- Plan: A1 Event date, B1 12/05/2026; A2 Buffer work days before the event, B2 0 (inputs); C1 note. Header row 4, rows 5:34 (5 sample tasks): A Task (in the order you will do them), B Work days needed (inputs); C Start, D Finish, E Status (OK, Check, Check: blank row; red). Summary G4:H10: Start by (first task's start) (H4), Work days in the plan (H5), Tasks listed (H6), Rows showing Check (H7), Last finish (H8), Day off typed as text (Holidays) (H9), CHECK (H10). Landscape.
- Holidays: header row 1, A Day off (your own), B Note (inputs, rows 2:20); D1 note. Sample 11/26/2026, 11/27/2026, 12/25/2026 "Shop closed" (fictional shop days off, not an official list).
- Deviations: every WORKDAY is wrapped in IFERROR(...,"Check") (a text day off that is not a date made WORKDAY return #VALUE!); text or blank day counts give Check in C instead of an error; B2 blank counts as 0, B2 text, negative or not whole gives Check; summary adds Last finish and the text day-off count. No WORKDAY.INTL and no calendar-day variant were built or tested, so the post has no "Going further" formulas for them.

Key formulas (row 5)
- D5 `=IF(A5="","",IF(A6="",IF(AND(ISNUMBER($B$1),OR($B$2="",AND(ISNUMBER($B$2),$B$2>=0))),IF($B$2=INT($B$2),IFERROR(WORKDAY($B$1,-1-$B$2,Holidays!$A$2:$A$20),"Check"),"Check"),"Check"),IF(ISNUMBER(C6),IFERROR(WORKDAY(C6,-1,Holidays!$A$2:$A$20),"Check"),"Check")))` (cores: last task `=WORKDAY(event,-1,holidays)`, task before `=WORKDAY(next start,-1,holidays)`)
- C5 `=IF(A5="","",IF(AND(ISNUMBER(B5),ISNUMBER(D5)),IFERROR(WORKDAY(D5,-(B5-1),Holidays!$A$2:$A$20),"Check"),"Check"))` (core `=WORKDAY(finish,-(days-1),holidays)`)
- E5 `=IF(A5="","",IF(NOT(ISNUMBER(B5)),"Check",IF(OR(B5<1,B5<>INT(B5)),"Check",IF(AND(A6="",COUNTA(A6:$A$35)>0),"Check: blank row",IF(OR(NOT(ISNUMBER(C5)),NOT(ISNUMBER(D5))),"Check",IF(C5>D5,"Check","OK"))))))`
- H4 `=IF(ISNUMBER(C5),C5,"")`; H5 `=SUM(B5:B34)`; H6 `=SUMPRODUCT(--(E5:E34<>""))`; H7 `=H6-COUNTIF(E5:E34,"OK")`; H8 `=IF(COUNT(D5:D34)=0,"",MAX(D5:D34))`; H9 `=SUMPRODUCT(--ISTEXT(Holidays!A2:A20))`; H10 `=IF(AND(H6>0,H7=0,H9=0,ISNUMBER(H8)),IF(H8<$B$1,"OK","Check"),"Check")`

Sample values (fictional; event Saturday 12/05/2026, buffer 0, days off 11/26, 11/27, 12/25; recalculated)
| Task | Work days | Start | Finish | Status |
|---|---|---|---|---|
| Order supplies | 2 | 11/16/2026 | 11/17/2026 | OK |
| Pour candles | 5 | 11/18/2026 | 11/24/2026 | OK |
| Cure and label | 3 | 11/25/2026 | 12/01/2026 | OK |
| Photograph and price | 2 | 12/02/2026 | 12/03/2026 | OK |
| Pack the booth | 1 | 12/04/2026 | 12/04/2026 | OK |
Summary H4:H10: 11/16/2026, 13, 5, 0, 12/04/2026, 0, OK. Quotable: Cure and label runs Wed 11/25, Mon 11/30, Tue 12/01 (Thanksgiving Thursday, the Friday day off and the weekend are skipped).
Variations (same tasks): event Wed 12/09/2026: Start by 11/18/2026, Pack the booth 12/08; event Sun 12/06/2026: same as Saturday (Start by 11/16/2026); event on the day off 12/25/2026: Pack the booth 12/24, Start by 12/08/2026; buffer 3: Pack the booth 12/01, Start by 11/11/2026; event 01/08/2027 with 12/24, 12/25, 12/31, 01/01 off: Start by 12/16/2026, Cure and label 12/29/2026 to 01/04/2027.

Edge cases tested (all matched a Python model stepping back over weekdays and a day-off set)
- Pour candles 0 days: start 11/25/2026 after finish 11/24/2026, Check, Start by 11/23/2026, CHECK Check.
- Days -2, "3" as text, blank (Check; C shows Check and every task above shows Check), 2.5 (Check; WORKDAY drops the decimal, so it gets 2 days: 12/02 to 12/03 before a 1-day task), a negative alone.
- Blank row after Pour candles: Pour candles is treated as the last task (11/30 to 12/04, "Check: blank row"), the tasks below the blank chain to the event on their own, CHECK Check.
- Event weekday, Saturday, Sunday, on a day off, with a time (10:00 AM, same dates), blank or text (Check). Days off on a weekend (no change), none, one with a time, "11/27/2026" as text (LibreOffice used it as a day off: same dates as the sample, H9 1, CHECK Check), "closed" as text (every row Check, no error). Buffer 0, 3, blank (0), "2" as text, -1, 1.5 (Check). 1-day tasks, 30 tasks, 0 tasks, a plan across New Year, 8 random plans (1 to 30 tasks, 0 to 19 days off, buffer 0 to 3, some bad day counts).

Pitfalls the post must state
- List tasks in order with no blank row; a blank row makes the task above it the last task.
- Work days must be whole numbers of 1 or more; WORKDAY drops a decimal (probe: WORKDAY(11/24/2026,-1.5) returned 11/23/2026, the same as -1).
- WORKDAY counts work days only. Curing time that runs over weekends is calendar time: say so; no calendar-day formula is in the file (not built or tested).
- Day offs are the reader's own dates typed as real dates, not an official holiday list. WORKDAY.INTL (custom weekends) was not tested: no "Going further" for it.
- WORKDAY(date,0) returns the date itself, even on a weekend (probe: 12/05/2026), which is why the last task uses -1.
- Not project-management advice.
- Not tested: Excel, Google Sheets (negative WORKDAY day counts, a fractional day count, a text day off in the holiday range are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- A task starts after it finishes: its work days are 0 (probe: Pour candles 0 days gave start 11/25/2026 and finish 11/24/2026).
- Two tasks overlap near the event: a blank row in the task list (template copy: "Check: blank row").
- Every row says Check: a day off typed as a word ("closed"); the core WORKDAY returned #VALUE! in the probe.
- A typed day off seems ignored or the CHECK says Check with right dates: a day off typed as text that looks like a date ("11/27/2026") was still used by LibreOffice (probe: WORKDAY from 11/30/2026 back 1 gave 11/25/2026); retype it as a date.
- A 2.5-day task got 2 days: WORKDAY drops the decimal (probe).

---
## 5. five-star-reviews-needed-to-reach-target-rating-roundup-excel-google-sheets
Title (from topics): How to Calculate How Many 5-Star Reviews You Need for a Target Rating in Excel and Google Sheets
Template: templates/tidy-tabs-reviews-needed-for-rating.xlsx (tabs: Calculator, From star counts, How to use)

Layout
- Calculator: A1:A2 notes (A2: the math assumes a plain average of every review). Header row 4, rows 5:24 (6 sample rows plus row 11 from star counts): A Shop or product, B Current average, C Number of reviews, D Target average (inputs; B11 and C11 are formulas from the From star counts tab); E 5-star reviews needed, F Average after those, G 1-star reviews before it drops below target, H Status (Below target, At or above target green, Not possible and Check red). Summary J5:K9: Rows listed (K5), Below target (K6), At or above target (K7), Not possible (target 5.00) (K8), Rows showing Check (K9). Landscape.
- From star counts: header row 4, A Stars 5 to 1, B Number of reviews (inputs, 152, 21, 6, 1, 3); A11:B13 Status, Total reviews, Average.
- Deviations: H is computed first (validity) and E/G read it; a current average equal to the target gives 0 needed even at target 5.00 (the topic's E checks D>=5 first); 0 reviews, a target of 1 or less, a target above 5, an average outside 1 to 5 and a non-whole review count are Check; the star-count tab has a Status (Check for text, negative or non-whole counts, No reviews for all zeros).

Key formulas (row 5)
- H5 `=IF(AND(TRIM(A5)="",B5="",C5="",D5=""),"",IF(AND(ISNUMBER(B5),ISNUMBER(C5),ISNUMBER(D5)),IF(AND(B5>=1,B5<=5,C5>=1,C5=INT(C5),D5>1,D5<=5),IF(B5>=D5,"At or above target",IF(D5>=5,"Not possible","Below target")),"Check"),"Check"))`
- E5 `=IF(OR(H5="",H5="Check"),"",IF(B5>=D5,0,IF(D5>=5,"Not possible",ROUNDUP(ROUND(C5*(D5-B5)/(5-D5),6),0))))` (core `=ROUNDUP(reviews*(target-current)/(5-target),0)`)
- F5 `=IF(ISNUMBER(E5),(B5*C5+5*E5)/(C5+E5),"")`
- G5 `=IF(OR(H5="",H5="Check"),"",IF(B5<D5,"Already below",ROUNDDOWN(ROUND(C5*(B5-D5)/(D5-1),6),0)))` (core `=ROUNDDOWN(reviews*(current-target)/(target-1),0)`)
- Calculator B11 `='From star counts'!B13`, C11 `='From star counts'!B12`
- From star counts B11 `=IF(COUNT(B5:B9)<>COUNTA(B5:B9),"Check",IF(SUMPRODUCT(--(B5:B9<0))+SUMPRODUCT(--(B5:B9<>INT(B5:B9)))>0,"Check",IF(SUM(B5:B9)=0,"No reviews","OK")))`; B12 `=IF(B11="OK",SUM(B5:B9),"")`; B13 `=IF(B11="OK",SUMPRODUCT(A5:A9,B5:B9)/SUM(B5:B9),"")` (core `=SUMPRODUCT(A5:A9,B5:B9)/SUM(B5:B9)`)
- K5 `=SUMPRODUCT(--(H5:H24<>""))`; K6:K9 COUNTIF of each status.

Sample values (fictional; recalculated)
| Shop or product | Current | Reviews | Target | 5-star needed | Average after | 1-star before below | Status |
|---|---|---|---|---|---|---|---|
| Candle shop | 4.72 | 183 | 4.80 | 74 | 4.8006 | Already below | Below target |
| Soap listing | 4.70 | 30 | 4.90 | 60 | 4.9000 | Already below | Below target |
| Tote bag listing | 4.60 | 100 | 4.80 | 100 | 4.8000 | Already below | Below target |
| Mug listing | 4.85 | 40 | 4.80 | 0 | 4.8500 | 0 | At or above target |
| Print shop | 4.95 | 240 | 4.90 | 0 | 4.9500 | 3 | At or above target |
| Jewelry shop | 4.90 | 50 | 5.00 | Not possible | | Already below | Not possible |
| Whole candle shop (from star counts) | 4.7377 | 183 | 4.80 | 57 | 4.8000 | Already below | Below target |
Star counts 152, 21, 6, 1, 3: Status OK, 183 reviews, average 4.7377 (4.73770491803279). Summary K5:K9: 7, 4, 2, 1, 0. Quotable: 73 five-star reviews leave the candle shop at 4.7998 and 74 reach 4.8006 (probe); the print shop at 4.95 can take 3 one-star reviews (4.9012) and a 4th drops it to 4.8852.

Edge cases tested (all matched a Python model with exact fractions; 209 checks that n-1 fails and n passes, and that k one-star reviews pass and k+1 fail)
- Target equal to current (0 and 0), below current (4.90 with 100 to 4.50: 0 needed, 11 one-star), 5.00 (Not possible), 4.99 (900 needed at 100 reviews), 4.995 (300 needed at 300 reviews), above 5, 1 and 0.5 (Check).
- Rounded average: 4.7 vs 4.74 at 120 reviews to 4.80 needs 60 vs 36. Raw whole number: 4.70 with 30 to 4.90 is 60 (raw 60.0000000000003 in the probe).
- 0, -5, 10.5 and "183" as text reviews; 5.2, 0.5 and "4.7" as text averages; blank target or average; a name only (all Check). 1 review at 3.0 to 4.0 (1 needed); 1.00 over 5,000 to 4.90 (195,000); 4.98 over 10,000 to 4.99 (10,000); 5.00 over 9,999 with target 4.50 (1,428 one-star); 1.00 to 1.01 (1 needed).
- Star counts with "21" as text, -1, 21.5 (Check), all zero (No reviews; Calculator row 11 Check), all 5-star (5.0000), large counts. 8 random sets of 20 rows (1 or 2 decimal averages, 1 to 20,000 reviews, targets near the average, some 0 reviews, 5.00 and text).

Pitfalls the post must state
- The math assumes the rating is a plain average of all reviews. The post must not say how Etsy, Google or Amazon compute or round their shown rating, and links nothing it has not read.
- Use star counts when you have them; a rating shown rounded changes the answer (4.7 vs 4.74: 60 vs 36).
- A target of 5.00 is not reachable unless every review is 5 stars; the core formula gives #DIV/0! (probe), the file says Not possible.
- The ROUND(...,6) guard: in LibreOffice ROUNDUP returned 60 for 60.0000000000003 even without it (probe). Excel and Sheets were not checked; present the guard as a safety step, not as a fix for any app.
- No advice on getting reviews and nothing on incentives (only "follow your platform's review policy" if at all). Fictional shops.
- Not tested: Excel, Google Sheets (ROUNDUP/ROUNDDOWN of float quotients near whole numbers, SUMPRODUCT on star counts are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- #DIV/0! for a target of 5.0: `=ROUNDUP(100*(5-4.9)/(5-5),0)` (probe); the file shows Not possible.
- The answer seems one review short: rounding to the nearest instead of up (73.2 needs 74; 73 leaves 4.7998, probe).
- The answer changes when you use star counts: the shown 4.7 is really 4.74 (60 vs 36 at 120 reviews), or 4.72 vs the star-count 4.7377 (74 vs 57 at 183 reviews).
- A row shows Check: 0 reviews, a review count typed as text, an average above 5, or a target of 1 or less.

---
What was NOT run (all five): Microsoft Excel, Google Sheets, any non-US language setting, more rows than the slots (204, 205, 104, 34, 24), a printer (PDF export only), the topic's optional "Going further" items (third option for combinations, WORKDAY.INTL, calendar-day plan). Nothing is claimed about Excel or Sheets.

---
Images 2026-11-15 (designer)

Heroes: `heroes_2026_11_15` in scripts/make-heroes.py, run with `python3 -B scripts/make-heroes.py --only-2026-11-15` (no earlier entry run or changed; on this machine run with /usr/bin/python3, which has pymupdf). Each hero line sits directly after the `> Works in` callout in the draft posts (blank line before and after).

Hero files (1600x900 PNG, checked with PIL, real LibreOffice PDF render of the template with its sample data, no error values or ### in any image). Hero alt texts below are 8 to 15 words and differ from the titles:
1. images/2026-11-15-team-01-reverse-list-order-newest-to-oldest-first-index-rows-excel-google-sheets/reverse-list-newest-to-oldest-template.png
   - caption: "Left, the export with the newest row first; right, the Flipped tab lists the same eight rows oldest first with a running balance (its summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2."
   - alt: "Bank export rows newest first beside the same rows flipped oldest first with running balance"
   - line: `![Bank export rows newest first beside the same rows flipped oldest first with running balance](images/2026-11-15-team-01-reverse-list-order-newest-to-oldest-first-index-rows-excel-google-sheets/reverse-list-newest-to-oldest-template.png) *Left, the export with the newest row first; right, the Flipped tab lists the same eight rows oldest first with a running balance (its summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*`
2. images/2026-11-15-team-02-list-all-size-and-color-variant-combinations-index-int-mod-excel-google-sheets/size-color-variant-combinations-template.png
   - caption: "Three sizes and four colors become twelve variant rows, each with its size's price and an OK status (the On hand column and the summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2."
   - alt: "Twelve candle variants pairing 8, 12 and 16 oz sizes with four colors and prices"
   - line: `![Twelve candle variants pairing 8, 12 and 16 oz sizes with four colors and prices](images/2026-11-15-team-02-list-all-size-and-color-variant-combinations-index-int-mod-excel-google-sheets/size-color-variant-combinations-template.png) *Three sizes and four colors become twelve variant rows, each with its size's price and an OK status (the On hand column and the summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*`
3. images/2026-11-15-team-03-find-next-scheduled-date-after-a-date-small-countif-excel-google-sheets/next-date-on-schedule-template.png
   - caption: "As of 11/14/2026, an unsorted list of seven market dates gives the next five in order, the two on 11/28 each with its own name (ZIP, helper and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2."
   - alt: "Unsorted market dates marked Past or Upcoming beside the next five events and days until"
   - line: `![Unsorted market dates marked Past or Upcoming beside the next five events and days until](images/2026-11-15-team-03-find-next-scheduled-date-after-a-date-small-countif-excel-google-sheets/next-date-on-schedule-template.png) *As of 11/14/2026, an unsorted list of seven market dates gives the next five in order, the two on 11/28 each with its own name (ZIP, helper and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*`
4. images/2026-11-15-team-04-plan-backward-from-event-date-task-start-dates-workday-excel-google-sheets/plan-backward-from-event-date-template.png
   - caption: "Five tasks planned back from a Saturday 12/05/2026 event skip weekends and the shop's 11/26 and 11/27 days off, so the plan starts 11/16/2026. PDF export from LibreOffice Calc 24.2."
   - alt: "Five candle-making tasks with work days, start and finish dates counted back from December 5"
   - line: `![Five candle-making tasks with work days, start and finish dates counted back from December 5](images/2026-11-15-team-04-plan-backward-from-event-date-task-start-dates-workday-excel-google-sheets/plan-backward-from-event-date-template.png) *Five tasks planned back from a Saturday 12/05/2026 event skip weekends and the shop's 11/26 and 11/27 days off, so the plan starts 11/16/2026. PDF export from LibreOffice Calc 24.2.*`
5. images/2026-11-15-team-05-five-star-reviews-needed-to-reach-target-rating-roundup-excel-google-sheets/five-star-reviews-needed-template.png
   - caption: "Seven fictional shops and listings: 74 five-star reviews take the candle shop from 4.72 to 4.80, and a 5.00 target shows Not possible (the Average after column and the summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2."
   - alt: "Shop ratings with current average, review count, target and the 5-star reviews still needed"
   - line: `![Shop ratings with current average, review count, target and the 5-star reviews still needed](images/2026-11-15-team-05-five-star-reviews-needed-to-reach-target-rating-roundup-excel-google-sheets/five-star-reviews-needed-template.png) *Seven fictional shops and listings: 74 five-star reviews take the candle shop from 4.72 to 4.80, and a 5.00 target shows Not possible (the Average after column and the summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*`
Shown in each hero: 1 Export rows 4-12 beside Flipped rows 4-12 (notes rows 1-3, the Rows found cells G1:H2 and the summary G4:H12 hidden with columns F:H); 2 Variants rows 1-17 without the spacer row 4, On hand (F) and the summary (H:J); 3 Schedule rows 1 and 4-11 without the note C1, Location (ZIP) (C), Sort key (D), Key (H) and the summary (L:N); 4 Plan rows 1-2 and 4-10 without the note C1, summary G4:H10 kept; 5 Calculator rows 4-11 without the notes rows 1-3, Average after (F) and the summary (I:K). All hiding is in the hero copy, not in the template files.

AI photo (`image_prompts`, `image_alt`), one scene each, 2 or 3 specific objects, no screens, no readable text, no logos, no laptop and coffee:
1. Reverse: image_prompts "A tall stack of plain paper bank statement pages being turned over on a wooden desk, the top half already lying face down in a shallow wooden in-tray, a large black bulldog clip set aside, all print blurred and unreadable, soft window light, realistic photo, no screens, no readable text, no logos" image_alt "Stack of statement pages being turned face down into a wooden in-tray"
2. Size-color: image_prompts "Twelve empty glass candle vessels on a wooden shelf in a small candle studio, arranged in three rows by size, small, medium and large, and four columns by glass tint, clear, black, amber and sage green, soft daylight, realistic photo, no screens, no readable text, no logos" image_alt "Twelve empty glass candle vessels in rows by size and columns by tint"
3. Next date: image_prompts "A collapsible garden wagon parked by a small shop's back door at dawn, loaded with a folded market table, a rolled canvas tablecloth and a folding stool, soft early light, realistic photo, no screens, no readable text, no logos" image_alt "Garden wagon loaded with a folded market table, rolled tablecloth and stool"
4. Plan: image_prompts "A sheet tray of freshly poured candles in plain glass jars, each wick held upright by a wooden centering bar, on a home studio counter beside a small heat gun, soft window light, realistic photo, no screens, no readable text, no logos" image_alt "Freshly poured candles with wooden wick bars on a sheet tray beside a heat gun"
5. Reviews: image_prompts "An open cloth-bound guest book with blank cream pages on the front table of a craft fair booth, a fountain pen resting in the gutter and a small ceramic vase of dried lavender beside it, soft daylight, realistic photo, no screens, no readable text, no logos" image_alt "Open guest book with blank pages, a fountain pen and dried lavender at a booth"

All ten alt texts (5 hero, 5 photo) were checked against every `![alt]` and `image_alt` in posts/ (ready and published, 462 existing alts): no duplicates. Photo scenes avoid objects already used in earlier prompts (receipt spike, soap bars, parcels and envelopes, wax pitcher, thank-you cards, canopy tent).
