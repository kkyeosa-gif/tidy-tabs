# Handoff 2026-11-17: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-17.md (researcher) and tasks/hooks-2026-11-17.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`smallest_box_that_fits`, `cost_per_usable_unit_after_waste`, `three_point_project_estimate`, `nps_survey_score`, `average_without_highest_lowest`, registered after `correlation_sales_vs_temperature()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-17 build and test"), one fix in tasks/hooks-2026-11-17.md (see section 4). Heroes are not made yet (designer, see the hero section below).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 112 edited workbooks recalculated and compared (52,974 cell comparisons, 0 mismatches against independent Python models), plus 2 probe workbooks of the short (core) formulas, the hook traps and TRIMMEAN on 3 to 100 values, read by hand; no error value in any recalculated template workbook. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text. Bugs found and fixed are in tasks/briefs.md.

Repo note: only the five new builder functions were run against templates/ (imported and run with /usr/bin/python3 -B), so no older template was rewritten and __pycache__ was not touched. The whole script was also run from a scratch directory, /tmp/claude-0/build-1117/full (exit 0, all 221 files written there, including the five new ones), so a full re-run still works. `git status` shows scripts/build-templates.py, the five new .xlsx files, this handoff, tasks/briefs.md and tasks/hooks-2026-11-17.md.

Post file names: `posts/ready/2026-11-17-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks Check (every status starting "Check") and a CHECK cell that is not OK; orange marks No box fits (box picker), Note: smaller than above (Boxes tab) and Skipped: blank / Skipped: text (average); green marks Promoter (NPS Group column). The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, MINIFS/MAXIFS, dynamic arrays (none used; TRIMMEAN, SUMSQ, MEDIAN, SUMIFS and COUNTIFS are plain, no _xlfn in any file). Help links were not checked by the builder; the script-writer checks every link as usual. The hook openings in tasks/hooks-2026-11-17.md were rechecked against the built files: every value they quote matches except opening 4's COUNTA trap, now corrected in the hooks file (section 4).

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-17-team-01-pick-smallest-shipping-box-that-fits-item-dimensions-match-index-excel-google-sheets | tidy-tabs-smallest-box-that-fits.xlsx | 28 workbooks, 14,364 comparisons |
| 2 | 2026-11-17-team-02-true-cost-per-usable-unit-after-waste-percent-excel-google-sheets | tidy-tabs-cost-per-usable-unit-after-waste.xlsx | 15 workbooks, 9,885 comparisons |
| 3 | 2026-11-17-team-03-three-point-estimate-project-hours-quote-range-pert-sumsq-excel-google-sheets | tidy-tabs-three-point-project-estimate.xlsx | 19 workbooks, 2,565 comparisons |
| 4 | 2026-11-17-team-04-net-promoter-score-from-customer-survey-0-10-countifs-excel-google-sheets | tidy-tabs-nps-survey-score.xlsx | 22 workbooks, 22,968 comparisons |
| 5 | 2026-11-17-team-05-average-without-highest-and-lowest-value-trimmean-excel-google-sheets | tidy-tabs-average-without-highest-lowest.xlsx | 28 workbooks, 3,192 comparisons |

tested_in text pattern (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets."

Hero files: for the designer (not made by the builder). Key columns for a hero (real LibreOffice PDF render of the template with its sample data, 1600x900):
1. Box: Orders tab B1 padding, A Order or item, B Length, C Width, D Height, E Status, F Box to use, G Box cost (rows 4 to 11, all 7 items; the Poster tube row shows No box fits in orange). H Box row (helper) and the summary J4:K11 can be hidden. Boxes tab A:E rows 4 to 12 is a good small inset (F:H sorted sides optional).
2. Waste: Materials tab A Material, C Pack cost, D Units per pack, E Waste %, F Cost per bought unit, H Cost per usable unit, I Shortcut, K Packs to buy (rows 4 to 9). B Unit, G, J, L, M and the summary O4:P8 can be hidden; the side by side of F, H and I ($1.20, $1.50, $1.44) is the point. Waste log G4:K7 (measured 18.00%) is an optional inset.
3. PERT: Estimate tab B1 rate, A Task, B O, C M, D P, E Expected, F Spread (rows 4 to 10) beside the summary I8:J15 (expected 45.83, spread 4.14, +1, +2, spreads added 8.50, the three quotes). G Status and I4:J7, I16:J18 can be hidden.
4. NPS: Summary tab A4:H8 (All responses 20 / 11 / 4 / 5 / 55.0% / 25.0% / 30, then Etsy 50, Market 14, Email 20); I (not rounded) and rows 11 to 14 can be hidden. Responses A:F rows 4 to 12 is an optional inset (Promoter cells green).
5. Average: Sales tab A Market date, B Sales, C Note (rows 4 to 14, the $1,480.00 Holiday rush and $128.00 Rained out rows) beside F4:G13 (plain $674.10, median $642.50, without highest and lowest $641.63, TRIMMEAN check $641.63, TRIMMEAN 0.1 $674.10). D Status and F14:G17 can be hidden.

Every workbook has a "How to use" tab and a US Letter page setup with the header row (row 4) repeated (data tabs landscape, fit to one page wide). Deviations from tasks/team-topics-2026-11-17.md are listed per section.

---
## 1. pick-smallest-shipping-box-that-fits-item-dimensions-match-index-excel-google-sheets
Title (from topics): How to Pick the Smallest Shipping Box That Fits Each Item in Excel and Google Sheets
Template: templates/tidy-tabs-smallest-box-that-fits.xlsx (tabs: Boxes, Orders, How to use)

Layout
- Boxes: A1:A2 notes. Header row 4, rows 5:24 (8 sample boxes, smallest first): A Box name, B Inside length, C Inside width, D Inside height (in), E Cost each (inputs); F Longest side, G Middle side, H Shortest side, I Volume (cu in) (helpers), J Status (OK; Note: smaller than above orange; Check: size, Check: cost, Check: no name red). Summary L4:M8: Boxes listed (M4), Boxes the Orders tab can use (M5), Rows showing Check (M6), Smaller than the box above (M7), Typed below row 24 (M8). Landscape.
- Orders: A1 Padding added to each side (in), B1 1 (input; blank counts as 0); C1 note. Header row 4, rows 5:104 (7 sample items): A Order or item, B Length, C Width, D Height (inputs); E Status (OK, Check: size, Check: padding), F Box to use (No box fits orange), G Box cost, H Box row (helper). Summary J4:K11: Items listed (K4), Items with a box (K5), No box fits (K6), Rows showing Check (K7), Total box cost (K8), Boxes tab rows showing Check (K9), Typed below row 104 (K10), CHECK (K11). Landscape.
- Deviations: H Box row helper (Box and Cost read the same row, so two boxes with the same name cannot mix up costs); the topic's `(Boxes!$A$5:$A$24<>"")` guard became `ISNUMBER(Boxes!$I$5:$I$24)` (in LibreOffice a "" compared with a number is TRUE, so a named box with a bad size would match); the middle side and every padded item side are rounded to 6 decimals (a 1 x 10.92 x 10 item has a middle side of 10.000000000000002 in binary; LibreOffice cleaned it, plain float math did not); a name with no sizes is Check: size (topic: blank); Boxes gets a Status and Volume; a box with a text cost is still used (cost blank, CHECK Check).

Key formulas (row 5)
- Boxes F5 `=IF(ISNUMBER(I5),MAX(B5:D5),"")`; G5 `=IF(ISNUMBER(I5),ROUND(SUM(B5:D5)-MAX(B5:D5)-MIN(B5:D5),6),"")`; H5 `=IF(ISNUMBER(I5),MIN(B5:D5),"")`; I5 `=IF(AND(TRIM(A5)<>"",COUNT(B5:D5)=3,MIN(B5:D5)>0),B5*C5*D5,"")`
- Boxes J6 `=IF(AND(TRIM(A6)="",COUNTA(B6:E6)=0),"",IF(TRIM(A6)="","Check: no name",IF(NOT(ISNUMBER(I6)),"Check: size",IF(AND(E6<>"",NOT(ISNUMBER(E6))),"Check: cost",IF(AND(ISNUMBER(I5),I6<I5),"Note: smaller than above","OK")))))` (J5 has no row-above test)
- Orders E5 `=IF(COUNTA(A5:D5)=0,"",IF(NOT(OR($B$1="",AND(ISNUMBER($B$1),$B$1>=0))),"Check: padding",IF(AND(COUNT(B5:D5)=3,MIN(B5:D5)>0),"OK","Check: size")))`
- Orders H5 `=IF(E5<>"OK","",IFERROR(MATCH(1,INDEX((Boxes!$F$5:$F$24>=ROUND(MAX(B5:D5)+N($B$1),6))*(Boxes!$G$5:$G$24>=ROUND(SUM(B5:D5)-MAX(B5:D5)-MIN(B5:D5)+N($B$1),6))*(Boxes!$H$5:$H$24>=ROUND(MIN(B5:D5)+N($B$1),6))*ISNUMBER(Boxes!$I$5:$I$24),0),0),""))` (core `=INDEX(boxes,MATCH(1,INDEX((long>=a)*(mid>=b)*(short>=c),0),0))`)
- Orders F5 `=IF(E5<>"OK","",IF(ISNUMBER(H5),INDEX(Boxes!$A$5:$A$24,H5),"No box fits"))`; G5 `=IF(ISNUMBER(H5),IF(ISNUMBER(INDEX(Boxes!$E$5:$E$24,H5)),INDEX(Boxes!$E$5:$E$24,H5),""),"")`
- K4 `=SUMPRODUCT(--(E5:E104<>""))`; K5 `=COUNT(H5:H104)`; K6 `=COUNTIF(F5:F104,"No box fits")`; K7 `=K4-COUNTIF(E5:E104,"OK")`; K8 `=SUM(G5:G104)`; K9 `=Boxes!M6`; K10 `=COUNTA(A105:D1000)`; K11 `=IF(AND(K4>0,K7=0,K9=0,K10=0,Boxes!M8=0,K5+K6=COUNTIF(E5:E104,"OK")),"OK","Check")`

Sample values (fictional; padding 1 in; recalculated). Boxes in list order: 6x4x4 $0.85, 8x6x4 $1.05, 10x8x4 $1.30, 9x6x6 $1.35, 12x9x6 $1.75, 12x12x8 $2.40, 16x12x8 $2.95, 18x14x10 $3.80 (all OK, volumes 96 to 2,520).
| Order or item | L x W x H | Status | Box to use | Box cost |
|---|---|---|---|---|
| Candle jar | 3 x 3 x 4 | OK | 6x4x4 | $0.85 |
| Mug | 4.5 x 3.5 x 4 | OK | 9x6x6 | $1.35 |
| Flat print | 11 x 8.5 x 0.25 | OK | 12x12x8 | $2.40 |
| Tray (typed on its side) | 5 x 9 x 3 | OK | 10x8x4 | $1.30 |
| Big quilt | 15 x 11 x 7 | OK | 16x12x8 | $2.95 |
| Poster tube | 24 x 3 x 3 | OK | No box fits | |
| Soap bar set | 5 x 3 x 3 | OK | 6x4x4 | $0.85 |
Summary K4:K11: 7, 6, 1, 0, $9.70, 0, 0, OK. Quotable: the Mug padded to 5.5 x 5 x 4.5 is too tall for the three 4 in boxes; the Flat print's padded middle side 9.5 fails 12x9x6; the Soap bar set padded to exactly 6 x 4 x 4 gets 6x4x4 (>= keeps an exact fit).

Edge cases tested (all matched a Python model: sort sides, first box in list order where each padded side fits)
- 5 x 3 x 3 exact fit (6x4x4), 5.01 x 3 x 3 (8x6x4), 3 x 5 x 3 turned (6x4x4), 8.25 x 2 x 2 (10x8x4); 0 x 0 x 0, -2, "8 in" text, a blank side, a name only: Check: size, box blank.
- Padding 0 (Mug 6x4x4, Flat print 12x9x6), 0.5, 2 (Candle jar 9x6x6, Tray 12x9x6, Big quilt 18x14x10), blank (as 0), "1 in" text and -1 (every row Check: padding, CHECK Check).
- Boxes out of size order (18x14x10 first: every item that fits gets 18x14x10, the boxes below show Note: smaller than above), two boxes the same size (the first listed is used), blank row inside the box list (skipped), no name, a text size "12", a 0 side, a text cost "$2.95" (Check: cost; box still used, cost blank), box sides typed in any order, no boxes (all No box fits), no orders, a box on row 25 and an order on row 105 (CHECK Check), 100 orders, 20 boxes x 100 orders, binary leftover exact fits (1 x 10.92 x 10 in a 17x10x6; 18 x 10 x 4.1 in a 21x10x10; 5.9 x 3.9 x 3.9 with 0.1 padding in a 6x4x4), 8 random sets (padding 0 to 2 or blank, 1 to 20 boxes, bad sizes).

Pitfalls the post must state
- Use INSIDE dimensions; straight fits only, no diagonal packing; no carrier size limits, dimensional weight or rates (link dimensional-weight-billable-weight and shipping-cost-weight-tier-lookup); padding is the reader's own choice.
- List boxes in the order you want them tried: the formula returns the first that fits in list order (template copy: a big box listed first takes every item).
- MATCH(1,INDEX(...,0),0) is the categorize-expenses-by-keyword trick; link it, do not re-explain at length.
- Not tested: Excel, Google Sheets (MATCH(1,INDEX(...,0),0) entered as a normal formula, a "" helper compared with a number inside the array test, the binary leftover in SUM-MAX-MIN are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- A bigger box than needed: sides compared as typed (L to L, W to W, H to H) with padding 1: Candle jar 9x6x6, Tray 12x12x8, Mug 9x6x6 (probe; the sorted version gives 6x4x4 and 10x8x4).
- Every item gets the same big box: the box list is not in the order you want them tried (template copy, Note: smaller than above).
- A row shows Check: size: a side is 0, negative, blank or text like "8 in".
- Box cost is blank but a box is named: that box's cost is typed as text (Boxes Status Check: cost).

---
## 2. true-cost-per-usable-unit-after-waste-percent-excel-google-sheets
Title (from topics): How to Find the True Cost per Usable Unit After Waste in Excel and Google Sheets
Template: templates/tidy-tabs-cost-per-usable-unit-after-waste.xlsx (tabs: Materials, Waste log, How to use)

Layout
- Materials: A1:A2 notes. Header row 4, rows 5:54 (5 sample rows): A Material, B Unit, C Pack cost, D Units per pack, E Waste %, J Usable units needed (inputs; E9 is the formula `='Waste log'!K5`); F Cost per bought unit, G Usable units per pack, H Cost per usable unit, I Shortcut: cost x (1 + waste), K Packs to buy, L Status (OK, Check: cost or units, Check: waste), M Usable units from those packs. Summary O4:P8: Materials listed (P4), Rows showing Check (P5), Units needed showing Check (P6), Typed below row 54 (P7), CHECK (P8). Landscape.
- Waste log: A1:A2 notes. Header row 4, rows 5:204 (5 sample rows): A Date, B Material, C Bought (units), D Usable (units) (inputs); E Status (OK, Check: no material, Check: number, Check: usable vs bought). Table G4:L24: G Material (input, 3 names), H Rows logged (OK), I Bought, J Usable, K Measured waste %, L Rows with Check. G26:H29: Log rows showing Check (H26), Log rows whose name is not in G (H27), Typed below row 204 (H28), CHECK (H29). Landscape.
- Deviations: K is **Packs to buy** (whole packs, the open question) plus M Usable units from those packs; the measured waste uses SUMIFS on OK log rows only (the post's core stays SUMIF); blank waste is Check: waste, not 0%; a text or negative Units needed shows Check in K only (the costs still show); log Status and the name check H27 were added.

Key formulas (row 5)
- L5 `=IF(AND(TRIM(A5)="",COUNTA(B5:E5)=0,J5=""),"",IF(NOT(AND(ISNUMBER(C5),ISNUMBER(D5))),"Check: cost or units",IF(OR(C5<=0,D5<=0),"Check: cost or units",IF(NOT(ISNUMBER(E5)),"Check: waste",IF(OR(E5<0,E5>=1),"Check: waste","OK")))))`
- F5 `=IF(OR(L5="",L5="Check: cost or units"),"",C5/D5)`; G5 `=IF(L5<>"OK","",D5*(1-E5))`
- H5 `=IF(L5<>"OK","",C5/(D5*(1-E5)))` (core `=pack_cost/(units_per_pack*(1-waste))`); I5 `=IF(L5<>"OK","",C5/D5*(1+E5))` (the shortcut)
- K5 `=IF(OR(L5<>"OK",J5=""),"",IF(AND(ISNUMBER(J5),J5>=0),ROUNDUP(ROUND(J5/(D5*(1-E5)),6),0),"Check"))`; M5 `=IF(ISNUMBER(K5),K5*D5*(1-E5),"")`
- P4 `=SUMPRODUCT(--(L5:L54<>""))`; P5 `=P4-COUNTIF(L5:L54,"OK")`; P6 `=COUNTIF(K5:K54,"Check")`; P7 `=COUNTA(A55:J300)`; P8 `=IF(AND(P4>0,P5=0,P6=0,P7=0),"OK","Check")`
- Log E5 `=IF(AND(A5="",TRIM(B5)="",C5="",D5=""),"",IF(TRIM(B5)="","Check: no material",IF(NOT(AND(ISNUMBER(C5),ISNUMBER(D5))),"Check: number",IF(OR(C5<=0,D5<0,D5>C5),"Check: usable vs bought","OK"))))`
- Log H5 `=IF(TRIM(G5)="","",COUNTIFS($B$5:$B$204,G5,$E$5:$E$204,"OK"))`; I5 `=IF(TRIM(G5)="","",SUMIFS($C$5:$C$204,$B$5:$B$204,G5,$E$5:$E$204,"OK"))`; J5 same on D; K5 `=IF(TRIM(G5)="","",IFERROR(1-J5/I5,""))` (core `=1-SUMIF(log_material,name,log_usable)/SUMIF(log_material,name,log_bought)`); L5 `=IF(TRIM(G5)="","",COUNTIF($B$5:$B$204,G5)-H5)`
- Log H27 `=SUMPRODUCT((TRIM(B5:B204)<>"")*(COUNTIF(G5:G24,B5:B204)=0))`; H29 `=IF(AND(H26=0,H27=0,H28=0),"OK","Check")`

Sample values (fictional; recalculated)
| Material | Pack cost | Units | Waste | Per bought unit | Per usable unit | Shortcut | Needed | Packs to buy | Usable from packs |
|---|---|---|---|---|---|---|---|---|---|
| Garden roses (stems) | $30.00 | 25 | 20.0% | $1.20 | $1.50 | $1.44 | 200 | 10 | 200 |
| Quilting cotton (yd) | $9.50 | 1 | 15.0% | $9.50 | $11.1765 | $10.925 | 30 | 36 | 30.6 |
| Pine board (ft) | $14.40 | 8 | 12.5% | $1.80 | $2.0571 | $2.025 | 60 | 9 | 63 |
| Green coffee (lb) | $6.75 | 1 | 18.0% | $6.75 | $8.2317 | $7.965 | 40 | 49 | 40.18 |
| Garden roses (waste from log) | $30.00 | 25 | 18.0% | $1.20 | $1.4634 | $1.416 | 200 | 10 | 205 |
Summary P4:P8: 5, 0, 0, 0, OK. Waste log: Garden roses 25 -> 21, 25 -> 19, 50 -> 42 (11/02, 11/05, 11/09/2026), Quilting cotton 10 -> 8.5, 6 -> 5.25; table G5:L7: Garden roses 3 rows, 100 bought, 82 usable, 18.00%; Quilting cotton 2, 16, 13.75, 14.06% (0.140625); Pine board 0 rows, measured blank. H26:H29: 0, 0, 0, OK. Quotable: the shortcut is 6 cents low per rose stem ($1.44 vs $1.50); averaging the three rose rows' own percents (16%, 24%, 16%) gives 18.67% (probe) instead of 18.00%.

Edge cases tested (all matched a Python model)
- 0% ($1.20 both ways, 8 packs for 200), 99% ($120.00 per usable stem, 800 packs), 100%, 120%, 20 typed instead of 20%, -10%, blank and "20%" text waste: Check: waste. Units per pack 0 and blank, "$30.00" as text, cost 0, a name only: Check: cost or units. Needed blank (K blank), "200" text and -5 (K Check), 0 (0 packs), 12.75 yd (15 packs). Units per pack 2.5.
- Log: usable more than bought and 0 bought (Check: usable vs bought), "garden roses " with a trailing space (not counted; H27 shows it), "GARDEN ROSES" (counted: LibreOffice SUMIFS and COUNTIFS ignore case), "8" as text (Check: number), no name, an unknown material (H27), a material with no rows linked from Materials (Check: waste), rows below 54 and 204, 50 materials x 200 log rows, 8 random sets.

Pitfalls the post must state
- The waste percent is the reader's own measurement; the sample percents are made up; no typical waste or roast loss figures.
- Measure waste as total usable over total bought, not the average of each row's percent (18.00% vs 18.67%).
- A log name with an extra space is not counted (link clean-customer-list-trim-proper).
- Not pricing or markup advice (link markup-vs-margin); not cost of goods sold, no inventory accounting or tax. Pack rounding: link round-order-quantity-up-to-case-size.
- Not tested: Excel, Google Sheets (SUMIFS/COUNTIFS case handling, COUNTIF with a range as the criteria inside SUMPRODUCT, ROUNDUP after ROUND are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (template copies and probe)
- The cost per usable unit looks too low: the shortcut (cost x (1 + waste)) was used ($1.44 vs $1.50).
- Status says Check: waste: waste typed as 20 instead of 20%, blank, 100% or more, or negative.
- The measured waste ignores some log rows: the name has a trailing space or another spelling (H27), or the row shows Check (usable more than bought).
- Packs to buy shows Check: the needed amount is text or negative.

---
## 3. three-point-estimate-project-hours-quote-range-pert-sumsq-excel-google-sheets
Title (from topics): How to Estimate Project Hours With a Three-Point (PERT) Estimate in Excel and Google Sheets
Template: templates/tidy-tabs-three-point-project-estimate.xlsx (tabs: Estimate, How to use)

Layout
- Estimate: A1 Hourly rate, B1 $85.00 (input); C1 note. Header row 4, rows 5:44 (6 sample tasks): A Task, B Best case hours (O), C Most likely hours (M), D Worst case hours (P) (inputs); E Expected hours (O+4M+P)/6, F Spread (P-O)/6, G Status (OK, Check red). Summary I4:J18: Tasks counted (J4), Best case total (J5), Most likely total (J6), Worst case total (J7), Expected total (J8), Project spread (J9), Expected + 1 spread (J10), Expected + 2 spreads (J11), Task spreads added (do not use, too wide) (J12), Quote at expected hours (J13), at + 1 spread (J14), at + 2 spreads (J15), Rows showing Check (J16), Typed below row 44 (J17), CHECK (J18). Landscape.
- Deviations: E and F stay blank on a Check row (topic: the 5/3/8 row still computed 4.1667), so every total uses OK rows only and J8, J9 are the plain core formulas; numbers with no task name are counted (OK), a name with no numbers is Check; rate blank, text or 0 blanks the quotes and makes CHECK Check; the spreads-added trap is in the sheet (J12).

Key formulas (row 5)
- G5 `=IF(AND(TRIM(A5)="",COUNTA(B5:D5)=0),"",IF(COUNT(B5:D5)<3,"Check",IF(AND(B5>=0,B5<=C5,C5<=D5),"OK","Check")))`
- E5 `=IF(G5<>"OK","",(B5+4*C5+D5)/6)` (core `=(O+4*M+P)/6`); F5 `=IF(G5<>"OK","",(D5-B5)/6)` (core `=(P-O)/6`)
- J4 `=COUNTIF(G5:G44,"OK")`; J5:J7 `=SUMIF(G5:G44,"OK",B5:B44)` (C, D); J8 `=SUM(E5:E44)`; J9 `=SQRT(SUMSQ(F5:F44))`; J10 `=J8+J9`; J11 `=J8+2*J9`; J12 `=SUM(F5:F44)`
- J13 `=IF(AND(ISNUMBER($B$1),$B$1>0),ROUND(J8*$B$1,2),"")` (J14 on J10, J15 on J11); J16 `=COUNTIF(G5:G44,"Check")`; J17 `=COUNTA(A45:D300)`; J18 `=IF(AND(J4>0,J16=0,J17=0,AND(ISNUMBER($B$1),$B$1>0)),"OK","Check")`

Sample values (fictional web project; rate $85.00; recalculated)
| Task | O | M | P | Expected | Spread | Status |
|---|---|---|---|---|---|---|
| Discovery call and brief | 1 | 2 | 4 | 2.17 | 0.50 | OK |
| Wireframes | 4 | 6 | 10 | 6.33 | 1.00 | OK |
| Design mockups | 8 | 12 | 20 | 12.67 | 2.00 | OK |
| Build pages | 12 | 16 | 30 | 17.67 | 3.00 | OK |
| Revisions | 2 | 4 | 12 | 5.00 | 1.67 | OK |
| Launch and handoff | 1 | 2 | 3 | 2.00 | 0.33 | OK |
Summary J4:J18: 6, 28.00, 42.00, 79.00, 45.83 (45.8333333333333), 4.14 (4.13991411612474), 49.97 (49.9732474494581), 54.11 (54.1131615655828), 8.50, $3,895.83, $4,247.73, $4,599.62, 0, 0, OK. Quotable: Design mockups 8 / 12 / 20 gives 12.67 expected hours, not 12.

Edge cases tested (all matched a Python model)
- O = M = P (3 / 3 / 3: expected 3, spread 0), O > M (5 / 3 / 8) and M > P (Check, left out of totals), negative O, text "2", blank P, a name only (Check), 0.25 / 0.5 / 1.75 (0.67 and 0.25), numbers with no name (counted), 0 / 0 / 0 (OK, 0). With those rows added to the sample: expected total 51.50, spread 4.16, 6 Check rows, CHECK Check.
- Blank rows between tasks, one task, no tasks (totals 0, CHECK Check), rate blank / "$85.00" text / 0 (quotes blank, CHECK Check), $112.50, a row on 45, 40 tasks, 8 random lists with bad rows and rates.

Pitfalls the post must state
- No probability claims (no "84% chance", no confidence levels); call +1 and +2 spreads a buffer, not a guarantee.
- Adding the task spreads (8.5 h) overstates the project spread; SQRT(SUMSQ()) is the usual way to combine spreads of independent tasks.
- The rate and the quote are the reader's choice; no pricing advice. Link freelance-hourly-rate-calculator and real-hourly-rate-fixed-price-project.
- Not tested: Excel, Google Sheets (SUMSQ on a range with "" cells, ROUND of hours x rate are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (template copies)
- A task is missing from the total: its Status is Check (hours out of order, negative, blank or text).
- The project buffer looks too big: the task spreads were added (J12, 8.50) instead of SQRT(SUMSQ()) (J9, 4.14).
- The quotes are blank: the rate in B1 is blank, text or 0.

---
## 4. net-promoter-score-from-customer-survey-0-10-countifs-excel-google-sheets
Title (from topics): How to Calculate a Net Promoter Score From a 0-10 Customer Survey in Excel and Google Sheets (hooks final: "From a 0 to 10 Customer Survey")
Template: templates/tidy-tabs-nps-survey-score.xlsx (tabs: Responses, Summary, How to use)

Layout
- Responses: A1:A2 notes. Header row 4, rows 5:504 (20 sample answers): A Date, B Channel (dropdown from Summary A6:A9, warning only), C Score 0-10, D Comment (inputs); E Valid score (helper), F Group (Promoter green, Passive, Detractor, Check red). Landscape.
- Summary: A1:A2 notes. Header row 4, A Channel, B Responses counted, C Promoters (9-10), D Passives (7-8), E Detractors (0-6), F % promoters, G % detractors, H NPS, I NPS (not rounded). Row 5 All responses; rows 6:9 channels (A6:A9 inputs: Etsy, Market, Email, blank). A11:B14: Rows showing Check (B11), Counted scores with no listed channel (B12), Typed below row 504 (B13), CHECK (B14). Landscape.
- Deviations: the summary is a table (rows = all and each channel) instead of a list; I (not rounded) and B12 were added; sample channels give Email 5 counted, NPS 20 (the topic's Email 6 / 17 came from a 21-score probe set that does not add up to the 20 sample scores; Etsy 50 and Market 14 are unchanged).

Key formulas (row 5)
- E5 `=IF(AND(ISNUMBER(C5),C5>=0,C5<=10),IF(C5=ROUND(C5,0),C5,""),"")`; F5 `=IF(C5="","",IF(ISNUMBER(E5),IF(E5>=9,"Promoter",IF(E5>=7,"Passive","Detractor")),"Check"))`
- Summary B5 `=COUNT(Responses!$E$5:$E$504)`; C5 `=COUNTIF(Responses!$E$5:$E$504,">=9")`; D5 `=COUNTIFS(Responses!$E$5:$E$504,">=7",Responses!$E$5:$E$504,"<=8")`; E5 `=COUNTIF(Responses!$E$5:$E$504,"<=6")`
- F5 `=IF(N(B5)=0,"",C5/B5)`; G5 `=IF(N(B5)=0,"",E5/B5)`; H5 `=IF(N(B5)=0,"",ROUND((C5-E5)/B5*100,0))` (core `=(COUNTIF(scores,">=9")-COUNTIF(scores,"<=6"))/COUNT(scores)*100`); I5 the same without ROUND
- B6 `=IF(TRIM($A6)="","",COUNTIFS(Responses!$B$5:$B$504,$A6,Responses!$E$5:$E$504,">=0"))`; C6 the same with ">=9", D6 with ">=7" and "<=8", E6 with "<=6"
- B11 `=COUNTIF(Responses!F5:F504,"Check")`; B12 `=B5-SUM(B6:B9)`; B13 `=COUNTA(Responses!A505:D2000)`; B14 `=IF(AND(B5>0,B11=0,B12=0,B13=0),"OK","Check")`

Sample values (fictional; 20 scores 10, 9, 9, 8, 6, 10, 7, 10, 9, 5, 3, 8, 10, 9, 7, 6, 10, 2, 9, 10, dated 11/02 to 11/11/2026; recalculated)
| Channel | Counted | Promoters | Passives | Detractors | % promoters | % detractors | NPS |
|---|---|---|---|---|---|---|---|
| All responses | 20 | 11 | 4 | 5 | 55.0% | 25.0% | 30 |
| Etsy | 8 | 5 | 2 | 1 | 62.5% | 12.5% | 50 |
| Market | 7 | 3 | 2 | 2 | 42.9% | 28.6% | 14 (14.29) |
| Email | 5 | 3 | 0 | 2 | 60.0% | 40.0% | 20 |
B11:B14: 0, 0, 0, OK.
Hook fix (tasks/hooks-2026-11-17.md, opening 4 corrected): raw-column COUNTA version (probe): 30 on the clean 20; 31.82 with an 11 and a text "9" added (COUNTA 22); 30.43 with an 11, a text "9" and an 8.5 added (COUNTA 23). The opening now names all three; all three show Check in the template.

Edge cases tested (all matched a Python model)
- Text "9", a blank, 11, 8.5, -1 and text " 10" added: still 20 counted, NPS 30, 5 rows Check (the blank stays blank), CHECK Check. Only 11 and text "9": 20 counted, NPS 30, 2 Check.
- All promoters (100), all detractors (-100), all passives (0), no responses (NPS blank, no #DIV/0!, CHECK Check). Boundaries 0, 6, 7, 8, 9, 10, 10.0 (counted), 6.0000001 and 9.5 (Check). Half-point NPS: 1 promoter and 7 passives gives 13 (12.5), 1 detractor and 7 passives gives -13.
- Channels: "etsy" and "ETSY" counted as Etsy (LibreOffice COUNTIFS ignores case), "Etsy " with a space, a blank channel and "Wholesale" not counted per channel (B12 3, CHECK Check); a fourth channel typed in A9 (Wholesale 2 counted, NPS 50); channel names removed (rows blank, B12 20). A row on 505, 500 rows, 8 random sets with text, decimals, out-of-range scores and odd channels.

Pitfalls the post must state
- Net Promoter Score and NPS are trademarks of their owners (Bain & Company, Satmetrix/NICE and Fred Reichheld): one descriptive line, no endorsement wording (the How to use tab carries it). No benchmark claims, no survey wording or sample size advice, no survey tool export named or described.
- Use COUNT on the validated helper, not COUNTA on the raw column.
- A channel name with an extra space is not counted for that channel; upper/lower case is ignored (LibreOffice).
- Not tested: Excel, Google Sheets (COUNTIFS case handling, ROUND of -12.5, COUNTIF ">=9" on a text "9" are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- The score is a little off: COUNTA counted text and out-of-range rows (probe: 31.82 and 30.43 instead of 30).
- A row shows Check: a score typed as text, a decimal such as 8.5, 11 or -1.
- The channel scores do not add up to the total: a channel typed "Etsy " or a channel not listed in A6:A9 (B12).
- NPS is blank: no valid scores yet (no #DIV/0!).

---
## 5. average-without-highest-and-lowest-value-trimmean-excel-google-sheets
Title (from topics): How to Average Without the Highest and Lowest Value in Excel and Google Sheets
Template: templates/tidy-tabs-average-without-highest-lowest.xlsx (tabs: Sales, How to use)

Layout
- Sales: A1 note. Header row 4, rows 5:104 (10 sample Saturdays): A Market date, B Sales, C Note (inputs); D Status (OK, Skipped: blank and Skipped: text orange, Check: 0 or less red). Summary F4:G17: Days counted (G4), Plain average (G5), Median (G6), Highest day (G7), Lowest day (G8), Average without highest and lowest (G9), TRIMMEAN check, percent 2/COUNT (G10), The two match? (G11), Drop top and bottom 2 (5+ days) (G12), TRIMMEAN(B5:B104,0.1), for comparison (G13), Rows skipped (G14), Rows showing Check (G15), Typed below row 104 (G16), CHECK (G17). Landscape.
- Deviations: D Status added; G7:G8 show the highest and lowest; G13 shows the 0.1 trap in the file; a 0 or negative day counts in the math but shows Check: 0 or less and CHECK Check.

Key formulas
- D5 `=IF(AND(A5="",B5="",TRIM(C5)=""),"",IF(B5="","Skipped: blank",IF(NOT(ISNUMBER(B5)),"Skipped: text",IF(B5<=0,"Check: 0 or less","OK"))))`
- G4 `=COUNT(B5:B104)`; G5 `=IF(G4=0,"",AVERAGE(B5:B104))`; G6 MEDIAN, G7 MAX, G8 MIN (same guard)
- G9 `=IF(COUNT(B5:B104)<3,"Need 3+ days",(SUM(B5:B104)-MAX(B5:B104)-MIN(B5:B104))/(COUNT(B5:B104)-2))` (core `=(SUM(r)-MAX(r)-MIN(r))/(COUNT(r)-2)`)
- G10 `=IF(COUNT(B5:B104)<3,"",TRIMMEAN(B5:B104,2/COUNT(B5:B104)))` (core `=TRIMMEAN(r,2/COUNT(r))`); G11 `=IF(G4<3,"",IF(ROUND(G9-G10,9)=0,"OK","Check"))`
- G12 `=IF(COUNT(B5:B104)<5,"Need 5+ days",TRIMMEAN(B5:B104,4/COUNT(B5:B104)))`; G13 `=IF(COUNT(B5:B104)<3,"",TRIMMEAN(B5:B104,0.1))`
- G14 `=COUNTIF(D5:D104,"Skipped: blank")+COUNTIF(D5:D104,"Skipped: text")`; G15 `=COUNTIF(D5:D104,"Check: 0 or less")`; G16 `=COUNTA(A105:C1000)`; G17 `=IF(AND(G4>=3,G11="OK",COUNTIF(D5:D104,"Skipped: text")=0,G15=0,G16=0),"OK","Check")`

Sample values (fictional Saturdays 09/05/2026 to 11/07/2026; recalculated): $612.00, $655.00, $590.00, $1,480.00 (Holiday rush, 09/26/2026), $640.00, $128.00 (Rained out, 10/10/2026), $700.00, $668.00, $623.00, $645.00.
| Summary | Value |
|---|---|
| Days counted (G4) | 10 |
| Plain average (G5) | $674.10 |
| Median (G6) | $642.50 |
| Highest / lowest (G7, G8) | $1,480.00 / $128.00 |
| Average without highest and lowest (G9) | $641.63 (641.625) |
| TRIMMEAN 2/COUNT (G10) | $641.63, G11 OK |
| Drop top and bottom 2 (G12) | $640.50 |
| TRIMMEAN 0.1 (G13) | $674.10 (nothing dropped) |
G14:G17: 0, 0, 0, OK.

TRIMMEAN rounding (open question, resolved in LibreOffice 24.2.7.2 only): the number dropped from each end is floor(count x percent / 2). 10 values: 0.1 -> $674.10 (plain average), 0.2 -> 641.625. 11 values: 0.1 -> plain average 677.36, 2/COUNT 649.22. 15 values (the 10 plus $700, $710, $560, $690, $1,200): average $706.73, 0.1 -> 706.73 (plain), 0.2 -> 691.769 (one from each end, same as SUM-MAX-MIN), 4/15 -> 657.545 (two from each end). 21 values: 0.1 drops one from each end (2.1 -> 2). TRIMMEAN(r,2/COUNT(r)) matched SUM-MAX-MIN for every count from 3 to 100 in the probe, including 49 and 98 where 2/n x n is 1.9999999999999998 in binary floating point (LibreOffice still dropped one from each end). Excel and Sheets not checked; the post explains the rounding only as seen in LibreOffice and links Microsoft's TRIMMEAN help page (read it before quoting it).

Edge cases tested (all matched a Python model)
- 2 values (Need 3+ days, G10 blank, CHECK Check; probe: #DIV/0! from the SUM version and #VALUE! from TRIMMEAN in LibreOffice), 3 values 10, 20, 90 (20), 4 values, 1 value, no rows, ties 500, 500, 600, 700, 700 (600: one 500 and one 700 dropped), all equal.
- Blank and "n/a" rows (skipped, COUNT 5), "$612.00" as text (9 days counted, 645.86, Skipped: text, CHECK Check), a 0 for a closed day (11 days, lowest $0.00, 584.56, Check: 0 or less), a negative refund, a note-only row, a row on 105, 49, 98 and 100 rows, 8 random lists of 4 to 100 values.

Pitfalls the post must state
- Not a rule for which average is right; plain, median and trimmed answer different questions (link median-order-value-vs-average and find-outlier-order-amounts-quartile-iqr). No forecast.
- A fixed percent can drop nothing: TRIMMEAN(r,0.1) on 10 values is the plain average (LibreOffice). 2/COUNT(r) adjusts as the list grows.
- Leave a closed day blank; a 0 counts as the lowest day. Text amounts are skipped silently (Days counted shows it).
- Ties: only one highest and one lowest value are dropped.
- Not tested: Excel, Google Sheets (TRIMMEAN rounding, 2/COUNT at 49 and 98 values, the 2-value error names are the open questions; do not name the error Excel or Sheets shows).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2 (probe workbook and template copies)
- TRIMMEAN gives the plain average: the percent is too small for the count (0.1 on 10 values).
- The trimmed average dropped a closed day instead of the real low day: a 0 was typed (template copy: lowest $0.00, Check: 0 or less).
- Fewer days than rows: a sales amount typed as text (Days counted 9, Skipped: text).
- #DIV/0! or #VALUE!: only 2 values (probe); the file shows Need 3+ days.

---
What was NOT run (all five): Microsoft Excel, Google Sheets, any non-US language setting, more rows than the slots (24 boxes / 104 orders, 54 materials / 204 log rows, 44 tasks, 504 responses, 104 days), a printer (PDF export only), diagonal packing, carrier rules, probability levels for the PERT buffer, NPS benchmarks, survey tool exports. Nothing is claimed about Excel or Sheets.
