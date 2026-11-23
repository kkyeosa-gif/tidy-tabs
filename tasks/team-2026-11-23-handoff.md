# Handoff 2026-11-23: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-23.md (researcher) and tasks/hooks-2026-11-23.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`invoices_add_up_to_payment` with `INVOICE_COMBO_MAX`, `print_size_dpi_checker` with `PRINT_SIZE_ROWS`, `maintenance_due_tracker`, `dsum_criteria_box_report`, `season_from_date`, registered after `bought_together_pairs()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-23 build and test"). Heroes are not made yet (designer, see each section).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 120 edited workbooks recalculated and compared (1,843,021 cell comparisons, 0 mismatches against independent Python models; a mutation run with one wrong formula per template produced mismatches in all five, so the harness catches errors), plus 6 probe workbooks of the traps read by hand; the only error values in the 120 are the expected #VALUE! cells of the "Amt" label trap. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text except How to use sentences that name a trap. One LibreOffice behavior differed from the researcher's probe (all-blank criteria box, topic 4) and seven layout problems were found and fixed (tasks/briefs.md).

Repo note: the whole `python3 scripts/build-templates.py` was not run (it rewrites older templates byte-wise); the five files were built by importing the module and calling only the five new functions. `git status`: scripts/build-templates.py, tasks/briefs.md, this handoff, five new .xlsx.

Post file names: `posts/ready/2026-11-23-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks every status starting "Check" and any CHECK cell that is not OK; light green marks "Yes" (Invoices) and Match 1 rows (Combos) and "OK" print sizes; orange marks "Too small" and "None" (Images) and "Overdue" (Equipment); light blue marks "Due soon". Functions used in the five files: ABS, AND, COUNT, COUNTA, COUNTIF, DATE, DAVERAGE, DAY, DCOUNT, DMAX, DMIN, DSUM, EDATE, IF, IFERROR, INDEX, INT, ISERROR, ISNUMBER, ISTEXT, LEFT, LEN, MATCH, MAX, MIN, MOD, MONTH, NOT, OR, ROUND, ROUNDDOWN, ROWS, SUM, SUMIF, SUMIFS, SUMPRODUCT, TEXT, TRIM. No TODAY, FILTER, UNIQUE, SORT, LET, LAMBDA, XLOOKUP, TEXTSPLIT, MINIFS/MAXIFS; 0 `_xlfn` in the files. Help links were not checked by the builder.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-23-team-01-find-which-invoices-add-up-to-a-payment-amount-mod-int-sumproduct-excel-google-sheets | tidy-tabs-invoices-that-add-up-to-payment.xlsx | 25 workbooks, 1,640,175 comparisons |
| 2 | 2026-11-23-team-02-check-image-pixels-for-print-size-300-dpi-crop-percent-min-max-excel-google-sheets | tidy-tabs-print-size-dpi-checker.xlsx | 19 workbooks, 132,772 comparisons |
| 3 | 2026-11-23-team-03-maintenance-due-date-by-months-or-miles-whichever-comes-first-edate-min-excel-google-sheets | tidy-tabs-maintenance-due-tracker.xlsx | 18 workbooks, 6,426 comparisons |
| 4 | 2026-11-23-team-04-total-sales-with-criteria-box-and-or-conditions-dsum-dcount-excel-google-sheets | tidy-tabs-dsum-criteria-box-report.xlsx | 36 workbooks, 18,900 comparisons |
| 5 | 2026-11-23-team-05-assign-season-and-rate-to-any-date-seasons-crossing-new-year-month-day-match-excel-google-sheets | tidy-tabs-season-from-date.xlsx | 22 workbooks, 44,748 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets." For the DSUM post add: "Criteria text matching and blank criteria rows were checked in LibreOffice Calc only; Excel's rule is quoted from Microsoft's help page."

Hooks check (tasks/hooks-2026-11-23.md, chosen openings): numbers match the built files for 1 (4,095 combinations; $416.24 -> 2 sets, INV-1042 + INV-1043 + INV-1049 and INV-1043 + INV-1048 + INV-1052), 2 (3600 x 5400: 8x10 450 DPI and crops 17%, 16x20 Too small: 225 DPI, 12 x 18 in; note the 450 is in the hidden helper column Z5:AO, the visible cell says "OK (crops 17%)"), 3 (van due 01/06/2027 instead of 01/10/2027, 44 days after 11/23/2026), 5 (12/31/2026 and 01/01/2027 Holiday $249; MATCH without the IF gave #N/A for 01/01 to 01/05: reproduced with LOOKUP and with plain MATCH on the template). Opening 4: $256.50 across 7 sales is correct but only after the reader types ="=Market" in C5; the download opens with ="=Etsy" alone ($104.50, 4 sales), so write "add Market on the next row and the total becomes $256.50". The hooks file's claims note for topic 4 says a box with only blank rows "totals everything": that does NOT hold for the built template (7-column box): LibreOffice returned $0.00 / 0 (in a plain workbook, everything for 1 to 3 columns, 0 for 4 or more). Drop that sentence; the post says to type at least one condition and makes no claim about an empty box. Formulas in the openings are the core forms; the cells hold guarded versions (below).

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=...)` fits the whole sheet, or its print area, on one page): Invoices (A1:I18) and Seasons (A1:I16) and Report (A1:J18) are small and render as they are. Combos has print area A1:P36, which render() keeps, so it renders the first 32 combinations. Images has print area A1:X104 and Equipment, Sales and Dates fill gray formulas to rows 54, 504 and 504, so a whole-sheet render of those is tiny: use a copy with the empty rows deleted or a print area (rows below).

---
## 1. find-which-invoices-add-up-to-a-payment-amount-mod-int-sumproduct-excel-google-sheets
Title (from topics): How to Find Which Invoices Add Up to a Payment in Excel and Google Sheets
Template: templates/tidy-tabs-invoices-that-add-up-to-payment.xlsx (tabs: Invoices, Combos, How to use)

Layout
- Invoices: A1 Payment received / B1 (input), A2 Allowed difference / B2 (input), C1:C2 notes. Header row 4, rows 5:16: A Invoice #, B Amount (inputs); C Status, D Match 1, E Match 2, F Match 3. C17/D17:F17 Total of this set, C18/D18:F18 Combination # (Combos tab). Summary H4:I12: Invoices with an amount (numbers) (I4), Amounts typed as text (must be 0) (I5), Combinations checked (2^n - 1) (I6), Matches found (I7), Result (I8), Rows showing Check (I9), Typed below row 16 (only 12 supported) (I10), Payment and allowed difference (I11), CHECK (I12). Landscape.
- Combos: A1:A2 notes, A3 "Amount" and B3:M3 the amounts, row 4 header A Combination #, B4:M4 invoice numbers by formula, N Total, O Match (1/0), P Match no.; rows 5:4099. Print area A1:P36. Landscape.
- Deviations from the topic: Combos is visible (not hidden or grouped) with a print area of the first 32 combinations; O5 also waits for I11 OK (payment a number above 0, difference a number of 0 or more), so $0.00 or a blank payment gives "Type the payment in B1"; rows 17/18 added; Status adds no amount, no invoice #, invoice # twice, more than 2 decimals; a negative amount is still used in the combinations but is a Check.

Key formulas (exact text in the file)
- Combos B3 `=INDEX(Invoices!$B$5:$B$16,1)` (2 ... 12 across to M3)
- Combos A5 `=IF(ROWS(A$5:A5)>2^COUNT(Invoices!$B$5:$B$16)-1,"",ROWS(A$5:A5))`
- Combos B5 `=IF($A5="","",MOD(INT($A5/2^0),2))`, C5 `...2^1...`, through M5 `=IF($A5="","",MOD(INT($A5/2^11),2))` (core `=MOD(INT(n/2^k),2)`)
- Combos N5 `=IF($A5="","",ROUND(SUMPRODUCT($B$3:$M$3,B5:M5),2))`
- Combos O5 `=IF(OR(N5="",Invoices!$I$11<>"OK"),"",IF(ABS(N5-ROUND(Invoices!$B$1,2))<=Invoices!$B$2,1,0))`
- Combos P5 `=IF(O5=1,COUNTIF($O$5:O5,1),"")`
- Invoices D5 `=IFERROR(IF(INDEX(Combos!$B$5:$M$4099,MATCH(1,Combos!$P$5:$P$4099,0),ROWS(D$5:D5))=1,"Yes",""),"")` (E uses MATCH(2,...), F MATCH(3,...))
- Invoices D17 `=IF(COUNTIF(D5:D16,"Yes")=0,"",SUMIF(D5:D16,"Yes",$B$5:$B$16))`; D18 `=IFERROR(INDEX(Combos!$A$5:$A$4099,MATCH(1,Combos!$P$5:$P$4099,0)),"")`
- Invoices C5 `=IF(AND(A5="",B5=""),IF(COUNTA(A5:$B$16)>0,"Check: blank row above a filled row",""),IF(B5="","Check: no amount",IF(NOT(ISNUMBER(B5)),"Check: amount is text",IF(TRIM(A5)="","Check: no invoice #",IF(COUNTIF($A$5:$A$16,A5)>1,"Check: invoice # twice",IF(B5<0,"Check: negative amount (credit memo)",IF(B5<>ROUND(B5,2),"Check: more than 2 decimals","OK")))))))`
- Invoices I4 `=COUNT(B5:B16)`; I5 `=COUNTA(B5:B16)-COUNT(B5:B16)`; I6 `=IF(I4=0,0,2^I4-1)`; I7 `=COUNTIF(Combos!$O$5:$O$4099,1)`; I8 `=IF(I11<>"OK","Type the payment in B1",IF(I7=0,"No match",IF(I7=1,"1 match",I7&" matches"&IF(I7>3," (first 3 listed)",""))))`; I11 `=IF(AND(ISNUMBER(B1),B1>0,ISNUMBER(B2),B2>=0),"OK","Check")`; I12 `=IF(AND(I4>0,I5=0,I9=0,I10=0,I11="OK"),"OK","Check")`

Sample values (fictional; recalculated). Payment $416.24, allowed difference $0.00.
| Invoice # | Amount | Match 1 | Match 2 |
|---|---|---|---|
| INV-1041 | $412.50 | | |
| INV-1042 | $189.99 | Yes | |
| INV-1043 | $76.25 | Yes | Yes |
| INV-1044 | $1,050.00 | | |
| INV-1045 | $230.10 | | |
| INV-1046 | $95.00 | | |
| INV-1047 | $318.40 | | |
| INV-1048 | $64.99 | | Yes |
| INV-1049 | $150.00 | Yes | |
| INV-1050 | $520.35 | | |
| INV-1051 | $88.80 | | |
| INV-1052 | $275.00 | | Yes |
| Total of this set | | $416.24 | $416.24 |
| Combination # | | 262 | 2180 |
Status all OK. I4:I12: 12, 0, 4,095, 2, 2 matches, 0, 0, OK, OK. Combos row 5: combination 1, INV-1041 only, $412.50; row 4099: combination 4,095, all 1s, $3,471.38.

Edge cases tested (all matched the Python model)
- $573.88: 1 match (INV-1042, INV-1045, INV-1048, INV-1051, combination 1170). $3,471.38: combination 4,095 (all 12).
- $413.24 with Allowed difference $3.00: 5 matches, "5 matches (first 3 listed)": INV-1041 $412.50 (combination 1), INV-1046 + INV-1047 $413.40 (96), the $416.24 set at 262; also $413.90 (1072) and the second $416.24 set (2180). The $3.00 edge counted as within.
- INV-1045 typed as text "150.00": I4 11, I5 1, 2,047 combinations, INV-1052 never tried, so the INV-1043 + INV-1048 + INV-1052 set was lost, and a false Match 2 appeared: INV-1042 + INV-1043 + INV-1045 + INV-1049 (the text counts as $0.00; D17 also shows $416.24 because SUMIF skips text). Status Check: amount is text, CHECK Check.
- Blank row in the middle (row 13 empty): 11 counted, INV-1052 never tried, No match, Check: blank row above a filled row.
- 7 invoices, payment no set reaches: 127 combinations, No match. 1 invoice: 1 combination, 1 match.
- Same amount twice (A-1 $50.00, A-3 $50.00, A-4 $20.00, payment $70.00): 2 matches that differ by invoice #. Credit memo CM-7 -$45.00: INV-1 + CM-7 = $255.00 matched; Check: negative amount (credit memo).
- Payment $0.00 or blank, or difference typed as text "3": I11 Check, "Type the payment in B1", 0 matches. 16 matches from 10 round amounts (first 3 listed); amounts 1 to 12 with payment 39: 124 matches.
- A 13th invoice on row 17: I10 2, CHECK Check (it is not used). Invoice # twice (X-1 and x-1, case ignored), 10.005 (more than 2 decimals), no invoice #, no amount: Check. 8 random sets of 1 to 12 invoices.

Pitfalls the post must state
- 12 invoices max (4,095 rows; 13 would need 8,191, 20 over a million): narrow the list by client and dates first.
- The sheet finds sets that add up; it does not decide which invoices the client meant. Confirm with the client. No bookkeeping or tax treatment of short payments or fees.
- ROUND is a guard; in our LibreOffice run the totals matched whole-cent sums. No claim that Excel or Sheets fail without it. No speed claim for 4,095 rows in any app.
- A text amount drops out of COUNT and counts as $0.00: compare COUNT with COUNTA (I5).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- No match on a payment that should match: an amount stored as text or a blank row in the list (both cut the combinations; I5 or Status shows it).
- A match lists an invoice whose amount is text: that amount counted as $0.00.
- More than one match: two sets add up to the same amount; ask the client.
- Nothing listed and Result says "Type the payment in B1": payment blank or $0.00, or difference not a number.

Hero: sheet "Invoices" as it is (A1:I18); Combos rows 1 to 20 as an inset (copy with rows 21:4099 deleted, or render() on Combos gives the print area A1:P36).

---
## 2. check-image-pixels-for-print-size-300-dpi-crop-percent-min-max-excel-google-sheets
Title (from topics): How to Check If an Image Has Enough Pixels for a Print Size (300 DPI) in Excel and Google Sheets
Template: templates/tidy-tabs-print-size-dpi-checker.xlsx (tabs: Images, Print sizes, How to use)

Layout
- Print sizes: A1 Minimum DPI / B1 (input, 300), C1 note, A2 note. Header row 4, rows 5:20: A Size name, B Short side (in), C Long side (in) (inputs); D Short (sorted), E Long (sorted), F Area (sq in), G Status. Summary I4:J11: Minimum DPI (B1) (J4), Sizes listed (J5), Rows showing Check (J6), Typed below row 20 (J7), Images listed (J8), Image rows showing Check (J9), Images typed below row 104 (J10), CHECK (J11). Landscape.
- Images: A1:A2 notes. Header row 4, rows 5:104: A File name, B Width px, C Height px (inputs); D Long px, E Short px, F Largest print at min DPI, G Largest listed size that prints, H:W one column per size row 5 to 20 (headers by formula: 4x6, 5x7, 8x10, 8.5x11, 11x14, 11x17, 12x18, 16x20, 18x24, 20x30, 24x36, then 5 empty), X Status. Helpers grouped and hidden: Z:AO effective DPI, AP:BE crop share, BF:BU area if OK (row 3 labels). Print area A1:X104. Landscape.
- Deviations from the topic: sizes are sorted on the Print sizes tab (D/E MIN/MAX), so either side can come first; the CHECK cell is on the Print sizes tab (J11) and covers the Images rows; Status is column X, after the size columns.

Key formulas (exact text in the file)
- Images D5 `=IF($X5<>"OK","",MAX(B5,C5))`; E5 `=IF($X5<>"OK","",MIN(B5,C5))` (core `=MAX(width,height)`, `=MIN(width,height)`)
- Images F5 `=IF(OR($X5<>"OK",'Print sizes'!$J$4<>"OK"),"",ROUNDDOWN(E5/'Print sizes'!$B$1,1)&" x "&ROUNDDOWN(D5/'Print sizes'!$B$1,1)&" in")`
- Images Z5 (DPI for size row 5; AA5 uses row 6, and so on) `=IF(OR($X5<>"OK",'Print sizes'!$J$4<>"OK",'Print sizes'!$G$5<>"OK"),"",MIN($D5/'Print sizes'!$E$5,$E5/'Print sizes'!$D$5))` (core `=MIN(long_px/long_in,short_px/short_in)`)
- Images AP5 (crop) `=IF(Z5="","",1-('Print sizes'!$D$5*'Print sizes'!$E$5*Z5^2)/($D5*$E5))`
- Images BF5 (area if OK) `=IF(Z5="","",IF(ROUND(Z5,6)>='Print sizes'!$B$1,'Print sizes'!$F$5,0))`
- Images H5 `=IF(Z5="","",IF(ROUND(Z5,6)>='Print sizes'!$B$1,"OK"&IF(AP5>=0.005," (crops "&TEXT(AP5,"0%")&")",""),"Too small: "&INT(Z5)&" DPI"))`
- Images G5 `=IF(F5="","",IF(MAX(BF5:BU5)=0,"None",INDEX('Print sizes'!$A$5:$A$20,MATCH(MAX(BF5:BU5),BF5:BU5,0))))`
- Print sizes D5 `=IF(G5<>"OK","",MIN(B5,C5))`, E5 `=IF(G5<>"OK","",MAX(B5,C5))`, F5 `=IF(G5<>"OK","",D5*E5)`
- Images X5 `=IF(AND(TRIM(A5)="",B5="",C5=""),"",IF(OR(NOT(ISNUMBER(B5)),NOT(ISNUMBER(C5))),"Check: pixels blank or text",IF(OR(B5<=0,C5<=0),"Check: 0 pixels",IF(TRIM(A5)="","Check: no file name","OK"))))`

Sample values (fictional; recalculated at 300 DPI)
| File | Pixels | Largest print | Largest listed | 5x7 | 8x10 | 11x14 | 16x20 | 24x36 |
|---|---|---|---|---|---|---|---|---|
| wildflowers-print.jpg | 3600 x 5400 | 12 x 18 in | 12x18 | OK (crops 7%) | OK (crops 17%) | OK (crops 15%) | Too small: 225 DPI | Too small: 150 DPI |
| cabin-landscape.jpg | 6000 x 4800 | 16 x 20 in | 16x20 | OK (crops 11%) | OK | OK (crops 2%) | OK | Too small: 166 DPI |
| square-logo.png | 3000 x 3000 | 10 x 10 in | 8x10 | OK (crops 29%) | OK (crops 20%) | Too small: 214 DPI | Too small: 150 DPI | Too small: 83 DPI |
| phone-photo.jpg | 3024 x 4032 | 10 x 13.4 in | 8.5x11 | OK (crops 5%) | OK (crops 6%) | Too small: 274 DPI | Too small: 189 DPI | Too small: 112 DPI |
| small-scan.png | 1200 x 1500 | 4 x 5 in | None | Too small: 214 DPI | Too small: 150 DPI | Too small: 107 DPI | Too small: 75 DPI | Too small: 41 DPI |
| botanical-set-04.jpg | 4961 x 7016 | 16.5 x 23.3 in | 16x20 | OK (crops 1%) | OK (crops 12%) | OK (crops 10%) | OK (crops 12%) | Too small: 194 DPI |
Other cells: wildflowers 4x6 OK, 8.5x11 OK (crops 14%), 11x17 OK (crops 3%), 12x18 OK, 18x24 Too small: 200 DPI; 8x10 helper DPI 450. cabin 4x6 OK (crops 17%), 18x24 Too small: 250 DPI. square-logo 8.5x11 Too small: 272 DPI (272.7). phone-photo 8.5x11 OK (crops 3%). small-scan 4x6 Too small: 250 DPI. botanical 18x24 Too small: 275 DPI. Print sizes J4:J11: OK, 11, 0, 0, 6, 0, 0, OK.

Edge cases tested (all matched the Python model)
- Exactly 300 DPI: 2400 x 3000 on 8x10, 4800 x 6000 on 16x20 and 2100 x 1500 on 5x7 all "OK" with no crop. 2399 x 3000 on 8x10: Too small: 299 DPI (299.875, rounded down).
- Portrait and landscape of the same pixels (3600 x 5400 and 5400 x 3600; 3024 x 4032 and 4032 x 3024): identical rows. Sizes typed long side first: same results.
- Square images and square sizes (12x12 and 8x8 added): 3600 x 3600 12x12 OK, 8x10 OK (crops 20%); 2400 x 2400 12x12 Too small: 200 DPI, 8x8 OK; 3600 x 2400 8x8 OK (crops 33%).
- Minimum DPI 150: wildflowers 24 x 36 in, largest 24x36 OK; square-logo 16x20 OK (crops 20%). Minimum "300" as text or 0: J4 Check, labels blank.
- A-series pixel files: 2480 x 3508 8x10 OK (crops 12%), 8.5x11 Too small: 291 DPI; 3508 x 4961 11x14 OK (crops 10%); 1748 x 2480 5x7 OK (crops 1%).
- Bad images (0, -10, "3600" text, blank height, no file name: Check; a blank row between files is fine). Bad sizes (blank side, "8" text, 0, a name typed twice: both rows flagged, no name). A size on row 21 and an image on row 105: CHECK Check. 100 random images; 6 random sets of 3 to 16 sizes with minimums 72 to 600.

Pitfalls the post must state
- DPI here is pixels per printed inch. 300 is the sample; the reader types the number their print lab's guide gives. No "required" DPI for any lab or marketplace.
- The DPI number stored in the file does not change the pixel count (one line, no app's export dialog).
- Crop share is the area lost when the print is filled edge to edge, not where the crop falls.
- INT rounds the DPI down on purpose (272.7 shows 272 DPI; 299.875 shows 299 DPI and is Too small). ROUND(,6) is the guard that lets exactly 300 count as OK.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A size shows Too small at 299 DPI: the file is a few pixels short (2399 instead of 2400 px).
- A sideways photo fails a portrait size: compare long to long (MAX/MIN); with them, 6000 x 4800 prints 16x20 at exactly 300 DPI.
- Every size is blank: Minimum DPI is text or 0 (J4 Check), or the row shows Check.
- Largest listed size says None: no listed size has enough pixels (small-scan.png).

Hero: sheet "Images", rows 1 to 10 (a copy with rows 11:104 deleted and print area A1:X10), hide_cols "STUVW" (the five empty size columns) so A:R and X show; Print sizes A4:G15 as an inset.

---
## 3. maintenance-due-date-by-months-or-miles-whichever-comes-first-edate-min-excel-google-sheets
Title (from topics): How to Track Maintenance Due by Months or Miles, Whichever Comes First, in Excel and Google Sheets
Template: templates/tidy-tabs-maintenance-due-tracker.xlsx (tabs: Equipment, How to use)

Layout
- Equipment: A1 As-of date / B1 (input, 11/23/2026), A2 Due soon (days) / B2 (input, 14), C1:C2 notes. Header row 4, rows 5:54: A Item and task, B Every N months, C Every N units, D Last service date, E Reading at last service, F Latest reading, G Latest reading date (inputs); H Due by months, I Units per day, J Due by usage, K Due date, L Limit that hits first, M Days left (format "0"), N Status. Summary P4:Q10: As-of date and due-soon days (Q4), Items listed (Q5), Overdue (Q6), Due soon (Q7), Rows showing Check (Q8), Typed below row 54 (Q9), CHECK (Q10). Landscape.
- Deviations from the topic: H, I and J have ISNUMBER guards so text inputs give blanks, not errors; Status has a full Check list and "No due date yet" (a usage-only task with no use since service, or reading date = service date); Overdue is orange and Due soon light blue (red is only for Check).

Key formulas (exact text in the file)
- H5 `=IF(OR(B5="",NOT(ISNUMBER(B5)),NOT(ISNUMBER(D5))),"",EDATE(D5,B5))` (core `=EDATE(last_service,months)`)
- I5 `=IF(OR(C5="",NOT(ISNUMBER(D5)),NOT(ISNUMBER(E5)),NOT(ISNUMBER(F5)),NOT(ISNUMBER(G5))),"",IF(G5<=D5,"",(F5-E5)/(G5-D5)))`
- J5 `=IF(OR(I5="",NOT(ISNUMBER(C5))),"",IF(F5>=E5+C5,G5,IF(I5<=0,"",G5+INT((E5+C5-F5)/I5))))` (core `=G5+INT((E5+C5-F5)/I5)`)
- K5 `=IF(AND(H5="",J5=""),"",MIN(H5,J5))`; L5 `=IF(K5="","",IF(AND(J5<>"",J5<H5),"Usage","Months"))`
- M5 `=IF(OR(K5="",NOT(ISNUMBER($B$1))),"",K5-$B$1)`
- N5 `=IF(AND(TRIM(A5)="",B5="",C5="",D5="",E5="",F5="",G5=""),"",IF(TRIM(A5)="","Check: no item name",IF(NOT(ISNUMBER(D5)),"Check: last service date blank or text",IF(AND(B5="",C5=""),"Check: no months or units",IF(AND(B5<>"",OR(NOT(ISNUMBER(B5)),B5<=0)),"Check: months not a number above 0",IF(AND(C5<>"",OR(NOT(ISNUMBER(C5)),C5<=0)),"Check: units not a number above 0",IF(AND(C5<>"",OR(NOT(ISNUMBER(E5)),NOT(ISNUMBER(F5)),NOT(ISNUMBER(G5)))),"Check: readings blank or text",IF(AND(C5<>"",F5<E5),"Check: reading lower than last service",IF(AND(C5<>"",G5<D5),"Check: reading date before service date",IF(NOT(ISNUMBER($B$1)),"Check: as-of date",IF(K5="","No due date yet",IF(M5<0,"Overdue",IF(M5<=$B$2,"Due soon","OK")))))))))))))`
- Q4 `=IF(AND(ISNUMBER(B1),B1>=1,ISNUMBER(B2),B2>=0),"OK","Check")`; Q10 `=IF(AND(Q4="OK",Q5>0,Q8=0,Q9=0),"OK","Check")`

Sample values (fictional; recalculated). As-of 11/23/2026, due soon 14 days.
| Item and task | Rule | Last service | Latest reading | By months | Per day | By usage | Due | Limit | Days | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Delivery van: oil change | 6 mo / 5,000 | 07/10/2026 at 48,210 | 51,890 on 11/20/2026 | 01/10/2027 | 27.67 | 01/06/2027 | 01/06/2027 | Usage | 44 | OK |
| Kiln: element check (firings) | 12 mo / 200 | 03/02/2026 at 1,140 | 1,302 on 11/21/2026 | 03/02/2027 | 0.61 | 01/21/2027 | 01/21/2027 | Usage | 59 | OK |
| Espresso machine: descale (shots) | 3 mo / 1,500 | 09/15/2026 at 20,500 | 21,480 on 11/22/2026 | 12/15/2026 | 14.41 | 12/28/2026 | 12/15/2026 | Months | 22 | OK |
| Sewing machine: service (hours) | 12 mo / 400 | 12/01/2025 at 3,100 | 3,360 on 11/18/2026 | 12/01/2026 | 0.74 | 05/26/2027 | 12/01/2026 | Months | 8 | Due soon |
| Box truck: tires | 24 mo | 01/31/2025 | | 01/31/2027 | | | 01/31/2027 | Months | 69 | OK |
| Delivery van: tire rotation | 6 mo / 7,500 | 05/31/2026 at 46,000 | 51,890 on 11/20/2026 | 11/30/2026 | 34.05 | 01/06/2027 | 11/30/2026 | Months | 7 | Due soon |
| Walk-in cooler: coil cleaning | 3 mo | 08/15/2026 | | 11/15/2026 | | | 11/15/2026 | Months | -8 | Overdue |
Q4:Q10: OK, 7, 1, 2, 0, 0, OK.

Edge cases tested (all matched the Python model)
- EDATE month ends: 01/31/2026 + 1 = 02/28/2026; 08/31/2026 + 6 = 02/28/2027; 03/31 + 1 = 04/30; 02/29/2028 + 12 = 02/28/2029; 08/31/2027 + 6 = 02/29/2028; 01/31/2028 + 1 = 02/29/2028; 12/31/2026 + 2 = 02/28/2027; 1.5 months used as 1.
- Already past the usage limit (65,400 against 60,000 + 5,000): due on the reading date 11/10/2026, -13, Overdue. Reading exactly at the limit: due on the reading date.
- No use since service with months: months only (04/01/2027, 129 days). Usage-only with no use, or reading date = service date: No due date yet.
- Only months, only units (12/11/2026, Usage), both blank (Check). Odometer typo (4,800 below 48,000): Check: reading lower than last service. Reading date before service: Check.
- Due on the as-of date: 0, Due soon; 14 days Due soon; 15 OK; -1 Overdue. Due soon 0. As-of before the service date (days count from the as-of date).
- Text service date, months, units, reading, reading date; units with no readings; 0 months; negative units; no item name: Check. As-of blank or text: every row Check: as-of date. A row on row 55: CHECK Check. 8 random sets of 10 to 50 rows.

Pitfalls the post must state
- Intervals come from the owner's manual or the maker; the sample rules are made up. No safety or warranty claim.
- The usage date is a straight-line projection from use so far; it moves when use changes. INT rounds it down (a day early rather than late).
- Type the as-of date; the template does not use TODAY. Days left is a number (format it as a number; a date format would show 44 as a 1900 date).
- Not a mileage log for taxes (link the 09-27 mileage log post).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- 01/31 plus 1 month shows 02/28: EDATE moves to the last day of the shorter month.
- Due by usage is blank: no use since the service, or the reading date equals the service date.
- Due date is the reading date and Overdue: the reading is already past the usage limit.
- Status Check: reading lower than last service: an odometer typo.

Hero: sheet "Equipment", rows 1 to 11 (a copy with rows 12:54 deleted), hide_cols "" (A:Q fit landscape).

---
## 4. total-sales-with-criteria-box-and-or-conditions-dsum-dcount-excel-google-sheets
Title (from topics): How to Total Sales With a Criteria Box (DSUM and DCOUNT) in Excel and Google Sheets (hook-writer final: How to Use DSUM With Multiple Criteria (AND and OR) in Excel and Google Sheets)
Template: templates/tidy-tabs-dsum-criteria-box-report.xlsx (tabs: Report, Sales, How to use)

Layout
- Report: A1:A2 notes. Criteria box A3:G6: row 3 labels Date | Date | Channel | Product | State | Amount | Amount (typed, header style), rows 4:6 conditions (inputs; C4 holds ="=Etsy"). Row 7 label checks (A7:G7). Results I3:J15: Criteria rows used (J3), Total (DSUM) (J4), Number of sales (DCOUNT) (J5), Average sale (DAVERAGE) (J6), Largest sale (DMAX) (J7), Smallest sale (DMIN) (J8), Conditions typed in the box (must be 1 or more) (J9), SUMIFS cross-check (one criteria row only) (J10), Difference (DSUM minus SUMIFS) (J11), Labels not found on Sales (J12), Blank criteria rows inside the box (J13), Sales rows showing Check (J14), CHECK (J15). A9:A18 "How to type conditions" lines. Landscape, no repeated row.
- Sales: A1:A2 notes. Header row 4, rows 5:504: A Date, B Channel, C Product, D State, E Amount (inputs); F Status. Summary H4:I8: Sales rows listed (I4), Total of all amounts (I5), Rows showing Check (I6), Amounts typed as text (I7), Typed below row 504 (I8). Landscape.
- Deviations from the topic: the box has 7 columns (Amount twice, so between-two-amounts works) and results are in column J; the box size is automatic (rows 4 to the last filled row) instead of three fixed versions; DMIN added; J9 requires at least one condition (see the all-blank finding); CHECK shows Check (not #VALUE!) when a label is wrong.

Key formulas (exact text in the file)
- J3 `=IF(COUNTA(A6:G6)>0,3,IF(COUNTA(A5:G5)>0,2,1))`
- J4 `=DSUM(Sales!$A$4:$E$504,"Amount",$A$3:INDEX($A$4:$G$6,$J$3,7))` (core `=DSUM(Sales!$A$4:$E$504,"Amount",$A$3:$G$4)`)
- J5 `=DCOUNT(Sales!$A$4:$E$504,"Amount",$A$3:INDEX($A$4:$G$6,$J$3,7))`
- J6 `=IFERROR(DAVERAGE(Sales!$A$4:$E$504,"Amount",$A$3:INDEX($A$4:$G$6,$J$3,7)),"")`
- J7 `=IF(J5=0,"",DMAX(Sales!$A$4:$E$504,"Amount",$A$3:INDEX($A$4:$G$6,$J$3,7)))`; J8 the same with DMIN
- J9 `=COUNTA(A4:G6)`
- J10 `=IF($J$3<>1,"",SUMIFS(Sales!$E$5:$E$504,Sales!$A$5:$A$504,IF(A4="","<>",A4),Sales!$A$5:$A$504,IF(B4="","<>",B4),Sales!$B$5:$B$504,IF(C4="","<>",C4),Sales!$C$5:$C$504,IF(D4="","<>",D4),Sales!$D$5:$D$504,IF(E4="","<>",E4),Sales!$E$5:$E$504,IF(F4="","<>",F4),Sales!$E$5:$E$504,IF(G4="","<>",G4)))`
- J11 `=IF(OR(J10="",ISERROR(J4)),"",ROUND(J4-J10,2))`; J13 `=(COUNTA(A4:G4)=0)*(J3>=2)+(COUNTA(A5:G5)=0)*(J3>=3)`; J15 `=IF(ISERROR(J4),"Check",IF(AND(J9>0,J12=0,J13=0,J14=0,OR(J11="",J11=0)),"OK","Check"))`
- A7 `=IF(A3="","",IF(COUNTIF(Sales!$A$4:$E$4,A3)=0,"Check: not a Sales label",IF(LEN(A3)<>LEN(TRIM(A3)),"Check: extra space","OK")))`
- C4 `="=Etsy"` (shows =Etsy)

Sample values (fictional; recalculated). Sales: 10/03/2026 Etsy Candle MA $28.00; 10/04 Market Soap MA $12.00; 10/11 Etsy Wholesale Candle NY $240.00; 10/18 Market Candle MA $56.00; 11/01 Etsy Soap CA $18.50; 11/02 Website Candle NJ $42.00; 11/08 Market Candle MA $84.00; 11/09 Etsy Candle TX $28.00; 11/15 Website Gift box NY $65.00; 11/20 etsy Candle MA $30.00 ($603.50 in all). Download as shipped (="=Etsy"): J3:J15 1, $104.50, 4, $26.13, $30.00, $18.50, 1, $104.50, $0.00, 0, 0, 0, OK.
| Criteria box (typed) | Total | Sales |
|---|---|---|
| Channel ="=Etsy" | $104.50 | 4 |
| Channel Etsy (plain) | $104.50 | 4 |
| Channel Etsy* | $344.50 | 5 |
| ="=Etsy" on row 4, ="=Market" on row 5 | $256.50 | 7 |
| Product Candle, State MA | $198.00 | 4 |
| Amount >50 | $445.00 | 4 |
| Candle + MA on row 4, Website on row 5 | $305.00 | 6 |
| ="=Etsy", Amount >=20, Amount <=29 | $56.00 | 2 |
| Channel <>Market | $451.50 | 7 |
| Date >=11/01/2026, Date <=11/30/2026 | $267.50 | 6 |

Edge cases tested (all matched the model of LibreOffice's rules)
- =">="&DATE(2026,11,1) and ="<="&DATE(2026,11,30): $267.50, the same as the typed dates. ETSY + candle: $86.00 / 3 (case ignored). Mark?t OR ?tsy: $256.50 / 7. =Amazon: $0.00, 0, average/largest/smallest blank. Amount 28 typed as a number: $56.00 / 2. Field 5 instead of "Amount": same results. Three OR rows: $335.50 / 4.
- Label "Amt": #VALUE! in J4, J5, J7, J8; F7 Check: not a Sales label; CHECK Check. Label "Amount " (trailing space): LibreOffice still matched ($86.00 for Etsy with >20); F7 Check anyway. Labels as formulas (=Sales!$A$4 ...): worked ($256.50).
- Blank row between filled rows: ignored ($244.50 for Etsy, blank, Market + Candle; J13 1, CHECK Check). Blank row 4 with Market on row 5: $152.00 (ignored; Check).
- All criteria blank: $0.00, 0 in this 7-column box (J9 0, CHECK Check). In a plain workbook, an all-blank criteria row matched everything for a 1- to 3-column box and nothing for 4 or 5 columns (probe4). The researcher's $603.50 came from a 2-column box.
- Sales with a text amount "12.00", a text date "11/01/2026" and a blank product: DSUM skips the text amount; <>Website gave $484.50 vs SUMIFS $400.50 (the blank-product sale is left out of SUMIFS), >=11/01/2026 left out the text-date sale ($249.00); Status Check rows, CHECK Check.
- 500 random sales; 10 random 1- to 3-row criteria sets with dates, wildcards, <>, numbers.

Pitfalls the post must state
- Text matching differs between apps: in LibreOffice Calc plain Etsy matched whole cells only (not Etsy Wholesale); Microsoft's criteria page says text without = matches values that begin with it and shows ="=Davolio" for exact matches. Show ="=Etsy". No Sheets claim.
- Size the criteria range to the rows you fill (the template does it with J3). Blank criteria rows were ignored in LibreOffice when another row was filled; an empty box returned $0.00 here. No Excel claim about blank rows; tell the reader to type at least one condition.
- Dates typed as text in the box depend on the locale (MM/DD/YYYY worked in our default en-US headless run); the DATE form is the safe one.
- DAVERAGE with no match is #DIV/0! without IFERROR (wrapped in J6). Not a pivot table post.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- #VALUE! in the results: a box label that is not a Sales header ("Amt").
- Total too high or too low vs SUMIFS: a wildcard (Etsy*) or a blank field in a sale; J11 shows the difference.
- $0.00 with nothing typed: the box is empty; type at least one condition.
- A sale is missing: its amount or date is text (Sales Status).

Hero: sheet "Report" as it is (A1:J18), after typing ="=Market" in C5 if the hero should show the OR example ($256.50, 7) (render from a copy); Sales A4:F14 as an inset.

---
## 5. assign-season-and-rate-to-any-date-seasons-crossing-new-year-month-day-match-excel-google-sheets
Title (from topics): How to Assign a Season and Rate to Any Date (Even Seasons That Cross New Year's) in Excel and Google Sheets (hook-writer final: How to Get the Season and Rate From a Date in Excel and Google Sheets)
Template: templates/tidy-tabs-season-from-date.xlsx (tabs: Seasons, Dates, How to use)

Layout
- Seasons: A1:A2 notes (each season runs until the next start; the last one wraps into January). Header row 4, rows 5:16: A Start month (1-12), B Start day, C Season name, D Rate $ (inputs); E Key (month*100+day), F Status. Summary H4:I7: Seasons listed (I4), Rows showing Check (I5), Typed below row 16 (I6), CHECK (I7). Landscape.
- Dates: A1:A2 notes. Header row 4, rows 5:504: A Date (input); B Key, C Season, D Rate $, E Status. Summary G4:H9: Dates listed (H4), Total of the rates (H5), Rows showing Check (H6), Seasons tab CHECK (H7), Typed below row 504 (H8), CHECK (H9). Landscape.
- Deviations from the topic: the order check sits on the lower row (F6 compares E6 with E5, "Check: sort by start", plus "Check: row above is blank or not valid"); Dates C/D stay blank until Seasons CHECK is OK, so an unsorted list never shows wrong seasons in the template; "Total of the rates" added.

Key formulas (exact text in the file)
- Seasons E5 `=IF(OR(NOT(ISNUMBER(A5)),NOT(ISNUMBER(B5))),"",A5*100+B5)`
- Dates B5 `=IF(OR(A5="",NOT(ISNUMBER(A5))),"",MONTH(A5)*100+DAY(A5))` (core `=MONTH(date)*100+DAY(date)`)
- Dates C5 `=IF(OR(B5="",Seasons!$I$7<>"OK"),"",IF(B5<MIN(Seasons!$E$5:$E$16),INDEX(Seasons!$C$5:$C$16,COUNT(Seasons!$E$5:$E$16)),INDEX(Seasons!$C$5:$C$16,MATCH(B5,Seasons!$E$5:$E$16,1))))` (core: the topic's C5); D5 the same on Seasons!$D$5:$D$16
- Seasons F6 `=IF(AND(A6="",B6="",TRIM(C6)="",D6=""),"",IF(OR(NOT(ISNUMBER(A6)),NOT(ISNUMBER(B6))),"Check: month or day blank or text",IF(OR(A6<1,A6>12,A6<>INT(A6),B6<1,B6<>INT(B6)),"Check: month 1-12, day a whole number",IF(B6>DAY(DATE(2024,A6+1,0)),"Check: day not valid for the month",IF(TRIM(C6)="","Check: no season name",IF(NOT(ISNUMBER(D6)),"Check: rate blank or text",IF(COUNTIF($E$5:$E$16,E6)>1,"Check: same start twice",IF(E5="","Check: row above is blank or not valid",IF(E6<E5,"Check: sort by start","OK")))))))))` (F5 has no order part)
- Seasons I7 `=IF(AND(I4>0,I5=0,I6=0),"OK","Check")`; Dates E5 `=IF(A5="","",IF(NOT(ISNUMBER(A5)),"Check: date is text",IF(Seasons!$I$7<>"OK","Check: fix the Seasons tab","OK")))`

Sample values (fictional; recalculated). Seasons: 1/6 Off-season $129.00 (key 106), 4/1 Spring $159.00 (401), 6/15 Peak $219.00 (615), 9/2 Shoulder $169.00 (902), 11/1 Off-season $129.00 (1101), 12/15 Holiday $249.00 (1215); all OK, I4:I7 6, 0, 0, OK.
| Date | Key | Season | Rate |
|---|---|---|---|
| 01/02/2026 | 102 | Holiday | $249.00 |
| 01/05/2026 | 105 | Holiday | $249.00 |
| 01/06/2026 | 106 | Off-season | $129.00 |
| 04/01/2026 | 401 | Spring | $159.00 |
| 06/14/2026 | 614 | Spring | $159.00 |
| 06/15/2026 | 615 | Peak | $219.00 |
| 09/02/2026 | 902 | Shoulder | $169.00 |
| 12/14/2026 | 1214 | Off-season | $129.00 |
| 12/15/2026 | 1215 | Holiday | $249.00 |
| 12/31/2026 | 1231 | Holiday | $249.00 |
| 01/01/2027 | 101 | Holiday | $249.00 |
| 07/04/2027 | 704 | Peak | $219.00 |
Other rows: 03/31/2026 Off-season, 09/01/2026 Peak, 10/31/2026 Shoulder, 11/01/2026 Off-season, 02/29/2028 Off-season. H4:H9: 17, $3,203.00, 0, OK, 0, OK.

Edge cases tested (all matched the Python model)
- 400 random dates 2024 to 2027 and all 366 days of 2028. Seasons starting 01/01 (no wrap). One season (every date). 12 seasons on the 1st; 12 seasons on the 10th (01/01 to 01/09 wrap to the 12/10 season).
- Same start twice: both Check: same start twice, Dates blank (Check: fix the Seasons tab). 02/29 as a start: 02/29/2028 Leap, 02/28/2027 Off-season, 03/01 March. 4/31: Check: day not valid for the month. Month 13, month 0, day 1.5, "4" as text, no name, no rate, "99" as text rate: Check.
- Text dates "12/20/2026", " " and "July 4": Check: date is text. Blank Dates rows in the middle: skipped. Unsorted list: 3 rows Check: sort by start, Dates blank. Blank row inside the Seasons list: Check: row above is blank or not valid. Rows on Seasons row 17 or Dates row 505: CHECK Check. 6 random sorted sets of 2 to 12 starts.
- Probe5 (copy with the guard removed): unsorted list gave 06/14 and 07/04 Off-season, 12/31 Shoulder, 01/06 #N/A; LOOKUP without the wrap gave #N/A for 01/02, 01/05, 01/01/2027; the same-year AND test gave "not holiday" for all 17 dates.

Pitfalls the post must state
- One date per row: a stay that crosses into another season is priced one night per row (link rental-price-weekday-weekend-day-rates and split-booking-days-across-months).
- The starts must be sorted; the template blanks the Dates tab until they are.
- Seasons and rates are the owner's own; no holiday calendar, no lodging tax or rental law.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- #N/A for 01/01 to 01/05: MATCH or LOOKUP without the IF that sends dates before the first start to the last season.
- A between-two-dates test says "not holiday" for every date: no date is after Dec 15 and before Jan 5 of the same year.
- Wrong seasons (06/14 Off-season): the Seasons list is not sorted.
- A date shows Check: date is text: it was typed or pasted as text.

Hero: sheet "Dates", rows 1 to 21 (a copy with rows 22:504 deleted), hide_cols "" (A:H show); Seasons A4:F10 as an inset.

---
Not built or tested (so not for the posts): recalculation time for 4,095 rows in any app, more than 12 invoices, which invoices a client meant, bookkeeping of short payments or fees, any print lab's or marketplace's DPI rule, where a crop falls, file metadata DPI or export dialogs, recommended maintenance intervals, warranty or safety, tax mileage, Excel's or Sheets' criteria text matching and blank criteria rows, date criteria under other locales, pivot tables, multi-night stays priced in one row, holiday calendars, lodging tax, typing by hand. Nothing here was run in Excel or Google Sheets.
