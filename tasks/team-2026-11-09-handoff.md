# Team handoff - 2026-11-09 (hook-writer + template-builder to script-writer)

Post files: posts/ready/2026-11-09-team-0N-<slug>.md. Hooks (titles and openings) are in tasks/hooks-2026-11-09.md (openings 4 and 5 corrected to the built formulas); test log in tasks/briefs.md (2026-11-09 builder results). Build: scripts/build-templates.py (functions `text_duration_to_minutes`, `fiscal_year_quarter`, `monthly_running_total`, `shade_groups`, `compare_price_lists`). All values below were recalculated in LibreOffice Calc 24.2.7.2 (headless) and matched independent Python models, except where a pitfall says otherwise. Excel and Google Sheets were NOT tested; `tested_in` must say LibreOffice only. Every workbook has a "How to use" tab, yellow = input cells, gray = formula cells (topic 4 is the exception, see its pitfalls). No topic was changed from tasks/team-topics-2026-11-09.md. Deviations from the topic file are listed per post.

Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays. Fictional data. The fiscal-year post must not give tax advice; the price post must not say what to charge.

Final template files: templates/tidy-tabs-text-duration-to-minutes.xlsx, tidy-tabs-fiscal-year-quarter.xlsx, tidy-tabs-monthly-running-total.xlsx, tidy-tabs-shade-groups.xlsx, tidy-tabs-compare-price-lists.xlsx.

---
## 1. convert-text-hours-minutes-1h-30m-to-minutes-find-mid-excel-google-sheets
Title: How to Convert Text Like 2h 30m Into Minutes in Excel and Google Sheets
Template: templates/tidy-tabs-text-duration-to-minutes.xlsx (tabs: Entries, How to use)

Key cells
- Entries: B1 hourly rate (input, 65.00), header row 3, data rows 4:103 (100 slots, 12 sample rows 4:15). A Date, B Time as typed (inputs, formatted as text), C Cleaned, D Hours, E Minutes, F Total minutes, G Decimal hours, H Amount, I Check.
- Totals block K3:L8: L3 entries, L4 total minutes, L5 total decimal hours, L6 amount, L7 entries showing Check, L8 rate check (B1 is a number).

Sample data (B4:B15, dates 11/02/2026 to 11/13/2026): "2h 30m", "45m", "1h", "1h15m", "3H 5M", " 90m " (spaces), "2h 0m", "1.5h", (blank), "30 min", "n/a", "0h 20m".

Formulas (row 4)
- C4 `=LOWER(SUBSTITUTE(TRIM(B4)," ",""))`
- D4 `=IF(ISNUMBER(FIND("h",C4)),VALUE(LEFT(C4,FIND("h",C4)-1)),0)`
- E4 `=IF(ISNUMBER(FIND("m",C4)),VALUE(MID(C4,IF(ISNUMBER(FIND("h",C4)),FIND("h",C4)+1,1),FIND("m",C4)-IF(ISNUMBER(FIND("h",C4)),FIND("h",C4)+1,1))),0)`
- F4 `=IF(C4="","",IF(ISERROR(D4+E4),"Check",IF(OR(D4<0,E4<0,ISNUMBER(VALUE(RIGHT(C4,1)))),"Check",ROUND(D4*60+E4,0))))` (the core `ROUND(D4*60+E4,0)` is what the post teaches; the wrapper is the guard)
- G4 `=IF(ISNUMBER(F4),F4/60,"")`; H4 `=IF(ISNUMBER(F4),IF(ISNUMBER($B$1),ROUND(G4*$B$1,2),"Check"),"")`; I4 `=IF(C4="","",IF(ISNUMBER(F4),IF(F4<=0,"Check",""),"Check"))`
- L3 `=SUMPRODUCT(--(C4:C103<>""))`, L4 `=SUM(F4:F103)`, L5 `=SUM(G4:G103)`, L6 `=SUM(H4:H103)`, L7 `=COUNTIF(I4:I103,"Check")`.
- D4:E103 have a conditional format that hides the 0 (gray text on gray) when C is blank.

Verified results (sample)
| Typed | Minutes | Decimal hours | Amount at $65.00 |
|---|---|---|---|
| 2h 30m | 150 | 2.50 | $162.50 |
| 45m | 45 | 0.75 | $48.75 |
| 1h | 60 | 1.00 | $65.00 |
| 1h15m | 75 | 1.25 | $81.25 |
| 3H 5M | 185 | 3.08 | $200.42 |
| " 90m " | 90 | 1.50 | $97.50 |
| 2h 0m | 120 | 2.00 | $130.00 |
| 1.5h | 90 | 1.50 | $97.50 |
| (blank) | blank | blank | blank |
| 30 min | 30 | 0.50 | $32.50 |
| n/a | 0 (Check) | 0.00 | $0.00 |
| 0h 20m | 20 | 0.33 | $21.67 |
Totals: 11 entries, 865 minutes, 14.42 hours (14.4167), $937.09, 1 entry showing Check (n/a).

Pitfalls
- Deviation from the topic: F4 and I4 carry extra guards (a number-only entry such as 5 or an entry ending in a digit such as "2h 30" shows Check instead of silently using 0 or dropping the 30), the totals skip those rows, and H shows Check when B1 is not a number. D4 and E4 are the topic's formulas unchanged.
- D or E shows #VALUE! (while F and I say Check) for entries such as "h30m", "30m2h", "1 hour 30 minutes", "M", "h", "hm", "1h m", "m30", "1.5.5h", "1,5h". That is how the sheet signals an unreadable entry; say so in the post. "2:30", "n/a", "abc", "0m", "0h" give 0 minutes and Check; "5" and "-30m" give Check.
- Lenient accepts that are checked and true in LibreOffice: "2 hours" = 120, "30mm" = 30, "1hh" = 60 (text after the h or m is ignored when nothing else follows). "1 hour 30 minutes" is NOT read (Check). Do not promise words.
- Non-breaking spaces: in LibreOffice a cell with a non-breaking space ("2<nbsp>h 30m", "2h<nbsp>30m", " 1h ") was read correctly (150, 150, 60) because LibreOffice's VALUE ignores it; we did NOT add SUBSTITUTE(...,UNICHAR(160),"") and Excel was not run, so the post must not claim Excel handles it. If someone's paste shows Check, wrap B4 in SUBSTITUTE(B4,UNICHAR(160)," ") (Excel 2013 and later, Sheets). A tab character is not removed (Check).
- A cell holding only a non-breaking space counts as an entry and shows Check.
- Amounts round each row to cents (ROUND(G*rate,2)), so the total is the sum of rounded rows. 0.075h = 4.5 minutes rounded to 5 matched the exact-decimal model.
- Date column A is optional and never used by a formula. Slots: 100 rows.

---
## 2. fiscal-year-quarter-from-date-custom-start-month-excel-google-sheets
Title: How to Get the Fiscal Year and Quarter From a Date in Excel and Google Sheets
Template: templates/tidy-tabs-fiscal-year-quarter.xlsx (tabs: Sales, Summary, How to use)

Key cells
- Sales: B1 first month of the year (input, 7, data validation whole number 1 to 12), B2 setting check; header row 3; A Date, B Amount (inputs), C Fiscal year, D Fiscal quarter, E Label, F Row check in rows 4:103 (100 slots, 18 sample rows 4:21, sorted by date).
- Summary: A4:A15 quarter labels (inputs, 6 typed), B Total, C Sales count; side block F3:G7: G3 total of all sales, G4 total of quarters listed, G5 not in a listed quarter (must be 0), G6 Sales rows showing Check, G7 CHECK.

Sample data (date, amount): 06/15/2026 140.00; 07/08/2026 85.50; 08/15/2026 120.00; 09/30/2026 64.25; 10/01/2026 210.00; 11/09/2026 47.25; 11/23/2026 88.00; 12/31/2026 150.00; 01/01/2027 38.75; 01/05/2027 95.00; 02/14/2027 72.00; 03/31/2027 180.50; 04/01/2027 56.00; 05/20/2027 133.00; 06/30/2027 99.99; 06/30/2027 25.00; 07/01/2027 110.00; 07/15/2027 41.00.

Formulas (row 4)
- B2 `=IF(ISNUMBER(B1),IF(AND(B1>=1,B1<=12,B1=INT(B1)),"OK","Check"),"Check")`
- C4 `=IF(A4="","",IF(OR($B$2<>"OK",NOT(ISNUMBER(A4))),"Check",YEAR(A4)+IF(AND($B$1>1,MONTH(A4)>=$B$1),1,0)))` (core: `=YEAR(A4)+IF(AND($B$1>1,MONTH(A4)>=$B$1),1,0)`)
- D4 `=IF(ISNUMBER(C4),INT(MOD(MONTH(A4)-$B$1,12)/3)+1,C4)` (core: `=INT(MOD(MONTH(A4)-$B$1,12)/3)+1`)
- E4 `=IF(ISNUMBER(C4),"FY"&C4&" Q"&D4,C4)`; F4 `=IF(COUNTA(A4:B4)=0,"",IF(AND(ISNUMBER(A4),OR(B4="",ISNUMBER(B4))),"OK","Check"))`
- Summary B4 `=IF(A4="","",SUMIFS(Sales!$B$4:$B$103,Sales!$E$4:$E$103,A4))` (the topic had $B$4:$B$21; the file runs to row 103), C4 `=IF(A4="","",COUNTIFS(Sales!$E$4:$E$103,A4))`; G5 `=ROUND(G3-G4,2)`, G7 `=IF(AND(Sales!B2="OK",G5=0,G6=0),"OK","Check")`.

Verified results (sample, start month 7)
| Date | FY | Quarter | Label |
|---|---|---|---|
| 06/15/2026 | 2026 | 4 | FY2026 Q4 |
| 08/15/2026 | 2027 | 1 | FY2027 Q1 |
| 11/09/2026 | 2027 | 2 | FY2027 Q2 |
| 01/05/2027 | 2027 | 3 | FY2027 Q3 |
| 06/30/2027 | 2027 | 4 | FY2027 Q4 |
| 07/01/2027 | 2028 | 1 | FY2028 Q1 |
Summary: FY2026 Q4 $140.00 (1 sale), FY2027 Q1 $269.75 (3), Q2 $495.25 (4), Q3 $386.25 (4), Q4 $313.99 (4), FY2028 Q1 $151.00 (2); total $1,756.24; not in a listed quarter $0.00; CHECK OK.

Pitfalls
- The FY is named for the calendar year it ENDS in (start month 1 gives FY = calendar year). Say it is a choice; say it labels dates and is not tax advice.
- Start month 0, 13, 7.5, text or blank: B2, every row and the Summary CHECK show Check (Excel and Sheets also get the validation dropdown rule; validation does not stop a paste).
- Summary labels are typed text; if the start month changes, the labels in column A must be retyped (the check G5 turns non-zero).
- A date typed as text shows Check; a date with a time works. Blank rows stay blank. A text amount shows Check in F and G6.
- Slots: 100 sales, 12 summary rows. Extra topic row: a sample sale on 11/23/2026 was added so Q2 has 4 sales.

---
## 3. running-total-that-restarts-each-month-sumifs-excel-google-sheets
Title: How to Make a Running Total That Restarts Each Month in Excel and Google Sheets
Template: templates/tidy-tabs-monthly-running-total.xlsx (tabs: Expenses, How to use)

Key cells
- Expenses: B1 monthly limit (input, 450.00), header row 3, rows 4:63 (60 slots, 8 sample rows 4:11): A Date, B What, C Amount (inputs), D Month to date, E Left this month, F Month total, G Row check. Side block I3:J6: J3 total of all amounts, J4 rows showing Check or Out of order, J5 rows over the limit, J6 CHECK.

Sample data (2026): 10/03 Supplies 120.00; 10/12 Booth fee 85.00; 10/27 Shipping 64.50; 11/02 Supplies 210.00; 11/09 Packaging 47.25; 11/18 Ads 150.00; 11/30 Shipping 38.75; 12/04 Supplies 95.00.

Formulas (row 4)
- D4 `=IF(A4="","",IF(ISNUMBER(A4),SUMIFS($C$4:C4,$A$4:A4,">="&DATE(YEAR(A4),MONTH(A4),1)),"Check"))` (core as in the hooks file)
- E4 `=IF(ISNUMBER(D4),IF(ISNUMBER($B$1),$B$1-D4,"Check"),D4)`
- F4 `=IF(A4="","",IF(ISNUMBER(A4),SUMIFS($C$4:$C$63,$A$4:$A$63,">="&DATE(YEAR(A4),MONTH(A4),1),$A$4:$A$63,"<"&EDATE(DATE(YEAR(A4),MONTH(A4),1),1)),"Check"))`
- G4 `=IF(COUNTA(A4:C4)=0,"",IF(AND(ISNUMBER(A4),OR(C4="",ISNUMBER(C4))),IF(A4<MAX($A$3:A3),"Out of order","OK"),"Check"))` (A3 is the header text, MAX ignores it)

Verified results (sample)
| Date | Amount | Month to date | Left | Month total |
|---|---|---|---|---|
| 10/03/2026 | 120.00 | 120.00 | 330.00 | 269.50 |
| 10/12/2026 | 85.00 | 205.00 | 245.00 | 269.50 |
| 10/27/2026 | 64.50 | 269.50 | 180.50 | 269.50 |
| 11/02/2026 | 210.00 | 210.00 | 240.00 | 446.00 |
| 11/09/2026 | 47.25 | 257.25 | 192.75 | 446.00 |
| 11/18/2026 | 150.00 | 407.25 | 42.75 | 446.00 |
| 11/30/2026 | 38.75 | 446.00 | 4.00 | 446.00 |
| 12/04/2026 | 95.00 | 95.00 | 355.00 | 95.00 |
J3 = 810.50, J4 = 0, J5 = 0, J6 = OK. A plain running total on 11/18/2026 would be 676.75 (120 + 85 + 64.50 + 210 + 47.25 + 150).

Pitfalls
- Unsorted rows: D is wrong (it only adds rows above), F is still right, G says Out of order. Test: 11/20 100, 10/05 50, 11/02 10, 12/03 1, 11/25 5 gave D = 100, 150, 110, 1, 116 where the true month to date by date is 110, 50, 10, 1, 115; F = 115, 50, 115, 1, 115; G flags the three rows dated before an earlier row. The fix is Data > Sort range by column A.
- A text amount (abc) is ignored by SUMIFS and shows Check in G; a text date shows Check in D, E, F, G; a refund is a negative amount and lowers the total. Limit blank or text: E shows Check; limit 0 shows negative Left (red).
- Rows with a time of day work (10/31 vs 11/01 boundaries tested). Blank rows in the middle stay blank and do not break later rows. Month and year boundaries (10/31, 11/01, 12/31, 01/01, 02/29/2028) matched.
- Slots: 60 rows (the F ranges end at row 63; adding rows below needs the ranges extended).

---
## 4. alternate-row-colors-by-group-order-number-conditional-formatting-excel-google-sheets
Title: How to Shade Alternating Groups of Rows by Order Number in Excel and Google Sheets
Template: templates/tidy-tabs-shade-groups.xlsx (tabs: Lines, Plain banding, How to use)

Key cells
- Lines: header row 3, rows 4:63 (60 slots, 24 sample rows 4:27), A Order, B Item, C Qty (inputs, no fill on purpose), D Group no. (formula, no fill). Conditional formatting on A4:D63.
- Plain banding (extra tab, deviation from the topic): the same 24 lines copied by formula in A4:C27 with the plain rule `=MOD(ROW(),2)=0` for contrast.

Sample data (order, item, qty): 5001 Soy candle, lavender 2 / Lip balm 4 / Gift box 1; 5002 Tote bag 1; 5003 Enamel pin 3 / Sticker pack 6 / Soy candle, cedar 1 / Gift box 1; 5004 Lip balm 2 / Tote bag 1; 5005 Sticker pack 10 / Enamel pin 2 / Lip balm 3; 5006 Gift box 2; 5007 Soy candle, lavender 1 / Soy candle, cedar 1 / Tote bag 2 / Sticker pack 4; 5008 Enamel pin 5 / Lip balm 1; 5009 Tote bag 1 / Gift box 1 / Lip balm 6 / Sticker pack 2.

Formulas
- D4 `=IF(A4="","",1)`; D5 `=IF(A5="","",IF(A5=A4,D4,MAX($D$4:D4)+1))` filled to D63 (the topic's plain `=IF(A5=A4,D4,D4+1)` also gives the same numbers on the sample but would number the empty rows below the data).
- Rule on A4:D63, custom formula `=MOD($D4,2)=0`, fill D9D9D9 (light gray).
- Helper-free rule, A4:D63: `=AND($A4<>"",MOD(SUMPRODUCT(--($A$4:$A4<>$A$3:$A3)),2)=0)`. The bare `=MOD(SUMPRODUCT(--($A$4:$A4<>$A$3:$A3)),2)=0` from the topic also matched on the sample, but the AND($A4<>"") guard is what the file's notes use so blank rows are never shaded.
- Plain contrast rule: `=MOD(ROW(),2)=0`.

Verified results
- D4:D27 = 1,1,1,2,3,3,3,3,4,4,5,5,5,6,7,7,7,7,8,8,9,9,9,9; D28 and below blank. Shaded (even groups): rows 7 (5002), 12 and 13 (5004), 17 (5006), 22 and 23 (5008), as in the PDF render.
- Rendered the xlsx to PDF (LibreOffice 24.2.7.2) and then to PNG with PyMuPDF, looked at the page: the shaded blocks follow the orders; the Plain banding tab shades every other row and cuts through 5001, 5003, 5005, 5007 and 5009. Pixel check of the 24 rows: helper rule and helper-free rule shade the same rows (sample), and the file survives a LibreOffice save and reopen with the rule intact (the cfRule expression `MOD($D4,2)=0` is in the saved sheet XML).
- Screenshots exist only in the scratchpad, not in the repo; the post needs its own real captured screenshots with the app and version in the caption.

Pitfalls
- The list must be sorted by order number. The same order in two places gets two groups (test 5003, 5001, 5001, 5003, 5002, 5003 gave 1, 2, 2, 3, 4, 5).
- A blank row between lines starts a new group in the helper version (one new group). The helper-free rule counts a blank row as two changes, so the shades after a blank row can differ from the helper version (checked by rendering both). Without the AND($A4<>"") guard the blank row itself can be shaded.
- Order numbers typed as text vs number: 5001 and "5001" are different groups. Capitals do not matter (a and A are the same group).
- Deviation: the yellow and gray static fills are left off this file because a static fill would hide the shading; column D is gray text only. Say that the shading comes from the rule, and where it sits (Home > Conditional Formatting > Manage Rules in Excel; Google Sheets steps come from Google help and were not run by us).
- A real row inserted inside the range, 200 rows, and Excel or Sheets rendering were NOT tested.
- Slots: 60 rows, the print is 1 page.

---
## 5. compare-old-and-new-price-lists-change-percent-index-match-excel-google-sheets
Title: How to Compare Two Price Lists and See What Changed in Excel and Google Sheets
Template: templates/tidy-tabs-compare-price-lists.xlsx (tabs: Old, New, Compare, How to use)

Key cells
- Old and New: A SKU, B Price (inputs, rows 2:41, 8 sample rows 2:9), C SKU used `=IF(TRIM(A2)="","",TRIM(A2))` (helper).
- Compare: counts in A2:G2 under the labels in row 1 (A Raised, B Lowered, C Same, D New, E Removed, F Check, G CHECK); header row 3; rows 4:43 follow the New tab: A SKU, B New price, C Old price, D Change, E Percent, F Status, G Row check. Removed block beside it: I Old SKU, J Old price, K Removed? in rows 4:43 (follows the Old tab).

Sample data
Old: CND-8 18.00; LIP-1 3.50; TOT-1 24.00; PIN-1 8.00; STK-4 4.50; BOX-1 32.00; CRD-1 5.00; MUG-1 16.00. New: CND-8 19.50; LIP-1 3.50; TOT-1 24.00; PIN-1 7.50; STK-4 4.50; BOX-1 34.00; SOAP-1 9.00; MUG-1 16.00.

Formulas (row 4)
- A4 `=IF(New!C2="","",New!C2)`; B4 `=IF(A4="","",New!B2)`
- C4 `=IF(A4="","",IFERROR(INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0)),""))`
- D4 `=IF(OR(NOT(ISNUMBER(B4)),NOT(ISNUMBER(C4))),"",ROUND(B4-C4,2))`; E4 `=IF(D4="","",IF(C4=0,"",D4/C4))`
- F4 `=IF(A4="","",IF(G4<>"OK","Check",IF(ISNA(MATCH(A4,Old!$C$2:$C$41,0)),"New",IF(D4>0,"Raised",IF(D4<0,"Lowered","Same")))))`
- G4 `=IF(A4="","",IF(NOT(ISNUMBER(B4)),"Check new price",IF(ISNA(MATCH(A4,Old!$C$2:$C$41,0)),"OK",IF(SUMPRODUCT(--(Old!$C$2:$C$41=A4))>1,"Duplicate in Old",IF(ISNUMBER(INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0))),"OK","Check old price")))))`
- I4 `=IF(Old!C2="","",Old!C2)`; J4 `=IF(I4="","",Old!B2)`; K4 `=IF(I4="","",IF(ISNUMBER(MATCH(I4,New!$C$2:$C$41,0)),"","Removed"))`
- Counts `=COUNTIF($F$4:$F$43,"Raised")` etc., E2 `=COUNTIF($K$4:$K$43,"Removed")`, G2 `=IF(SUM(A2:D2)+F2=COUNTIF($A$4:$A$43,"?*"),"OK","Check")`.

Deviations from the topic: ranges run to row 41 (40 products) instead of row 9, so the formula in the post is `INDEX(Old!$B$2:$B$41,MATCH(A4,Old!$C$2:$C$41,0))` (hooks file corrected); the helper column C on each list trims the SKU (the topic only said to document it); the Removed block sits beside the Compare table, not below it.

Verified results (sample)
| SKU | New | Old | Change | Percent | Status |
|---|---|---|---|---|---|
| CND-8 | $19.50 | $18.00 | $1.50 | 8.3% | Raised |
| LIP-1 | $3.50 | $3.50 | $0.00 | 0.0% | Same |
| TOT-1 | $24.00 | $24.00 | $0.00 | 0.0% | Same |
| PIN-1 | $7.50 | $8.00 | -$0.50 | -6.3% | Lowered |
| STK-4 | $4.50 | $4.50 | $0.00 | 0.0% | Same |
| BOX-1 | $34.00 | $32.00 | $2.00 | 6.3% | Raised |
| SOAP-1 | $9.00 | blank | blank | blank | New |
| MUG-1 | $16.00 | $16.00 | $0.00 | 0.0% | Same |
CRD-1 shows Removed. Counts: Raised 2, Lowered 1, Same 4, New 1, Removed 1, Check 0, CHECK OK. Percent -6.25% displays as -6.3% and 6.25% as 6.3% with the 0.0% format (say "6.3%" only if the post uses that format).

Pitfalls
- MATCH ignores capitals (cnd-8 matches CND-8). TRIM removes outer spaces and collapses inner doubles on both lists; a SKU typed as the number 1001 on one list and "1001" on the other matches (both go through TRIM). Non-breaking spaces and tabs are not removed.
- MATCH reads * ? ~ as wildcards in a SKU; not escaped here (the 11-08 state file escapes them). Say SKUs with those characters are not supported.
- A SKU twice in Old: the first price is used and Status shows Check (Row check says Duplicate in Old). A SKU twice in New: both rows are compared with the same old price. Old price 0: Percent blank (no #DIV/0!), status from Change. Old price blank or text, or New price text or blank: Check (Row check says which). A blank Old price is NOT treated as 0.
- Prices with half a cent: 1.005 against 1.00 gave Same (floating point: ROUND of -0.00499999999999989 is 0) where an exact-decimal model says Lowered by 0.01. Real prices have two decimals; the post can skip it.
- Blank rows in the middle of a list stay blank in the Compare row but still take up a row (the New tab row number is the Compare row minus 2).
- A SKU removed from the old list and a new SKU in the same row do not pair up; they show New and Removed separately.
- Slots: 40 products per list. The sheet shows what changed; the post must not say what to charge.
