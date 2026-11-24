# Handoff 2026-11-24: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-24.md (researcher) and tasks/hooks-2026-11-24.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`hourly_rental_daily_max`, `cheapest_pack_mix` with `PACK_MIX_MAX`, `ab_test_significance` with `AB_LIMIT`, `replace_words_from_list` with `REPLACE_PAIRS` and `REPLACE_TEXT_ROWS`, `find_transposed_digits`, registered after `season_from_date()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-24 build and test"). Heroes are not made yet (designer, see each section; trial renders were made in scratch only).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 243 edited workbooks recalculated and compared (2,567,738 cell comparisons, 0 mismatches against independent Python models, 0 error values in any of the 243; a mutation run with one wrong formula per template produced mismatches in all five, so the harness catches errors), plus 12 probe workbooks of the traps read by hand. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values, Err: or ### in the PDF text. No LibreOffice behavior differed from the researcher's probe; three topic expectations changed in the built files (topic 2 tie rule added, topic 3 "Over 1,000,000" instead of 7,052,685, topic 5 "Extra entry?" for the $1,172.04 case), and nine layout problems were found and fixed (tasks/briefs.md).

Repo note: the whole `python3 scripts/build-templates.py` was not run (it rewrites older templates byte-wise); the five files were built by importing the module and calling only the five new functions. `git status`: scripts/build-templates.py (590 lines added, 0 removed), tasks/briefs.md, this handoff, five new .xlsx.

Post file names: `posts/ready/2026-11-24-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks every status starting "Check" and any CHECK or settings cell that is not OK; light blue marks "Capped" and light orange "Minimum" (Bookings); light green marks the chosen mix (Mixes), the cheapest total (Order), "Likely a real difference" (Tests) and "Yes" (Text Changed?); light orange marks "Caution" statuses (Tests, Replacements) and any listed candidate (Entries K:N). Functions used in the five files: ABS, AND, CEILING, COLUMNS, COUNT, COUNTA, COUNTIF, EXACT, FIND, IF, INDEX, INT, ISNUMBER, ISTEXT, LARGE, LEFT, LEN, MATCH, MAX, MID, MIN, MOD, NORMSDIST, NORMSINV, NOT, OR, ROUND, ROUNDUP, ROWS, SIGN, SMALL, SQRT, SUBSTITUTE, SUM, SUMPRODUCT, TEXT, TRIM. No TODAY, FILTER, UNIQUE, SORT, LET, LAMBDA, REDUCE, XLOOKUP, TEXTSPLIT, MINIFS/MAXIFS; 0 `_xlfn` in the files. Help links were not checked by the builder.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-24-team-01-hourly-rental-fee-with-daily-maximum-and-minimum-hours-int-ceiling-min-excel-google-sheets | tidy-tabs-hourly-rental-daily-max.xlsx | 36 workbooks, 43,488 comparisons |
| 2 | 2026-11-24-team-02-cheapest-mix-of-pack-sizes-to-buy-at-least-a-quantity-int-mod-roundup-min-excel-google-sheets | tidy-tabs-cheapest-pack-mix.xlsx | 96 workbooks, 1,923,648 comparisons |
| 3 | 2026-11-24-team-03-ab-test-two-listing-versions-conversion-rate-significance-normsdist-sqrt-excel-google-sheets | tidy-tabs-ab-test-significance.xlsx | 23 workbooks, 11,684 comparisons |
| 4 | 2026-11-24-team-04-find-and-replace-multiple-words-from-a-list-whole-word-substitute-trim-excel-google-sheets | tidy-tabs-replace-words-from-list.xlsx | 22 workbooks, 159,060 comparisons |
| 5 | 2026-11-24-team-05-find-transposed-digits-when-totals-dont-match-divisible-by-9-mid-mod-excel-google-sheets | tidy-tabs-find-transposed-digits.xlsx | 66 workbooks, 429,858 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets." For the rental post add: "The raw-hours rounding error was measured in LibreOffice Calc only (3,000 random bookings)." For the A/B post add: "NORMSDIST and NORMSINV results were checked against Python's math.erf and NormalDist only."

Hooks check (tasks/hooks-2026-11-24.md, every "to confirm by builder" number):
- 1 chosen: confirmed. 11/06/2026 9:00 AM to 11/08/2026 11:15 AM = 3,015 minutes, 2 blocks + 135 minutes, 2.5 hours, $487.50; the web formula gave $600.00 on a template copy. The four formulas match the file's cores; the cell adds the grace setting: G8 `=IF(F8<=$B$4,0,CEILING(F8,$B$3)/60)` and H8 wraps the charge in ROUND(...,2). Opening 2: confirmed (9-hour booking $200.00 instead of $315.00, Capped; 1-hour booking $70.00, Minimum). Opening 3: confirmed (`=ROUNDUP((end-start)*24*2,0)/2` gave 1.5 hours for 10:00 to 11:00 AM; the raw hours read back as 1.00000000046556). The CEILING-on-raw-hours version is now measured: `=CEILING((end-start)*24,0.5)` also gave 1.5 hours for 10:00 to 11:00 AM and was wrong on 13 of the researcher's 3,000 random bookings (ROUNDUP on raw hours 12 of 3,000, ROUND-to-minutes 0). The post may state both, as LibreOffice Calc results.
- 2 chosen: confirmed (95 jars: two 48-cases $104.00, 1 extra; biggest-first one 48-case + three 12-cases + 11 singles $116.80; $12.80 more). Opening 2: confirmed. Opening 3: confirmed (10 jars: one 12-case $15.00, 2 extra; biggest-first 10 singles $18.00). Claims note correction: the template has a tie rule (lowest cost, then fewest units bought, then fewest packs, then the first row), not "the first row MATCH finds"; state that rule.
- 3 chosen: confirmed (37 / 1,840 = 2.01% vs 58 / 1,910 = 3.04%, z 1.998 shows 2.00, p 0.0457 shows 0.046, Likely a real difference at 95%, Not enough evidence yet at 99%). Opening 2: confirmed (2.50% vs 3.05%, p 0.4132, Not enough evidence yet, 13,959 visits per version at 95% / 80%). Opening 3: the 104% rate, p 0 and #DIV/0! happen only in a copy with the Status guard removed (reproduced, probe p3: 0 orders in both gave #DIV/0! in z, p and Result; 520 on 500 gave rate 104%, p 0, "Likely a real difference" and an error in visits needed). In the template both rows show Check and no result. Write it as "without a check, ..." and say the template flags them.
- 4 chosen: confirmed ("Soy Candle Lg Lav w/ Wood Wick" -> "Soy Candle Large Lavender with Wood Wick"; "Smoky" unchanged). In the file the padding sits in its own helper column C5 `=IF(A5="","",SUBSTITUTE(" "&A5&" "," ","  "))` and the pair columns are D5:W5 (INDEX with COLUMNS($D5:D5)). Opening 2: confirmed, exact LibreOffice text with the sample's five probe pairs: "Smalloky Small Candle 8 oz" (with only Sm -> Small it would keep "oz."; say "turns Smoky into Smalloky"). Opening 3: confirmed ("Large Lg Gift Set" with single spaces, "Large Large Gift Set" with doubled spaces).
- 5 chosen: confirmed ($1,214.04 vs $1,223.04, gap 900 cents, four entries $78.50 (Maybe $87.50), $45.60 (Maybe $54.60), $112.80 (Maybe $121.80), $301.25 (Maybe $310.25)). The file's swap cell also multiplies by SIGN(B6) and blanks padding pairs. Opening 2: confirmed. Opening 3: confirmed ($989.04, gap -22,500 cents = -$225.00, no swap, slide Maybe $25.00 on $250.00); the file writes the slide test in whole numbers, `$B$3*10=-9*ROUND(B6*100,0)`, which is the opening's `gap=-0.9*cents` for a positive amount.

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=...)` fits the whole sheet, or its print area, on one landscape page and crops to the drawn cells; `short_sheet_copy(src, sheet, last_row, name)` hides the empty rows). Trial renders at 1600x900 were made in scratch (/tmp/claude-0/build-1124/hero/) under /usr/bin/python3 (the /usr/local python cannot import PIL here); the calls below rendered legibly. Bookings, Tests, Text and Entries fill gray formulas to rows 207, 55, 304 and 505, so a whole-sheet render without short_sheet_copy is tiny.

---
## 1. hourly-rental-fee-with-daily-maximum-and-minimum-hours-int-ceiling-min-excel-google-sheets
Title (hook-writer final, same as topics): How to Charge an Hourly Rental Rate With a Daily Maximum in Excel and Google Sheets
Template: templates/tidy-tabs-hourly-rental-daily-max.xlsx (tabs: Bookings, How to use)

Layout
- Bookings: settings A1:B5 (inputs B1:B5): A1 Hourly rate / B1 $35.00, A2 Maximum per 24 hours / B2 $200.00, A3 Billing step (minutes) / B3 30, A4 Grace minutes after full 24 hours / B4 10, A5 Minimum hours / B5 2; C1:C5 notes. Header row 7, rows 8:207: A Customer, B Start date-time, C End date-time (inputs, MM/DD/YYYY h:mm AM/PM); D Minutes, E Full 24-hour blocks, F Minutes left over, G Hours billed for the left over, H Charge, I Status. Summary K7:L14: Settings (B1:B5) (L7), Bookings listed (L8), Total charged (L9), Bookings capped (L10), Bookings at the minimum (L11), Rows showing Check (L12), Typed below row 207 (not read) (L13), CHECK (L14). Landscape, row 7 repeated.
- Deviations from the topic: the minimum is `MAX($B$5*$B$1, ...)` without IF(E8=0) (the settings check needs minimum hours x rate <= maximum, so it is the same rule); D:H stay blank unless start and end are numbers, end > start and L7 is OK; Status shows Capped / Minimum / OK / Check; H is rounded to cents.

Key formulas (exact text in the file)
- D8 `=IF(OR(NOT(ISNUMBER(B8)),NOT(ISNUMBER(C8)),$L$7<>"OK"),"",IF(C8<=B8,"",ROUND((C8-B8)*1440,0)))` (core `=ROUND((C8-B8)*1440,0)`)
- E8 `=IF(D8="","",INT(D8/1440))`; F8 `=IF(D8="","",D8-E8*1440)`
- G8 `=IF(D8="","",IF(F8<=$B$4,0,CEILING(F8,$B$3)/60))`
- H8 `=IF(D8="","",ROUND(MAX($B$5*$B$1,E8*$B$2+MIN(G8*$B$1,$B$2)),2))`
- I8 `=IF(AND(TRIM(A8)="",B8="",C8=""),"",IF(OR(NOT(ISNUMBER(B8)),NOT(ISNUMBER(C8))),"Check: start or end blank or text",IF(C8<=B8,"Check: end is not after start",IF(TRIM(A8)="","Check: no customer name",IF($L$7<>"OK","Check: settings (B1:B5)",IF(H8<ROUND((E8*24+G8)*$B$1,2),"Capped",IF(H8>ROUND(E8*$B$2+MIN(G8*$B$1,$B$2),2),"Minimum","OK")))))))`
- L7 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B4),ISNUMBER(B5)),IF(AND(B1>0,B2>0,B3>=1,B3<=1440,B3=INT(B3),B4>=0,B4<1440,B5>=0,B5*B1<=B2),"OK","Check"),"Check")`; L14 `=IF(AND(L7="OK",L8>0,L12=0,L13=0),"OK","Check")`

Sample values (fictional; recalculated). $35.00 an hour, $200.00 per 24 hours, 30-minute step, 10 grace minutes, 2-hour minimum.
| Customer | Start | End | Minutes | Blocks | Left over | Hours billed | Charge | Status |
|---|---|---|---|---|---|---|---|---|
| Rivera Family Photos | 11/02/2026 9:00 AM | 11/02/2026 1:30 PM | 270 | 0 | 270 | 4.5 | $157.50 | OK |
| Northside Bakery Co. | 11/03/2026 9:00 AM | 11/03/2026 1:40 PM | 280 | 0 | 280 | 5.0 | $175.00 | OK |
| Lena Park Headshots | 11/04/2026 8:00 AM | 11/04/2026 5:00 PM | 540 | 0 | 540 | 9.0 | $200.00 | Capped |
| Oak & Ember Catering | 11/05/2026 6:00 PM | 11/06/2026 10:00 AM | 960 | 0 | 960 | 16.0 | $200.00 | Capped |
| Tri-County Wedding Expo | 11/06/2026 9:00 AM | 11/08/2026 11:15 AM | 3,015 | 2 | 135 | 2.5 | $487.50 | Capped |
| J. Alvarez Studio | 11/09/2026 10:00 AM | 11/09/2026 11:00 AM | 60 | 0 | 60 | 1.0 | $70.00 | Minimum |
| Hillside Pop-up Market | 11/10/2026 9:00 AM | 11/11/2026 9:00 AM | 1,440 | 1 | 0 | 0.0 | $200.00 | Capped |
| Maya Chen Portraits | 11/12/2026 7:15 AM | 11/12/2026 1:25 PM | 370 | 0 | 370 | 6.5 | $200.00 | Capped |
| Green Fork Meal Prep | 11/13/2026 9:10 AM | 11/13/2026 1:40 PM | 270 | 0 | 270 | 4.5 | $157.50 | OK |
L7:L14: OK, 9, $1,847.50, 5, 1, 0, 0, OK.

Edge cases tested (all matched the Python model; booking start 11/02/2026 9:00 AM unless noted, sample settings)
| Booking | Charge |
|---|---|
| exactly 24 h | $200.00 (Capped) |
| 24 h 5 min, 24 h 10 min | $200.00 (grace) |
| 24 h 11 min, 24 h 15 min | $217.50 |
| 24 h 31 min | $235.00 |
| 47 h 59 min | $400.00 |
| 48 h | $400.00 |
| 12/31/2026 10:00 PM to 01/02/2027 1:00 AM | $305.00 (1 block + 3 h) |
| 11/30/2026 8:00 PM to 12/01/2026 9:30 AM | $200.00 |
| 5 minutes | $70.00 (Minimum) |
| date only 11/20 to 11/21 (midnight) | $200.00 |
- Step 15: 24 h 15 min $208.75; step 60: $235.00; grace 0: 24 h 5 min $217.50; minimum 0: 5 minutes $0.00. Also step 1, 45, 1440, minimum 1.5 h, a maximum below rate x step, rate $37.25 / max $245.50 / step 20, 02/28/2028 to 03/01/2028 ($435.00), 10 days 5 h 55 min ($2,200.00), 20 seconds (0 minutes, minimum), 1 h 40 s (61 minutes, 1.5 h, minimum $70.00).
- Check: end before start and end = start (Check: end is not after start), a text start and a blank end (Check: start or end blank or text), no customer (Check: no customer name; the charge still shows), settings (minimum charge above the maximum, rate as text, step 0, blank or 7.5, grace 1440: L7 Check and every row Check: settings), a row on row 208 (L13 1). 16 random sets of 150 to 200 bookings.

Pitfalls the post must state
- The blocks are 24 hours from the start time, not calendar days; a per-calendar-day cap is a different rule (one sentence, link rental-price-weekday-weekend-day-rates-networkdays).
- The rate, maximum, step, grace and minimum are the owner's settings; the samples are not advice. No deposits, damage, cancellation fees or lodging tax.
- Count whole minutes first. Raw hours ((end-start)*24) can land a hair above a whole hour in LibreOffice Calc; no claim about Excel's or Sheets' counts.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A 1-hour booking bills 1.5 hours: rounding up raw hours (ROUNDUP and CEILING on (end-start)*24 both did it for 10:00 to 11:00 AM; 12 and 13 of 3,000 random bookings).
- A 50-hour booking costs 3 full days ($600.00): the web formula `MIN(hours*rate,daily_max*ROUNDUP(hours/24,0))` caps per started day; the template's blocks + left over gives $487.50.
- Charge blank and Status Check: settings (B1:B5): step 0 or text, grace 1440 or more, or the minimum charge above the maximum.
- Charge blank: end not after start, or a start typed as text.

Hero: `render(short_sheet_copy(t+"tidy-tabs-hourly-rental-daily-max.xlsx","Bookings",16,"rent"),"Bookings",name="rent")` (A1:L16 incl. settings and summary; legible).

---
## 2. cheapest-mix-of-pack-sizes-to-buy-at-least-a-quantity-int-mod-roundup-min-excel-google-sheets
Title (hook-writer final, same as topics): How to Find the Cheapest Mix of Pack Sizes to Buy in Excel and Google Sheets
Template: templates/tidy-tabs-cheapest-pack-mix.xlsx (tabs: Order, Mixes, How to use)

Layout
- Order: A1 Units needed / B1 (input, 95), C1 and A2 notes. Header row 4, rows 5:7: A Pack name, B Units per pack, C Price per pack (inputs, any order, 2 or 3 packs); D Price per unit, E Status. A9 title, header row 10, rows 11:13 "Packs sorted by size": A Role (Large, Medium, Smallest (fills the rest)), B Units per pack, C Price per pack, D Counts tried, E Biggest-first count. Summary G4:H9: Packs listed (2 or 3) (H4), Pack rows showing Check (H5), Units needed (B1) (H6), Packs ready (H7), Mixes to try (2,000 at most) (H8), CHECK (H9). Result G10:H20: Cheapest total (H11), Large packs (H12), Medium packs (H13), Smallest packs (H14), Units bought (H15), Extra units (H16), Cost per unit needed (H17), Biggest-first total (H18), Biggest-first costs more by (H19), Mix # chosen (Mixes tab) (H20). Landscape.
- Mixes: A1:A3 notes, A4 "Lowest" with E4, H4, I4. Header row 5, rows 6:2005: A Mix #, B Large packs, C Medium packs, D Smallest packs, E Cost, F Units bought, G Packs in all, H Units if lowest cost, I Packs if lowest cost and fewest units, J Chosen (1). Print area A1:J32 on one page (27 mixes). Landscape, row 5 repeated.
- Deviations from the topic: the sheet sorts the packs (rows 11:13) so any order works, and 2 packs work (Medium row 0, D12 = 1); tie rule added (below); Order H11 reads Mixes!E4 (the topic's MIN); mixes over 2,000 give "Check: too many mixes" in H9 and no result.

Key formulas (exact text in the file)
- Order D11 `=IF(OR(B11="",$H$6<>"OK"),"",ROUNDUP($B$1/B11,0)+1)`; D12 `=IF(OR(B12="",$H$6<>"OK"),"",IF(B12=0,1,ROUNDUP($B$1/B12,0)+1))`
- Order B11 `=IF($H$7<>"OK","",LARGE($B$5:$B$7,1))`; B12 `=IF($H$7<>"OK","",IF($H$4=3,LARGE($B$5:$B$7,2),0))`; B13 `=IF($H$7<>"OK","",SMALL($B$5:$B$7,1))`; C11 `=IF(B11="","",INDEX($C$5:$C$7,MATCH(B11,$B$5:$B$7,0)))`
- Mixes A6 `=IF(OR(Order!$H$9<>"OK",ROWS(A$6:A6)>Order!$H$8),"",ROWS(A$6:A6)-1)`
- Mixes B6 `=IF($A6="","",INT($A6/Order!$D$12))`; C6 `=IF($A6="","",MOD($A6,Order!$D$12))` (core `=INT(mix/medium_options)`, `=MOD(mix,medium_options)`)
- Mixes D6 `=IF($A6="","",MAX(0,ROUNDUP((Order!$B$1-B6*Order!$B$11-C6*Order!$B$12)/Order!$B$13,0)))`
- Mixes E6 `=IF($A6="","",ROUND(B6*Order!$C$11+C6*Order!$C$12+D6*Order!$C$13,2))`; F6 units; G6 `=IF($A6="","",B6+C6+D6)`
- Tie rule: E4 `=MIN(E6:E2005)`; H6 `=IF($A6="","",IF(E6=$E$4,F6,""))`; H4 `=MIN(H6:H2005)`; I6 `=IF(H6="","",IF(H6=$H$4,G6,""))`; I4 `=MIN(I6:I2005)`; J6 `=IF(I6="","",IF(I6=$I$4,1,""))`
- Order H11 `=IF($H$9<>"OK","",Mixes!$E$4)`; H20 `=IF($H$9<>"OK","",INDEX(Mixes!$A$6:$A$2005,MATCH(1,Mixes!$J$6:$J$2005,0)))`; H12 `=IF($H$9<>"OK","",INDEX(Mixes!$B$6:$B$2005,$H$20+1))` (H13:H15 the same on C, D, F)
- Biggest-first: E11 `=IF(OR(B11="",$H$6<>"OK"),"",INT($B$1/B11))`; E12 `=IF(OR(B12="",$H$6<>"OK"),"",IF(B12=0,0,INT(($B$1-E11*B11)/B12)))`; E13 `=IF(OR(B13="",$H$6<>"OK"),"",MAX(0,ROUNDUP(($B$1-E11*B11-E12*B12)/B13,0)))`; H18 `=IF($H$9<>"OK","",ROUND(E11*C11+E12*C12+E13*C13,2))`
- H8 `=IF(AND(H6="OK",H7="OK"),D11*D12,"")`; H9 `=IF(AND(H6="OK",H7="OK"),IF(H8<=2000,"OK","Check: too many mixes"),"Check")`
- E5 `=IF(AND(TRIM(A5)="",B5="",C5=""),"",IF(OR(NOT(ISNUMBER(B5)),NOT(ISNUMBER(C5))),"Check: units or price blank or text",IF(OR(B5<1,B5<>INT(B5)),"Check: units must be a whole number",IF(C5<=0,"Check: price must be above 0",IF(TRIM(A5)="","Check: no pack name",IF(COUNTIF($B$5:$B$7,B5)>1,"Check: same units as another pack","OK"))))))`

Tie rule (decided by the builder; state it in the post): when two mixes cost the same, the sheet picks the one with the fewest units bought (fewest extra), then the fewest packs, then the first row (fewer large packs).

Sample values (fictional; recalculated). Packs: 48-jar case $52.00, 12-jar case $15.00, Single jar $1.80 (price per unit $1.08, $1.25, $1.80). 95 needed: 27 mixes; H11:H20 $104.00, 2, 0, 0, 96, 1, $1.09, $116.80, $12.80, 18. Sorted rows: Large 48 $52.00 counts 3 biggest-first 1; Medium 12 $15.00 counts 9 biggest-first 3; Smallest 1 $1.80 biggest-first 11.
| Jars needed | Cheapest mix | Cost | Extra | Biggest-first | Biggest-first cost |
|---|---|---|---|---|---|
| 95 | 2 x 48 | $104.00 | 1 | 1 x 48 + 3 x 12 + 11 singles | $116.80 |
| 130 | 2 x 48 + 3 x 12 | $149.00 | 2 | 2 x 48 + 2 x 12 + 10 singles | $152.00 |
| 47 | 1 x 48 | $52.00 | 1 | 3 x 12 + 11 singles | $64.80 |
| 40 | 1 x 48 | $52.00 | 8 | 3 x 12 + 4 singles | $52.20 |
| 10 | 1 x 12 | $15.00 | 2 | 10 singles | $18.00 |
| 7 | 7 singles | $12.60 | 0 | same | $12.60 |
| 200 | 4 x 48 + 8 singles | $222.40 | 0 | same | $222.40 |
Mixes rows for 95 (first rows): mix 0 = 0/0/95 $171.00; mix 8 = 0/8/0 $120.00 (96 units); mix 12 = 1/3/11 $116.80; mix 13 = 1/4/0 $112.00; mix 18 = 2/0/0 $104.00 (chosen).

Edge cases tested (all matched the Python model)
- Ties: 24 needed with 24 @ $24.00, 12 @ $12.00, 1 @ $1.00 (every mix of exactly 24 costs $24.00): 1 x 24-case (fewest packs). 23 needed with 24 @ $23.00, 12 @ $12.00, 1 @ $1.00: 1 x 12 + 11 singles (23 units, no extra) instead of the 24-case. Plain MATCH on the lowest cost (mutation) picked 24 singles and 23 singles.
- Packs typed singles-first or medium-first (same answer); singles at $1.60 and 130 needed ($149.00, 48 mixes; biggest-first $150.00); a 6-pack as the smallest (95: $104.00 vs $113.80); two packs only (48 and singles, also with the middle row blank: 3 mixes, $104.00 vs $136.60); a 48-case priced $99.00 (never chosen: 95 singles at $1.00, $95.00); 48, 49, 96, 12, 1, 500 needed.
- 3,000 needed: 16,064 mixes, Check: too many mixes, result blank. 1,998 with 1998/2/1 packs: exactly 2,000, OK. 2,300 with 500/100/10: $1,855.00.
- Check: need 0, blank, "95" text, 9.5; one pack only; two packs with 48 units; price as text; price 0; units 1.5; no pack name. 60 random needs and price sets.

Pitfalls the post must state
- Three pack sizes at most (two counted, the smallest fills the rest); up to 2,000 mixes, then a Check, never a wrong answer.
- The supplier and prices are fictional; no claim about any supplier. Minimum orders and shipping are outside the sheet (link round-order-quantity-up-to-case-size-minimum-order).
- The Mixes counter shows counts 0, 1, 2 ... (not Yes/No).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The by-hand order costs more: biggest-first $116.80 vs $104.00 for 95.
- No result and H9 says Check: too many mixes: the need is large compared with the medium and large packs (3,000 jars: 16,064 mixes).
- No result and a pack row says Check: same units as another pack, units must be a whole number, or price as text.
- Two mixes cost the same: the tie rule picks fewest extra units, then fewest packs (plain MATCH would take the first row, 24 singles in the tie test).

Hero: `save(out, render(t+"tidy-tabs-cheapest-pack-mix.xlsx","Order",hide_cols="E",name="packs"), render(short_sheet_copy(t+"tidy-tabs-cheapest-pack-mix.xlsx","Mixes",24,"mixes"),"Mixes",name="mixes"))` (two cards: Order with the Status column hidden, Mixes rows 1 to 24 ending on the green chosen mix 18; legible, the Order card is the smaller one).

---
## 3. ab-test-two-listing-versions-conversion-rate-significance-normsdist-sqrt-excel-google-sheets
Title (hook-writer final): How to Check If an A/B Test Result Is Significant in Excel and Google Sheets (topics title: How to Tell If an A/B Test Result Is Real (Conversion Rate Significance) ...)
Template: templates/tidy-tabs-ab-test-significance.xlsx (tabs: Tests, How to use)

Layout
- Tests: A1 Confidence level / B1 95%, A2 Power for visits needed / B2 80%, A3 Fewest orders per version / B3 10 (inputs), C1:C3 notes. Header row 5, rows 6:55: A Test name, B Visits A, C Orders A, D Visits B, E Orders B (inputs); F Rate A, G Rate B, H Lift (B vs A), I Pooled rate, J Standard error, K z, L p-value (two-tailed), M Result, N Visits needed per version, O Status. Summary Q5:R12: Settings (B1:B3) (R5), Tests listed (R6), Likely a real difference (R7), Not enough evidence yet (R8), Caution: few orders (R9), Rows showing Check (R10), Typed below row 55 (not read) (R11), CHECK (R12). Landscape, row 5 repeated.
- Deviations from the topic: N shows "Rates are equal" when the rates are equal and "Over 1,000,000" above 1,000,000 (so the 950/19 vs 960/19 row shows Over 1,000,000, not 7,052,685); F:N stay blank on Check rows; Status adds whole-number, 0 visits, every-visit-ordered and no-name checks.

Key formulas (exact text in the file)
- F6 `=IF(OR($O6="",LEFT($O6,5)="Check"),"",C6/B6)`; G6 `=IF(F6="","",E6/D6)`; H6 `=IF(F6="","",IF(F6=0,"",G6/F6-1))`
- I6 `=IF(F6="","",(C6+E6)/(B6+D6))`; J6 `=IF(F6="","",SQRT(I6*(1-I6)*(1/B6+1/D6)))`; K6 `=IF(F6="","",(G6-F6)/J6)`
- L6 `=IF(F6="","",2*(1-NORMSDIST(ABS(K6))))`
- M6 `=IF(F6="","",IF(L6<=1-$B$1,"Likely a real difference","Not enough evidence yet"))`
- N6 `=IF(F6="","",IF(G6=F6,"Rates are equal",IF((NORMSINV(1-(1-$B$1)/2)*SQRT(2*((F6+G6)/2)*(1-(F6+G6)/2))+NORMSINV($B$2)*SQRT(F6*(1-F6)+G6*(1-G6)))^2/(G6-F6)^2>1000000,"Over 1,000,000",ROUNDUP((NORMSINV(1-(1-$B$1)/2)*SQRT(2*((F6+G6)/2)*(1-(F6+G6)/2))+NORMSINV($B$2)*SQRT(F6*(1-F6)+G6*(1-G6)))^2/(G6-F6)^2,0))))` (core `=ROUNDUP((NORMSINV(1-(1-conf)/2)*SQRT(2*p*(1-p))+NORMSINV(power)*SQRT(rA*(1-rA)+rB*(1-rB)))^2/(rB-rA)^2,0)` with p the average of the two rates)
- O6 `=IF(AND(TRIM(A6)="",B6="",C6="",D6="",E6=""),"",IF(OR(NOT(ISNUMBER(B6)),NOT(ISNUMBER(C6)),NOT(ISNUMBER(D6)),NOT(ISNUMBER(E6))),"Check: visits or orders blank or text",IF(OR(B6<0,C6<0,D6<0,E6<0,B6<>INT(B6),C6<>INT(C6),D6<>INT(D6),E6<>INT(E6)),"Check: counts must be whole numbers",IF(OR(B6=0,D6=0),"Check: 0 visits",IF(OR(C6>B6,E6>D6),"Check: more orders than visits",IF(C6+E6=0,"Check: no orders in either version",IF(C6+E6=B6+D6,"Check: every visit ordered",IF(TRIM(A6)="","Check: no test name",IF($R$5<>"OK","Check: settings (B1:B3)",IF(MIN(C6,E6)<$B$3,"Caution: few orders","OK"))))))))))`
- R5 `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B1>0,B1<1,B2>0,B2<1,B3>=0),"OK","Check"),"Check")`

Sample values (fictional; recalculated at 95% confidence, 80% power, caution below 10 orders)
| Test | A visits / orders | B visits / orders | Rate A | Rate B | Lift | z | p-value | Result | Visits needed | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Listing photo: flat lay vs lifestyle | 1,840 / 37 | 1,910 / 58 | 2.01% | 3.04% | 51.0% | 1.998 | 0.0457 | Likely a real difference | 3,669 | OK |
| Price: $24.00 vs $26.00 | 1,200 / 30 | 1,180 / 36 | 2.50% | 3.05% | 22.0% | 0.818 | 0.4132 | Not enough evidence yet | 13,959 | OK |
| Newsletter subject line | 400 / 8 | 420 / 15 | 2.00% | 3.57% | 78.6% | 1.362 | 0.1731 | Not enough evidence yet | 1,721 | Caution: few orders |
| Free shipping banner | 950 / 19 | 960 / 19 | 2.00% | 1.98% | -1.0% | -0.033 | 0.9740 | Not enough evidence yet | Over 1,000,000 | OK |
| Shop photos: old vs new (one year) | 52,000 / 500 | 50,000 / 570 | 0.96% | 1.14% | 18.6% | 2.797 | 0.0052 | Likely a real difference | 51,246 | OK |
Pooled rate and standard error for the first row: 2.53%, 0.0051. R5:R12: OK, 5, 2, 3, 1, 0, 0, OK.

Edge cases tested (all matched the Python model)
- Confidence 99%: listing photo Not enough evidence yet (5,460 needed), only the shop photos test likely real (76,253). 90%: 2,890 / 10,996 / 1,356 needed, same results as 95%. Power 90%: 4,912 / 18,687 / 2,303.
- 1 vs 0 orders of 500 (p 0.3171, 3,920, Caution: few orders; lift -100%), 0 vs 1 (lift blank), 100% conversion in B (50/10 vs 40/40: z 7.589, 5 visits), equal rates (z 0, p 1, Rates are equal), B worse than A (z -1.998, same p 0.0457), 100,000 / 2,000 vs 100,000 / 2,001 (Over 1,000,000), 900,000 / 18,000 vs 18,900 (z 4.734, 315,206), a blank row in the middle (skipped).
- Check: decimal visits, 0 orders in both, 520 orders on 500 visits, -3 orders, "1840" as text, blank orders, 0 visits, every visit ordered, no test name; settings "95%" as text, 100% or power 0 (R5 Check, every row Check: settings); a row on row 56. 12 random sets of 50 tests.

Pitfalls the post must state
- Pooled two-proportion z test, normal approximation, two-tailed; fine for visits in the hundreds and up; not for revenue per visitor.
- "Not enough evidence yet" is not proof the versions are the same. The 10-order caution is a setting, not a rule.
- One-tailed p (half the two-tailed one) and stopping the day p dips under 0.05 are cautions only, not measured.
- NORMSDIST and NORMSINV were checked against Python math.erf / NormalDist only. Etsy named only as the reader's own stats; no claim about Etsy tools.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- #DIV/0! in z and p: zero orders in both versions (copy without the Status guard; the template shows Check: no orders in either version).
- A 104% rate and p 0 ("Likely a real difference"): orders typed in the visits column, 520 on 500 (copy without the guard; visits needed showed an error); the template shows Check: more orders than visits.
- Visits needed says Over 1,000,000 or Rates are equal: the two rates are almost or exactly the same (950/19 vs 960/19).
- The result flips between 95% and 99%: the listing photo test (p 0.0457) passes at 95% and not at 99%.

Hero: `render(short_sheet_copy(t+"tidy-tabs-ab-test-significance.xlsx","Tests",10,"ab"),"Tests",hide_cols="IJPQR",name="ab")` (pooled rate, standard error and the summary hidden so the five tests read larger; with hide_cols "IJ" only the summary shows too but rows 11 and 12 of it are cut by last_row 10).

---
## 4. find-and-replace-multiple-words-from-a-list-whole-word-substitute-trim-excel-google-sheets
Title (hook-writer final, same as topics): How to Find and Replace Multiple Words at Once From a List in Excel and Google Sheets
Template: templates/tidy-tabs-replace-words-from-list.xlsx (tabs: Replacements, Text, How to use)

Layout
- Replacements: A1:A2 notes. Header row 4, rows 5:24: A Find, B Replace with (inputs, text format), C Status. Summary E4:F8: Pairs used (F4), Rows showing Check (F5), Rows with a caution (F6), Typed below row 24 (not used) (F7), CHECK (F8). Landscape.
- Text: A1:A2 notes. Header row 4, rows 5:304: A Original (input), B Cleaned, C Padded (spaces doubled), D:W After pair 1 to After pair 20 (C:W grouped and hidden), X Changed?, Y Pairs applied. Summary AA4:AB8: Rows with text (AB4), Rows changed (AB5), Replacements tab CHECK (AB6), Typed below row 304 (not read) (AB7), CHECK (AB8). Landscape.
- Deviations from the topic: the padding is its own helper column C, so the pair columns are D:W and B is `=TRIM(W5)` (topic: C:V and TRIM(V5)); Changed? is X and Pairs applied (pairs that changed the row) is Y; Replacements Status has no "find inside another find" check (harmless with whole words) and shows the order trap as "Caution: a later pair changes this replacement".

Key formulas (exact text in the file)
- Text C5 `=IF(A5="","",SUBSTITUTE(" "&A5&" "," ","  "))`
- Text D5 `=IF(INDEX(Replacements!$A$5:$A$24,COLUMNS($D5:D5))="",C5,SUBSTITUTE(C5," "&INDEX(Replacements!$A$5:$A$24,COLUMNS($D5:D5))&" "," "&INDEX(Replacements!$B$5:$B$24,COLUMNS($D5:D5))&" "))` (filled right to W5: COLUMNS($D5:W5) = 20, previous column V5) (core `=SUBSTITUTE(prev," "&find&" "," "&replace&" ")`)
- Text B5 `=IF(A5="","",TRIM(W5))`; X5 `=IF(A5="","",IF(EXACT(A5,B5),"","Yes"))`; Y5 `=IF(A5="","",SUMPRODUCT(--NOT(EXACT(C5:V5,D5:W5))))`
- Replacements C5 `=IF(AND(A5="",B5=""),"",IF(A5="","Check: find is blank",IF(LEN(A5)<>LEN(SUBSTITUTE(A5," ","")),"Check: find has a space (one word only)",IF(SUMPRODUCT(--EXACT($A$5:$A$24,A5))>1,"Check: same find twice",IF(SUMPRODUCT(($A$6:$A$24<>"")*ISNUMBER(FIND(" "&$A$6:$A$24&" "," "&B5&" ")))>0,"Caution: a later pair changes this replacement",IF(B5="","OK (deletes the word)","OK"))))))` (C24 has no caution part)
- Replacements F8 `=IF(AND(F4>0,F5=0,F7=0),"OK","Check")`; Text AB8 `=IF(AND(AB4>0,AB6="OK",AB7=0),"OK","Check")`

Sample values (fictional; recalculated). Pairs: Lg -> Large, Sm -> Small, Med -> Medium, w/ -> with, Lav -> Lavender, Vnl -> Vanilla, oz. -> oz, pk -> pack (all OK).
| Original | Cleaned | Changed? | Pairs applied |
|---|---|---|---|
| Soy Candle Lg Lav w/ Wood Wick | Soy Candle Large Lavender with Wood Wick | Yes | 3 |
| Smoky Sm Candle 8 oz. | Smoky Small Candle 8 oz | Yes | 2 |
| Vnl Bean Soy Candle Med | Vanilla Bean Soy Candle Medium | Yes | 2 |
| Beeswax Tea Lights 12 pk | Beeswax Tea Lights 12 pack | Yes | 1 |
| Lav Soap Bar w/ Oatmeal | Lavender Soap Bar with Oatmeal | Yes | 2 |
| Sm Lav Tin Candle | Small Lavender Tin Candle | Yes | 2 |
| Gift Set Lg w/ 3 Sm Candles | Gift Set Large with 3 Small Candles | Yes | 3 |
| Tea Light Lav (Lg) | Tea Light Lavender (Lg) | Yes | 1 |
| Smoky Vnl Wax Melts 3 oz. | Smoky Vanilla Wax Melts 3 oz | Yes | 2 |
| Cedar Wood Wick Candle | Cedar Wood Wick Candle | | 0 |
F4:F8: 8, 0, 0, 0, OK. AB4:AB8: 10, 9, OK, 0, OK.

Traps (same 8 pairs unless noted; LibreOffice results)
| Text | Template result | Plain nested SUBSTITUTE | Single-space padding (Lg only) |
|---|---|---|---|
| Smoky Sm Candle 8 oz. | Smoky Small Candle 8 oz | Smalloky Small Candle 8 oz | |
| Lg Lg Gift Set | Large Large Gift Set | Large Large Gift Set | Large Lg Gift Set |
| Lg, Sm and XL | Lg, Small and XL | Large, Small and XL | |
| Tea Light Lav (Lg) | Tea Light Lavender (Lg) | Tea Light Lavender (Large) | |
| sm lav tin | sm lav tin (unchanged) | sm lav tin | |
| "  Lg   Tin  " | Large Tin | | |
The plain nested column used the researcher's five pairs (Lg, Sm, w/, Lav, oz.).

Edge cases tested (all matched the token model)
- "Lg Lg Lg" -> "Large Large Large"; "Lg  Lg" -> "Large Large"; "Lg" alone -> "Large"; "LgLg", "Lg-Tin" and Lg<tab>Tin unchanged; 2026 and 12.5 typed as numbers -> text "2026" and "12.5" (Changed? blank).
- Blank pair rows in the middle: skipped. Blank replacement for Sm: "Smoky Sm Candle 8 oz." -> "Smoky Candle 8 oz." (Status OK (deletes the word)). A number as a find (8 -> eight, oz. -> ounce): "8 oz. jar" -> "eight ounce jar", "18 oz. jar" -> "18 ounce jar". lg as its own pair: "lg tin" -> "Large tin". 20 pairs; find = whole text.
- Order trap: Lg -> Large above Large -> XL: "Lg Soy Candle" -> "XL Soy Candle" (row 5 Caution); swapped order: "Large Soy Candle". Sm -> S above S -> Small: "Smoky S Sm" -> "Smoky Small Small".
- Check: find blank with a replacement, "Wood Wick" (space), Lg twice; a pair on row 25, text on row 305; no pairs (CHECK Check). 8 random sets of 1 to 20 pairs on 300 random titles with extra and doubled spaces.

Pitfalls the post must state
- 20 pairs, one helper column each; more pairs means more columns (How to use 10).
- Whole words are space-separated only: a word next to punctuation, a tab or a hyphen is not matched; SUBSTITUTE is case-sensitive. Fixes: extra pairs such as (Lg) or lowercase rows, or clean punctuation first. LibreOffice Calc results.
- Pairs run in list order; the Status caution flags a replacement that a later pair changes.
- No REDUCE, LAMBDA or macros by design so it works in older versions; no claim it was tested in Excel or Sheets. Known limit (not reproduced): a multi-word replacement containing the same later find twice.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- "Smoky" became "Smalloky": plain SUBSTITUTE replaces letters inside words.
- Only the first of two neighbors changed ("Large Lg Gift Set"): single-space padding; double the spaces.
- "Lg," or "(Lg)" or "sm" not replaced: punctuation or case.
- "Lg" ended up as "XL": a later pair changed the replacement (order trap, Caution in Status).

Hero: `render(short_sheet_copy(t+"tidy-tabs-replace-words-from-list.xlsx","Text",14,"repl"),"Text",name="repl")` (Original, Cleaned, Changed?, Pairs applied and the summary; very legible). Replacements A4:C12 as an optional second card.

---
## 5. find-transposed-digits-when-totals-dont-match-divisible-by-9-mid-mod-excel-google-sheets
Title (hook-writer final): How to Find Transposed Digits When Totals Don't Match in Excel and Google Sheets (topics title: How to Find a Transposed Number When Your Totals Are Off by an Amount Divisible by 9 ...)
Template: templates/tidy-tabs-find-transposed-digits.xlsx (tabs: Entries, How to use)

Layout
- Entries: A1 Deposit (should total) / B1 (input, $1,223.04), A2 Log total / B2, A3 Gap in cents / B3, A4 Divisible by 9? / B4; C1:C4 notes. Header row 5, rows 6:505: A Date, B Amount (inputs); C Cents, D Digits, E Length, F:J Swap 1-2 to Swap 5-6 (grouped and hidden), K Possible swap (should be), L Possible slide (should be), M Extra entry, N Wrong sign (should be), O Status. Summary Q5:R14: Entries with an amount (R5), Amounts typed as text (not in the total) (R6), Possible swaps (R7), Possible slides (R8), Extra entry or typed twice (R9), Wrong sign (R10), Result (R11), Rows showing Check (R12), Typed below row 505 (not read) (R13), CHECK (R14). Landscape, row 5 repeated.
- Deviations from the topic: swap cells also blank a swap that would put a 0 first; K/L/N show "Maybe $87.50" style corrected amounts; M is "Typed twice?" only when the same amount is in the log more than once, otherwise "Extra entry?"; summary has a Result line.

Key formulas (exact text in the file)
- B2 `=SUM(B6:B505)`; B3 `=IF(ISNUMBER(B1),ROUND((B1-B2)*100,0),"")`; B4 `=IF(B3="","",IF(B3=0,"Totals match",IF(MOD(ABS(B3),9)=0,"Yes","No")))`
- C6 `=IF($O6<>"OK","",ROUND(ABS(B6)*100,0))`; D6 `=IF(C6="","",TEXT(C6,"000000"))`; E6 `=IF(C6="","",LEN(C6))`
- F6 (pair 1-2) `=IF(C6="","",IF(OR(1<7-$E6,AND(1=7-$E6,MID($D6,2,1)="0")),"",(MID($D6,2,1)-MID($D6,1,1))*9*10^4*SIGN($B6)))`; H6 (pair 3-4) `=IF(C6="","",IF(OR(3<7-$E6,AND(3=7-$E6,MID($D6,4,1)="0")),"",(MID($D6,4,1)-MID($D6,3,1))*9*10^2*SIGN($B6)))`; J6 (pair 5-6) `...*9*10^0*SIGN($B6)))` (core `=(MID(digits,k+1,1)-MID(digits,k,1))*9*10^(5-k)`)
- K6 `=IF(OR(C6="",$B$3="",$B$3=0),"",IF(COUNTIF(F6:J6,$B$3)>0,"Maybe "&TEXT(B6+$B$3/100,"$#,##0.00"),""))`
- L6 `=IF(OR(C6="",$B$3="",$B$3=0),"",IF($B$3=9*ROUND(B6*100,0),"Maybe "&TEXT(B6*10,"$#,##0.00"),IF($B$3*10=-9*ROUND(B6*100,0),"Maybe "&TEXT(B6/10,"$#,##0.00"),"")))`
- M6 `=IF(OR(C6="",$B$3="",$B$3=0),"",IF($B$3=-ROUND(B6*100,0),IF(COUNTIF($B$6:$B$505,B6)>1,"Typed twice?","Extra entry?"),""))`
- N6 `=IF(OR(C6="",$B$3="",$B$3=0),"",IF($B$3=-2*ROUND(B6*100,0),"Maybe "&TEXT(-B6,"$#,##0.00"),""))`
- O6 `=IF(AND(A6="",B6=""),"",IF(B6="","Check: no amount",IF(NOT(ISNUMBER(B6)),"Check: amount is text (not in the total)",IF(ROUND(B6,2)<>B6,"Check: more than 2 decimals",IF(ABS(B6)>=10000,"Check: $10,000.00 or more (not checked)","OK")))))`
- R11 `=IF(B3="","Type the deposit in B1",IF(B3=0,"Totals match",IF(SUM(R7:R10)=0,"No single entry explains the gap",SUM(R7:R10)&" possible fixes to recheck")))`; R14 `=IF(AND(ISNUMBER(B1),R5>0,R6=0,R12=0,R13=0),"OK","Check")`

Sample values (fictional; recalculated). Dates 11/14/2026 (first 7) and 11/15/2026 (last 7). Log total $1,214.04, deposit $1,223.04, gap 900 cents, divisible Yes.
| Amount | Digits | Possible swap |
|---|---|---|
| $42.00 | 004200 | |
| $18.50 | 001850 | |
| $78.50 | 007850 | Maybe $87.50 |
| $126.00 | 012600 | |
| $33.75 | 003375 | |
| $9.00 | 000900 | |
| $64.25 | 006425 | |
| $250.00 | 025000 | |
| $15.00 | 001500 | |
| $87.40 | 008740 | |
| $45.60 | 004560 | Maybe $54.60 |
| $112.80 | 011280 | Maybe $121.80 |
| $29.99 | 002999 | |
| $301.25 | 030125 | Maybe $310.25 |
R5:R14: 14, 0, 4, 0, 0, 0, 4 possible fixes to recheck, 0, 0, OK. Hidden swap values for $78.50: H 900, I -270, J -45.

Other deposits with the same 14 entries (type in B1)
| Deposit | Gap (cents) | Divisible by 9? | Listed | Result |
|---|---|---|---|---|
| $1,223.04 | 900 | Yes | 4 swaps (above) | 4 possible fixes to recheck |
| $989.04 | -22,500 | Yes | $250.00 slide: Maybe $25.00 | 1 possible fixes to recheck |
| $1,146.54 | -6,750 | Yes | $33.75 wrong sign: Maybe -$33.75 | 1 possible fixes to recheck |
| $1,172.04 | -4,200 | No | $42.00 Extra entry? | 1 possible fixes to recheck |
| $1,178.04 | -3,600 | Yes | none | No single entry explains the gap |
| $1,214.04 | 0 | Totals match | none | Totals match |
Note the Result wording "1 possible fixes to recheck" (plural kept by the formula); quote it as it is or round it to "one entry to recheck" in prose. With $42.00 typed twice (log $1,256.04, deposit $1,214.04): gap -4,200, both $42.00 rows "Typed twice?".

Edge cases tested (all matched the Python model)
- Refunds -18.50 and -250.00 (SIGN flips the swap change), $9,999.99 (6 digits), 0.05 and 0 (no pairs), 1000.00, 55.55 (equal digits, change 0, never a match), 101.00.
- Check: $12,345.67 ($10,000.00 or more), 10.004 (more than 2 decimals), "45.00" text (not in the total, R6 1), a date with no amount; deposit blank or "1223.04" as text ("Type the deposit in B1", CHECK Check); a row on row 506.
- 48 random logs of 20 to 499 entries with one planted error each (12 adjacent swaps, 12 slides, 12 typed twice, 12 wrong signs): the planted entry was listed in its column in all 48.

Pitfalls the post must state
- Adjacent-digit swaps, decimal slides by one place, extra or doubled entries and flipped signs only; non-neighbor swaps, a missing entry or two errors at once are not found (one line each).
- Divisible by 9 is a clue, not proof; four entries match the sample's $9.00 gap, so check each against receipts. The sheet points at entries; no bookkeeping, audit or tax advice.
- Amounts up to $9,999.99; text amounts are not in the total (Status says so).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A suggestion like $402.00 for a $42.00 entry: the swap formula without the length guard swaps a padding zero of TEXT(cents,"000000") (probe p5, deposit $1,574.04 also gave $405.60 for $45.60); the template blanks those pairs.
- Divisible by 9 but nothing listed: $1,178.04 (gap -$36.00); look for two errors, a non-neighbor swap or a missing entry.
- Gap not divisible by 9: not a swap; check the Extra entry column ($1,172.04: $42.00).
- A refund typed as a sale: the gap is minus 2 x the amount (Wrong sign).

Hero: `render(short_sheet_copy(t+"tidy-tabs-find-transposed-digits.xlsx","Entries",19,"trans"),"Entries",hide_cols="CDEPQR",name="trans")` (B1:B4, the 14 entries and the K:O columns; very legible; C also hides the C1:C4 notes). With hide_cols "" the digits and summary show too, smaller.

---
Not built or tested (so not for the posts): calendar-day caps, deposits, damage or cancellation fees, lodging tax, more than 3 pack sizes, minimum orders or shipping, any supplier's prices, more than 2,000 mixes, one-tailed tests, stopping a test early, revenue per visitor, Bayesian tests, Etsy's own tools, punctuation-aware or case-insensitive matching, more than 20 pairs, REDUCE/LAMBDA versions, non-adjacent swaps, two errors at once, a missing entry, amounts of $10,000.00 or more, bookkeeping or audit practice, typing by hand. Nothing here was run in Excel or Google Sheets.
