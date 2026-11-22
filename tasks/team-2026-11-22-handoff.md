# Handoff 2026-11-22: five templates built and tested (LibreOffice Calc 24.2.7.2 only), for the designer and the script-writers

Builder stage for the five topics in tasks/team-topics-2026-11-22.md (researcher) and tasks/hooks-2026-11-22.md (hook-writer). Nothing is committed or pushed; no post file exists yet in posts/ready. Files: scripts/build-templates.py (`alphabetical_sort_formula`, `combined_shipping_calculator`, `random_coupon_codes` with `COUPON_MAX_CHARS`, `production_queue_finish_dates`, `bought_together_pairs` with `PAIRS_MAX_PRODUCTS`, registered after `url_handle_from_title()`), five new files in templates/, the record in tasks/briefs.md ("2026-11-22 build and test"). Heroes are not made yet (designer, see each section).

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 128 edited workbooks recalculated and compared (284,842 cell comparisons, 0 mismatches against independent Python models; a mutation run with one wrong formula per template produced mismatches in all five, so the harness catches errors), plus 2 probe workbooks of the traps read by hand; no error value in any recalculated template workbook. All five files exported to PDF: every page US Letter (612 x 792 or 792 x 612 points), no error values or ### in the PDF text. Two formula problems and five layout problems were found and fixed (tasks/briefs.md).

Repo note: `python3 scripts/build-templates.py` ran (exit 0) but rewrote 230 older templates byte-wise; those were restored with `git checkout -- templates/`, and the five new files were rebuilt by running only the five new functions. `git status`: scripts/build-templates.py, tasks/briefs.md, this handoff, five new .xlsx.

Post file names: `posts/ready/2026-11-22-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; red marks every status starting "Check" (and "Duplicate" on the Codes tab) and any CHECK cell that is not OK; orange marks "Late by ..." on the Queue tab; light green marks the 1 cells on the Baskets tab. Functions used in the five files: AND, COUNT, COUNTA, COUNTIF, COUNTIFS, FIND, IF, IFERROR, INDEX, INT, ISNUMBER, ISTEXT, LEFT, LEN, MATCH, MAX, MID, NETWORKDAYS, NOT, OR, RANDBETWEEN, ROUND, ROUNDUP, ROW, ROWS, SUBSTITUTE, SUM, SUMIF, SUMIFS, SUMPRODUCT, TEXT, TRIM, WEEKDAY, WORKDAY. No TODAY, FILTER, UNIQUE, SORT, LET, LAMBDA, XLOOKUP, TEXTSPLIT, MINIFS/MAXIFS, ISFORMULA; 0 `_xlfn` in the files. Help links were not checked by the builder.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-22-team-01-sort-list-alphabetically-with-formula-no-sort-function-countif-index-match-excel-google-sheets | tidy-tabs-alphabetical-sort-formula.xlsx | 26 workbooks, 41,705 comparisons |
| 2 | 2026-11-22-team-02-combined-shipping-first-item-each-additional-item-countifs-sumif-excel-google-sheets | tidy-tabs-combined-shipping-calculator.xlsx | 24 workbooks, 46,842 comparisons |
| 3 | 2026-11-22-team-03-generate-random-unique-coupon-codes-randbetween-mid-excel-google-sheets | tidy-tabs-random-coupon-codes.xlsx | 33 workbooks, 83,094 comparisons |
| 4 | 2026-11-22-team-04-production-queue-finish-dates-from-daily-capacity-workday-roundup-excel-google-sheets | tidy-tabs-production-queue-finish-dates.xlsx | 27 workbooks, 20,736 comparisons |
| 5 | 2026-11-22-team-05-products-frequently-bought-together-pair-counts-countifs-sumproduct-index-excel-google-sheets | tidy-tabs-bought-together-pairs.xlsx | 18 workbooks, 92,465 comparisons |

tested_in sentence (fill the counts from the table): "LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, N workbooks and N cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets." For the coupon post add: "Codes are random, so the test checked format, length, allowed characters and the duplicate counts, not fixed codes."

Hooks check (tasks/hooks-2026-11-22.md, chosen openings): every number matches the built files: 1 (19 names, 1st Street Coffee, apple Annie's Pies, Apple Annie's Pies, Ash Grove Apiary, 1st Street Coffee $45.00), 2 (2 candles + 1 mug: mug pays first, $8.00 + 2 x $1.50 = $11.00, every-line-first $14.50, shown in column K), 4 (40 a day from 11/23/2026, the order that brings the total to 160 finishes 11/30/2026, the 300-unit order promised 12/16/2026 finishes 12/18/2026, Status "Late by 2 days"), 5 (Lavender candle + Wick trimmer 3 orders, most often with Wick trimmer, 3 of 5, 60.0%). Opening 3: the format is right (10 characters, MKT- plus 6 from the 32-character set) but MKT-3GBXMZ is a probe code; the template shows new codes on every recalculation, so write "codes such as MKT-3GBXMZ" and never "the download shows MKT-3GBXMZ". Formulas in the openings are the core forms; the cells hold guarded versions (below).

Hero renders (scripts/make-heroes.py `render(src, sheet, hide_cols=...)` fits the whole sheet on one page): the gray formula fill runs to the last row (Vendors and A to Z 204, Orders 304, Codes 504, Queue 104, Items 504, Baskets 304), so a whole-sheet render of those tabs will be tiny. Render the Pairs tab as it is, and for the others use a copy with the empty rows deleted (or a print area) before rendering. The Codes tab changes on every recalculation, so a hero of it shows whatever codes that render produced; caption it as an example.

---
## 1. sort-list-alphabetically-with-formula-no-sort-function-countif-index-match-excel-google-sheets
Title (from topics): How to Sort a List Alphabetically With a Formula (No SORT Function) in Excel and Google Sheets
Template: templates/tidy-tabs-alphabetical-sort-formula.xlsx (tabs: Vendors, A to Z, How to use)

Layout
- Vendors: A1:A2 notes. Header row 4, rows 5:204: A Vendor name, B Booth fee (or any second value) (inputs); C Rank (A to Z), D Rank (Z to A), E Status, F Rank, SUMPRODUCT form (fix for the Check rows). Summary H4:I9: Names entered (I4), Rows showing Check (I5), Ranks that repeat (must be 0) (I6), Rows where F differs from C (I7), Typed below row 204 (I8), CHECK (I9). Landscape.
- A to Z: A1:A2 notes. Header row 4, rows 5:204: A #, B Vendor (A to Z), C Booth fee, D gap, E Vendor (Z to A), F Booth fee. Summary H4:I8: Names entered (Vendors) (I4), Names shown A to Z (I5), Names shown Z to A (I6), Repeated ranks (must be 0) (I7), CHECK (I8). Print area A1:I204. Landscape.
- Deviations from the topic: F is a visible column (not hidden); Status adds "Check: number stored as text", "Check: extra spaces" (also a doubled space inside), "Check: * ? ~ or < > = in name" and "Check: value with no name"; the A to Z list uses the COUNTIF rank (C), as in the topic; a blank fee stays blank on A to Z.

Key formulas (exact text in the file)
- Vendors C5 `=IF(A5="","",COUNTIF($A$5:$A$204,"<"&A5)+COUNTIF($A$5:A5,A5))` (core `=COUNTIF(names,"<"&name)+COUNTIF(names_so_far,name)`)
- Vendors D5 `=IF(A5="","",COUNTIF($A$5:$A$204,">"&A5)+COUNTIF($A$5:A5,A5))`
- Vendors F5 `=IF(A5="","",SUMPRODUCT(($A$5:$A$204<>"")*($A$5:$A$204<A5))+SUMPRODUCT(($A$5:A5<>"")*($A$5:A5=A5)))`
- Vendors E5 `=IF(A5="",IF(B5="","","Check: value with no name"),IF(ISNUMBER(A5),"Check: number, not text",IF(ISNUMBER(--A5),"Check: number stored as text",IF(LEN(A5)<>LEN(TRIM(A5)),"Check: extra spaces",IF(OR(ISNUMBER(FIND("*",A5)),ISNUMBER(FIND("?",A5)),ISNUMBER(FIND("~",A5)),ISNUMBER(FIND(LEFT(A5,1),"<>="))),"Check: * ? ~ or < > = in name",IF(OR(COUNTIF($C$5:$C$204,C5)>1,COUNTIF($D$5:$D$204,D5)>1),"Check: rank repeats","OK"))))))`
- Vendors I6 `=SUMPRODUCT((C5:C204<>"")*(COUNTIF(C5:C204,C5:C204)>1))`
- A to Z B5 `=IFERROR(INDEX(Vendors!$A$5:$A$204,MATCH(ROWS(A$5:A5),Vendors!$C$5:$C$204,0)),"")` (core `=INDEX(names,MATCH(ROWS(counter),rank,0))`)
- A to Z C5 `=IFERROR(IF(INDEX(Vendors!$B$5:$B$204,MATCH(ROWS(A$5:A5),Vendors!$C$5:$C$204,0))="","",INDEX(Vendors!$B$5:$B$204,MATCH(ROWS(A$5:A5),Vendors!$C$5:$C$204,0))),"")`
- A to Z E5 and F5: the same on Vendors!$D$5:$D$204; A5 `=IF(AND(B5="",E5=""),"",ROWS(A$5:A5))`

Sample values (fictional; recalculated)
Vendors in sign-up order (A5:A24, B fee): Maple Hollow Farm $35.00, apple Annie's Pies $40.00, Blue Door Bakery $30.00, Zeke's Honey $25.00, blue door bakery $30.00, Bea's Bees $25.00, 1st Street Coffee $45.00, The Candle Barn $35.00, Oak & Ivy Soap $30.00, Maple Hollow Farm $35.00, (blank row 15), Dee-Lish Jams $25.00, De Luca Pasta $40.00, Apple Annie's Pies $40.00, Zeke's Honey $25.00, zinnia Flower Co $30.00, Cedar Ridge Eggs $20.00, Bea's Bees $25.00, Kettle Corn King $35.00, Ash Grove Apiary $25.00. Ranks C5:C24: 13, 2, 7, 17, 8, 5, 1, 16, 15, 14, blank, 11, 10, 3, 18, 19, 9, 6, 12, 4.
| # | Vendor (A to Z) | Booth fee | Vendor (Z to A) | Booth fee |
|---|---|---|---|---|
| 1 | 1st Street Coffee | $45.00 | zinnia Flower Co | $30.00 |
| 2 | apple Annie's Pies | $40.00 | Zeke's Honey | $25.00 |
| 3 | Apple Annie's Pies | $40.00 | Zeke's Honey | $25.00 |
| 4 | Ash Grove Apiary | $25.00 | The Candle Barn | $35.00 |
| 5 | Bea's Bees | $25.00 | Oak & Ivy Soap | $30.00 |
| 6 | Bea's Bees | $25.00 | Maple Hollow Farm | $35.00 |
| 7 | Blue Door Bakery | $30.00 | Maple Hollow Farm | $35.00 |
| 8 | blue door bakery | $30.00 | Kettle Corn King | $35.00 |
| 9 | Cedar Ridge Eggs | $20.00 | Dee-Lish Jams | $25.00 |
| 10 | De Luca Pasta | $40.00 | De Luca Pasta | $40.00 |
Rows 11 to 19 A to Z: Dee-Lish Jams $25.00, Kettle Corn King $35.00, Maple Hollow Farm $35.00, Maple Hollow Farm $35.00, Oak & Ivy Soap $30.00, The Candle Barn $35.00, Zeke's Honey $25.00, Zeke's Honey $25.00, zinnia Flower Co $30.00. Vendors I4:I9: 19, 0, 0, 0, 0, OK. A to Z I4:I8: 19, 19, 19, 0, OK.

Edge cases tested (all matched the Python model or the stated LibreOffice result)
- Trap list (12 names: Bob's Jams, 42, "*Star* Kettle Corn", Star Farm, "  Sunny Acres", Sunny Acres, apple, Apple, Ölberg Olive Oil, Zed, "~Tilde Tea", Oak): COUNTIF ranks 6, 1, 2, 9, 1, 10, 4, 5, 8, 11, 2, 7; the A to Z tab listed 10 names and lost "  Sunny Acres" and "~Tilde Tea"; I4:I9 12, 6, 4, 11, 0, Check; A to Z CHECK Check. Column F gave 1 to 12 with no gaps: 42, "  Sunny Acres", "*Star* Kettle Corn", "~Tilde Tea", apple, Apple, Bob's Jams, Oak, Ölberg Olive Oil, Star Farm, Sunny Acres, Zed.
- A name starting with > (">Zed") got rank 200 (COUNTIF reads "<>Zed" as "not equal"); it shows Check: * ? ~ or < > = in name.
- Case ties apple, Apple, APPLE, aPPle: list order kept. 200 copies of one name: ranks 1 to 200. Names of 124, 150 and 250 characters sorted with no error.
- Booth 1, Booth 10, Booth 2, Booth 20, Booth 3 (text order). 10 Acre Farm, 100 Mile Jams, 1st Street Coffee, 2nd Street Pies, Ash Grove.
- Accents in LibreOffice: Eclair Bakery, Éclair Bakery, Ezra, Oak, Olberg Farm, Ölberg Olive Oil, Ostrich Eggs.
- 13 random lists of up to 200 names (20% lowercase, repeats, blank rows) matched a case-insensitive Python sort.

Pitfalls the post must state
- The order of spaces, symbols, accents and upper/lower case ties is LibreOffice Calc's; say so and link the COUNTIF help pages instead of claiming Excel or Sheets order.
- Not a "The"-skipping sort and not a natural number sort (Booth 10 before Booth 2).
- SORT in newer Excel and Sheets gets one line, no steps.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A name is missing and a blank row appears in the A to Z list: two names got the same rank (a leading space, a number, or * ? ~ in a name). Vendors I6 shows the count; column F is the fix.
- A name starting with > or < sorts to the end or the start: COUNTIF reads it as a comparison.
- Booth 10 comes before Booth 2: the names are text.
- apple and Apple: both kept, in the order typed.

Hero: sheet "A to Z", rows 1 to 24, hide_cols "GHI" (or a copy with rows 25:204 deleted); Vendors A4:E24 as an inset.

---
## 2. combined-shipping-first-item-each-additional-item-countifs-sumif-excel-google-sheets
Title (from topics): How to Charge Shipping for the First Item and Each Additional Item in Excel and Google Sheets
Template: templates/tidy-tabs-combined-shipping-calculator.xlsx (tabs: Profiles, Orders, How to use)

Layout
- Profiles: A1:A2 notes. Header row 4, rows 5:24: A Profile name, B First item $, C Each additional $ (inputs); D Status. Landscape.
- Orders: A1 note, A2 the rule. Header row 4, rows 5:304: A Order #, B Item, C Profile (dropdown from Profiles!$A$5:$A$24, warning style), D Qty (inputs); E First item $, F Each additional $, G Pays first-item price (1 or 0), H Line shipping, I Order shipping (first line of the order), J Status, K Every line pays first item (wrong, comparison). Summary M4:N12: Order lines listed (N4), Orders (N5), Lines showing Check (N6), Profiles rows showing Check (N7), Sum of order shipping (I) (N8), Sum of line shipping (H) (N9), Difference (must be $0.00) (N10), Typed below row 304 (N11), CHECK (N12). Landscape.
- Rule (Orders A2 and How to use line 3): "The item with the highest First item price pays it once; every other unit pays its Each additional price. On a tie, the first line listed pays it."
- Deviations from the topic: E/F end with +0 (a blank price becomes $0.00 and Profiles shows Check: price blank or text); K added; Profiles Status added; "Check: order # is text" fires only for a number stored as text, so text order numbers such as SO-1001 are fine.

Key formulas (exact text in the file)
- E5 `=IFERROR(INDEX(Profiles!$B$5:$B$24,MATCH(C5,Profiles!$A$5:$A$24,0))+0,"")`; F5 the same on column C
- G5 `=IF(AND(A5="",TRIM(B5)="",TRIM(C5)="",D5=""),"",IF(OR(A5="",E5="",NOT(ISNUMBER(D5))),0,IF(AND(D5>0,COUNTIFS($A$5:$A$304,A5,$D$5:$D$304,">0",$E$5:$E$304,">"&E5)=0,COUNTIFS($A$5:A5,A5,$D$5:D5,">0",$E$5:E5,E5)=1),1,0)))` (core: the topic's two-COUNTIFS flag)
- H5 `=IF(OR(A5="",E5="",F5="",D5=""),"",IF(G5=1,E5+(D5-1)*F5,D5*F5))` (core `=IF(flag=1,first+(qty-1)*additional,qty*additional)`)
- I5 `=IF(A5="","",IF(COUNTIF($A$5:A5,A5)=1,SUMIF($A$5:$A$304,A5,$H$5:$H$304),""))`
- K5 `=IF(I5="","",SUMIF($A$5:$A$304,A5,$H$5:$H$304)+SUMIFS($E$5:$E$304,$A$5:$A$304,A5,$D$5:$D$304,">0",$G$5:$G$304,0)-SUMIFS($F$5:$F$304,$A$5:$A$304,A5,$D$5:$D$304,">0",$G$5:$G$304,0))`
- J5 `=IF(AND(A5="",TRIM(B5)="",TRIM(C5)="",D5=""),"",IF(A5="","Check: no order #",IF(AND(ISTEXT(A5),ISNUMBER(--A5)),"Check: order # is text",IF(TRIM(C5)="","Check: no profile",IF(E5="","Check: profile not found",IF(D5="","Check: no qty",IF(NOT(ISNUMBER(D5)),"Check: qty is text",IF(OR(D5<0,D5<>INT(D5)),"Check: qty not a whole number","OK"))))))))`
- N10 `=ROUND(N8-N9,2)`; N12 `=IF(AND(N4>0,N6=0,N7=0,N10=0,N11=0),"OK","Check")`

Sample values (fictional; recalculated)
Profiles: Small $5.00 / $1.50, Fragile $8.00 / $3.00, Flat $4.00 / $0.75, Heavy $12.00 / $6.00.
| Order # | Lines | Line flagged (G=1) | Order shipping | Every line pays first (K) |
|---|---|---|---|---|
| 1001 | Soy candle Small x2, Mug Fragile x1 | Mug | $11.00 | $14.50 |
| 1002 | Art print Flat x3 (row 7), Sticker pack Flat x1 (row 16) | Art print | $6.25 | $9.50 |
| 1003 | Soy candle Small x1, Cast iron trivet Heavy x1, Art print Flat x2 | Trivet | $15.00 | $21.75 |
| 1004 | Mug Fragile x2, Vase Fragile x1 | Mug | $14.00 | $19.00 |
| 1005 | Soy candle Small x0, Art print Flat x1 | Art print | $4.00 | $4.00 |
| 1006 | Soy candle Small x1 | Candle | $5.00 | $5.00 |
| 1007 | Vase Fragile x1, Mug Fragile x1 (tie) | Vase (listed first) | $11.00 | $16.00 |
Line shipping H5:H18: $3.00, $8.00, $5.50, $1.50, $12.00, $1.50, $11.00, $3.00, $0.00, $4.00, $5.00, $0.75, $8.00, $3.00. N4:N12: 14, 7, 0, 0, $66.25, $66.25, $0.00, 0, OK.

Edge cases tested (all matched the Python model or the stated LibreOffice result)
- Three Flat lines qty 2, 1, 3 in one order: $7.75 (the first line pays). Fragile 1 + Heavy qty 0 + Fragile 2: $14.00 (qty 0 never pays first).
- A $0.00 free sample profile: 2 samples $0.00; sample + mug $8.00. All lines qty 0: $0.00.
- Profiles A $6.00/$5.00 and B $5.50/$1.00 in one order: $7.00 by this rule (K $11.50). A "largest gap" rule would charge $10.50; it is a different rule and the post must not offer it as the same thing.
- Lines of one order not next to each other: same totals. Profile typed "small" or "FRAGILE": found (LibreOffice MATCH ignores case).
- "Fragile " with a trailing space on an order line: Check: profile not found, and the order total leaves that line out (order shows $5.00 for its Small line).
- Text order number "2" next to the number 2: H $5.00 and $8.00, I shows $5.00 on the number line and nothing on the text line, N10 -$8.00, J Check: order # is text, CHECK Check. Qty typed as text "2": G 0, H $3.00, order still $11.00 by luck, Check: qty is text.
- Qty 1.5 and -1 (Check: qty not a whole number), blank qty, blank order #, blank profile, unknown profile; Profiles name twice, trailing space, blank price, negative price, no name (N7 6, CHECK Check). 20 profiles with 300 lines; 10 random price lists with 20 to 90 orders.

Pitfalls the post must state
- No platform's shipping rule is described or claimed; "check your shop's shipping settings help". No carrier rates; no tax on shipping.
- Text vs number order numbers: SUMIF/COUNTIF treat them differently from what you expect; the Status flags it. LibreOffice result only.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Order total too high: every line was charged its own first-item price (order 1001 $14.50 instead of $11.00; column K).
- Order total too low or Difference not $0.00: an order number typed as text (N10 -$8.00 in the test).
- A line shows Check: profile not found: a trailing space or typo in the profile name.
- Two lines tie on first-item price: the first one listed pays it (1007).

Hero: sheet "Orders", rows 1 to 18, hide_cols "LMN" (or a copy with rows 19:304 deleted), so G, H, I, J and K show; Profiles A4:C8 as an inset.

---
## 3. generate-random-unique-coupon-codes-randbetween-mid-excel-google-sheets
Title (from topics): How to Generate Random Coupon Codes in Excel and Google Sheets
Template: templates/tidy-tabs-random-coupon-codes.xlsx (tabs: Settings, Codes, How to use)

Layout
- Settings: A1:A4 labels, B1 Character set, B2 Prefix, B3 Random characters (1 to 8), B4 How many codes (1 to 500) (inputs), C1:C4 notes. Checks A6:B16: Characters in the set (B7), Characters that appear more than once (B8), Spaces, * ? or ~ in the set (B9), Character set (B10), Prefix (B11), Random characters (B12), How many codes (B13), Code length (prefix + random) (B14), Possible different codes (B15), Settings CHECK (B16). Landscape, no repeated title row.
- Codes: A1:A2 notes. Header row 4, rows 5:504: A #, B Code (live, changes on every recalculation), C Times in live list, D Pasted codes (paste values here) (input, text format), E Times in pasted list, F Status. Summary H4:I11: Codes asked for (I4), Live codes shown (I5), Live codes that repeat (I6), Pasted codes (I7), Pasted codes that repeat (I8), Pasted codes with the wrong length (I9), Settings CHECK (I10), Pasted list CHECK (after you paste) (I11). Landscape.
- Deviations from the topic: codes are made only while B16 is OK (a length of 9 or 10 shows Check and no codes, so the probe's 12-character result never appears); Settings rows are 7 to 16, not right under the inputs; I11 shows Check until something is pasted in D.

Key formulas (exact text in the file)
- Codes B5 (8 pieces; the post shows 6 and explains the IF pieces): `=IF(OR(Settings!$B$16<>"OK",ROWS(B$5:B5)>Settings!$B$4),"",Settings!$B$2&IF(Settings!$B$3>=1,MID(Settings!$B$1,RANDBETWEEN(1,LEN(Settings!$B$1)),1),"")&IF(Settings!$B$3>=2,MID(Settings!$B$1,RANDBETWEEN(1,LEN(Settings!$B$1)),1),"")&...&IF(Settings!$B$3>=8,MID(Settings!$B$1,RANDBETWEEN(1,LEN(Settings!$B$1)),1),""))`
- Core: `=MID(chars,RANDBETWEEN(1,LEN(chars)),1)`; short form for the first example `="MKT-"&MID($B$1,RANDBETWEEN(1,32),1)&...` (6 times), as in the topic.
- C5 `=IF(B5="","",COUNTIF($B$5:$B$504,B5))`; E5 `=IF(D5="","",COUNTIF($D$5:$D$504,D5))`
- F5 `=IF(D5<>"",IF(E5>1,"Duplicate",IF(LEN(D5)<>Settings!$B$14,"Check: length","OK")),IF(B5="","",IF(C5>1,"Duplicate (live)","Live (paste as values)")))`
- Settings B8 `=IF(LEN(B1)=0,0,SUMPRODUCT(--(IF(LEN($B$1)=0,0,LEN($B$1)-LEN(SUBSTITUTE($B$1,MID($B$1,ROW($A$1:INDEX($A:$A,MAX(1,LEN($B$1)))),1),"")))>1)))`
- Settings B12 `=IF(AND(ISNUMBER(B3),B3>=1,B3<=8,B3=INT(B3)),"OK","Check: a whole number from 1 to 8")`; B15 `=IF(AND(B10="OK",B12="OK"),B7^B3,"")`; B16 `=IF(AND(B10="OK",LEFT(B11,2)="OK",B12="OK",B13="OK"),"OK","Check")`
- Codes I6 `=COUNTIF(C5:C504,">1")` (COUNTIF, not SUMPRODUCT, so the "" cells are skipped)

Sample values (fictional settings; codes are random)
Settings: ABCDEFGHJKLMNPQRSTUVWXYZ23456789, MKT-, 6, 200. B7:B16: 32, 0, 0, OK, OK, OK, OK, 10, 1,073,741,824, OK. Three recalculations of the template: 200 codes each, all 10 characters, all 32 characters used, 200 distinct, every C = 1, different codes each time (first codes in the final run MKT-8MRFPP, MKT-BEK3Y7, MKT-WWWQL8). I4:I11 200, 200, 0, 0, 0, 0, OK, Check. The post shows codes "such as" these; never a fixed list as what the download shows.

Edge cases tested (all matched the Python model)
- 500 codes of 6 (500 distinct); SAVE + 4 characters (8 characters, 1,048,576 possible); 8 characters (12-character codes); no prefix (6 characters); prefix "FALL " (OK (has a space)).
- Forced repeats: prefix X + 1 character, 300 codes (32 distinct, every row Duplicate (live), I6 300); a 10-character set with 2 characters (100 possible, I6 499 of 500). Random sets of 27 and 29 characters with 3 characters: 4 rows flagged each, matching Python.
- Settings Check: a repeated letter (B8 2), a space or * in the set (B9 1), 8 characters, empty set, random characters 9, 10, 0, "6" as text, 6.5, how many 501 and 0, prefix "*MKT" and "<MKT": no codes, B16 Check.
- Pasted list: 200 typed codes (I11 OK); one code typed twice (both Duplicate, I8 2); MKT-ABC and MKT-ABCDEFG (Check: length, I9 2); MKT-ABCDEF and mkt-abcdef (Duplicate: LibreOffice COUNTIF ignores case).
- Probe: `=CHAR(RANDBETWEEN(65,90))` gave O or I in 36 of 500 cells (the topic's probe: 158 to 166 of 2,000).

Pitfalls the post must state
- Codes are random, not secret or secure: no passwords, gift cards or payments.
- RANDBETWEEN picks new codes on every recalculation; paste as values (Paste Special) before printing or importing. The paste step itself was not run by the builder (only typed values in D were tested); link the Paste Special help pages. Do not claim when each app recalculates.
- No claim about any shop's coupon import format or limits.
- Duplicates are rare but possible (1,073,741,824 possible codes for 6 from 32); the C column is the guard.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The codes changed after I printed them: the live column recalculated (three recalculations, three different lists).
- No codes appear: Settings CHECK is Check (length 9 or 10, a repeated or space character, a short set).
- A code shows Duplicate: two equal codes (forced with a 1-character length); type a new one over it in D.
- Pasted code shows Check: length: it is not prefix + random characters long.

Hero: sheet "Settings" as it is (inputs and checks), or "Codes" rows 1 to 20 with hide_cols "GHI" from a copy; caption the codes as one recalculation's example.

---
## 4. production-queue-finish-dates-from-daily-capacity-workday-roundup-excel-google-sheets
Title (from topics): How to Estimate When Each Order Will Be Finished From Your Daily Capacity in Excel and Google Sheets
Template: templates/tidy-tabs-production-queue-finish-dates.xlsx (tabs: Queue, Days off, How to use)

Layout
- Queue: A1 Start date / B1 (input), A2 Units per work day / B2 (input), C1:C2 notes. Header row 4, rows 5:104: A Order, B Promised date, C Units (inputs); D Units through this order, E First work day #, F Last work day #, G Start date, H Finish date, I Status (late = calendar days), J Work days late. Summary L4:M13: Start date and units per day (M4), Orders listed (M5), Total units (M6), Work days needed (M7), Last finish date (M8), Orders late (M9), Rows showing Check (M10), Days off typed as text (M11), Typed below row 104 (M12), CHECK (M13). Landscape.
- Days off: header row 1: A Day off (input, A2:A30), B Day, C Status; D1 note; D4/E4 Days off typed as text. Landscape, one page.
- Deviations from the topic: late is counted in calendar days in Status ("Late by 1 day", "Late by 2 days", singular fixed) and in work days in J with `NETWORKDAYS(INT(B5)+1,H5,days off)` (the topic's NETWORKDAYS(B5,H5)-1 undercounts when the promised date is a weekend); D uses SUMIF(">0") so text, 0 and negative units are not added; D:H are blank on Check rows except a row with units and no order name (it is still scheduled and shows Check: no order name); "No promised date" is not a Check.

Key formulas (exact text in the file)
- D5 `=IF(OR($M$4<>"OK",NOT(ISNUMBER(C5)),C5<=0),"",SUMIF($C$5:C5,">0"))` (core `=SUM($C$5:C5)`)
- E5 `=IF(D5="","",INT(ROUND((D5-C5)/$B$2,9))+1)`
- F5 `=IF(D5="","",ROUNDUP(ROUND(D5/$B$2,9),0))` (core `=ROUNDUP(units_so_far/units_per_day,0)`)
- G5 `=IF(E5="","",WORKDAY($B$1-1,E5,'Days off'!$A$2:$A$30))`; H5 `=IF(F5="","",WORKDAY($B$1-1,F5,'Days off'!$A$2:$A$30))`
- I5 `=IF(AND(TRIM(A5)="",B5="",C5=""),"",IF(TRIM(A5)="","Check: no order name",IF(C5="","Check: no units",IF(NOT(ISNUMBER(C5)),"Check: units are text",IF(C5<=0,"Check: units 0 or less",IF($M$4<>"OK","Check: start date or units per day",IF(B5="","No promised date",IF(NOT(ISNUMBER(B5)),"Check: promised date is text",IF(H5>B5,"Late by "&(H5-INT(B5))&IF(H5-INT(B5)=1," day"," days"),"On time")))))))))`
- J5 `=IF(OR(H5="",NOT(ISNUMBER(B5))),"",IF(H5>B5,NETWORKDAYS(INT(B5)+1,H5,'Days off'!$A$2:$A$30),0))`
- M4 `=IF(AND(ISNUMBER(B1),B1>=1,ISNUMBER(B2),B2>0),"OK","Check")`; M9 `=COUNTIF(I5:I104,"Late*")`
- Days off C2 `=IF(A2="","",IF(NOT(ISNUMBER(A2)),"Check: date is text",IF(WEEKDAY(A2,2)>5,"Weekend (no effect)","OK")))`

Sample values (fictional; recalculated). Start Monday 11/23/2026, 40 units a day, days off 11/26/2026, 11/27/2026, 12/24/2026, 12/25/2026.
| Order | Promised | Units | Through | Days | Start | Finish | Status | Work days late |
|---|---|---|---|---|---|---|---|---|
| W-101 | 11/25/2026 | 60 | 60 | 1-2 | 11/23/2026 | 11/24/2026 | On time | 0 |
| W-102 | 11/30/2026 | 100 | 160 | 2-4 | 11/24/2026 | 11/30/2026 | On time | 0 |
| W-103 | 12/02/2026 | 20 | 180 | 5-5 | 12/01/2026 | 12/01/2026 | On time | 0 |
| W-104 | 12/03/2026 | 140 | 320 | 5-8 | 12/01/2026 | 12/04/2026 | Late by 1 day | 1 |
| W-105 | 12/11/2026 | 80 | 400 | 9-10 | 12/07/2026 | 12/08/2026 | On time | 0 |
| W-106 | 12/16/2026 | 300 | 700 | 11-18 | 12/09/2026 | 12/18/2026 | Late by 2 days | 2 |
| W-107 | 12/23/2026 | 40 | 740 | 18-19 | 12/18/2026 | 12/21/2026 | On time | 0 |
M4:M13: OK, 7, 740, 19, 12/21/2026, 2, 0, 0, 0, OK. Days off B/C: Thu OK, Fri OK, Thu OK, Fri OK; E4 0.

Edge cases tested (all matched the Python model or the stated LibreOffice result)
- No days off: W-102 finishes 11/26/2026 (Thanksgiving) and W-103 11/27/2026; last finish 12/17/2026.
- Start Saturday 11/21/2026: day 1 is Monday 11/23/2026 (same dates). Start on the day off 11/26/2026: day 1 is Monday 11/30/2026.
- Exact capacity: 80 units at 40 finish on day 2 and the next order starts on day 3; 1 unit and 39 units share day 4; finishing on the promised date is On time, one day after is Late by 1 day.
- Capacity 0.1 with 0.3 units: 3 days (ROUND guard). Units 0, -20, "40" as text, blank: Check and no dates. Promised date as text: Check: promised date is text. Units per day 0 or text, start date text or blank: every row Check: start date or units per day.
- A Saturday on the Days off tab: Weekend (no effect). Blanks in the days-off list: no effect. 100 orders over 2027 and 10 random queues matched a Python work-day counter.
- Text day off "11/26/2026": LibreOffice still skipped it (same dates as the sample), but Days off shows Check: date is text and CHECK is Check. Do not say what Excel or Sheets do.
- Mutation WORKDAY($B$1,n) without -1: every date one work day late (W-101 11/24/2026 to 11/25/2026); this is the post's off-by-one trap.

Pitfalls the post must state
- Steady capacity in the listed order only; no staffing, overtime or wage talk. Days off are the shop's own closed days; no holiday list.
- WORKDAY basics: link business-days-ship-by-date and plan-backward-from-event-date; bars: link simple-gantt-chart.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Every date is one work day late: WORKDAY(start, n) instead of WORKDAY(start-1, n).
- An order finishes on a closed day: the Days off range is missing or the date is not on the Days off tab (W-102 on 11/26/2026).
- Dates are blank: units 0, negative or text, or B1/B2 not a date or not above 0 (M4 Check).
- A day off stored as text: Days off Status Check: date is text.

Hero: sheet "Queue", rows 1 to 11, hide_cols "KLM" (or a copy with rows 12:104 deleted).

---
## 5. products-frequently-bought-together-pair-counts-countifs-sumproduct-index-excel-google-sheets
Title (from topics): How to Find Which Products Are Bought Together Most Often in Excel and Google Sheets
Template: templates/tidy-tabs-bought-together-pairs.xlsx (tabs: Items, Baskets, Pairs, How to use)

Layout
- Items: A1:A2 notes. Header row 4, rows 5:504: A Order #, B Product (inputs); C Order no., D Status. Summary F4:G8: Item lines listed (G4), Orders (G5), Rows showing Check (G6), Lines not on the Pairs list (G7), Typed below row 504 (G8). Landscape.
- Baskets: A1:A2 notes. Row 4: A Order #, B Products in order, C4:N4 product names by formula from Pairs A5:A16. Rows 5:304: A order, B count, C:N 1 or 0 (1 is light green). Landscape.
- Pairs: A1:A2 notes. A5:A16 product names (input, up to 12); B4:M4 the names by formula; B5:M16 orders with both (diagonal empty); N gap; O Orders with this product, P Most often with, Q Share of this product's orders, R Partners tied for most, S Status. Checks A18:F24 (values in F): Orders on the Items tab (F19), Orders listed on the Baskets tab (up to 300) (F20), Item lines whose product is not on this list (F21), Rows showing Check (Items and this tab) (F22), Products listed (F23), CHECK (F24). Landscape.
- Deviations from the topic: R "Partners tied for most" and S Status added; Items Status adds "Repeat in this order (counted once)" and "Check: order # is text" (number stored as text); checks sit under the grid.

Key formulas (exact text in the file)
- Items C5 `=IF(A5="","",1)`; C6 `=IF(A6="","",IF(COUNTIF($A$5:A6,A6)=1,MAX($C$5:C5)+1,""))`
- Baskets A5 `=IFERROR(INDEX(Items!$A$5:$A$504,MATCH(ROWS(A$5:A5),Items!$C$5:$C$504,0)),"")`
- Baskets C5 `=IF(OR($A5="",C$4=""),"",IF(COUNTIFS(Items!$A$5:$A$504,$A5,Items!$B$5:$B$504,C$4)>0,1,0))` (core `=IF(COUNTIFS(order_col,order,product_col,product)>0,1,0)`)
- Pairs C5 (every off-diagonal grid cell; B5 is the diagonal) `=IF(OR($A5="",C$4=""),"",SUMPRODUCT((INDEX(Baskets!$C$5:$N$304,0,MATCH($A5,Baskets!$C$4:$N$4,0))=1)*(INDEX(Baskets!$C$5:$N$304,0,MATCH(C$4,Baskets!$C$4:$N$4,0))=1)))`
- O5 `=IF($A5="","",SUM(INDEX(Baskets!$C$5:$N$304,0,MATCH($A5,Baskets!$C$4:$N$4,0))))`
- P5 `=IF(OR($A5="",MAX(B5:M5)=0),"",INDEX($B$4:$M$4,MATCH(MAX(B5:M5),B5:M5,0)))`; Q5 `=IF(P5="","",MAX(B5:M5)/O5)`; R5 `=IF(P5="","",COUNTIF(B5:M5,MAX(B5:M5)))`
- F24 `=IF(AND(F19>0,F19=F20,F22=0,Items!G8=0),"OK","Check")`

Sample values (fictional; recalculated). Orders: 5001 Lavender candle, Wick trimmer, Gift box; 5002 Soap bar, Lip balm; 5003 Lavender candle, Wick trimmer; 5004 Cedar candle, Wick trimmer, Cedar candle; 5005 Soap bar, Lip balm, Gift box; 5006 Lavender candle; 5007 Lavender candle, Cedar candle, Wick trimmer; 5008 Lip balm, Soap bar; 5009 Gift box, Lavender candle; 5010 Soap bar (22 lines).
| Product | Lavender | Cedar | Wick | Soap | Lip balm | Gift box | Orders | Most often with | Share |
|---|---|---|---|---|---|---|---|---|---|
| Lavender candle | | 1 | 3 | 0 | 0 | 2 | 5 | Wick trimmer | 60.0% |
| Cedar candle | 1 | | 2 | 0 | 0 | 0 | 2 | Wick trimmer | 100.0% |
| Wick trimmer | 3 | 2 | | 0 | 0 | 1 | 4 | Lavender candle | 75.0% |
| Soap bar | 0 | 0 | 0 | | 3 | 1 | 4 | Lip balm | 75.0% |
| Lip balm | 0 | 0 | 0 | 3 | | 1 | 3 | Soap bar | 100.0% |
| Gift box | 2 | 0 | 1 | 1 | 1 | | 3 | Lavender candle | 66.7% |
Partners tied 1 on every row; Status OK. Items G4:G8 22, 10, 0, 0, 0; row 14 (5004 Cedar candle again) Repeat in this order (counted once). Baskets B5:B14 (products per order) 3, 2, 2, 2, 3, 1, 3, 2, 2, 1. Pairs F19:F24 10, 10, 0, 0, 6, OK.

Edge cases tested (all matched the Python model or the stated LibreOffice result)
- Lines shuffled (orders not grouped): same counts. 30 one-item orders: all pairs 0, Most often with blank.
- "soap bar" and "LIP BALM": counted (LibreOffice COUNTIFS ignores case; Soap bar + Lip balm 4). "Soap bar " with a trailing space: Check: extra spaces, not counted. A product not on the list, a blank order # and a blank product: Check, CHECK Check.
- Ties: Lavender candle with Soap bar, Lip balm and Gift box 2 each: Most often with Soap bar (first in the list), Partners tied 3.
- Duplicate list names (Lavender candle / lavender candle): both Check: name twice. Empty Items: all 0, CHECK Check.
- 12 products, 300 orders, 500 lines; 310 orders: Baskets shows 300, F19 310, F20 300, CHECK Check.
- Text "5001" next to the number 5001: Order no. numbered them once but the Baskets COUNTIFS left the text line out (pair 0), Check: order # is text.
- Probe on the template: `SUMPRODUCT(INDEX(...)*INDEX(...))` without the =1 tests gave #VALUE!; COUNTIF of Cedar candle on Items gave 3 lines while O6 gives 2 orders.

Pitfalls the post must state
- Counts only: "orders that contain both"; no lift, statistics or "bundles sell more" claims.
- Up to 12 products and 300 orders (say how to extend). If an export puts all items in one cell, split them first (link count-orders-with-exact-tag-comma-separated-cell or pull-size-and-scent-from-variations only if relevant).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Counts are too high: COUNTIF on the item lines counts lines, not orders (Cedar candle 3 lines vs 2 orders).
- Every pair cell shows #VALUE!: the two INDEX columns were multiplied without the =1 tests.
- A product is not counted: a trailing space or a name missing from the Pairs list (Items Status).
- Most often with picks one of several equal partners: ties return the first product in the list (R shows how many tie).

Hero: sheet "Pairs", rows 1 to 10, hide_cols "HIJKLMN" (the six used products stay in B:G), with O:S showing; Baskets A4:H14 as an inset.

---
Not built or tested (so not for the posts): a "The"-skipping or natural number sort, a sort done with **Data > Sort**, any platform's shipping rule, carrier rates, tax on shipping, secure codes, any shop's coupon import format, Paste Special done by hand, recalculation timing in Excel or Sheets, staffing or overtime, lift or other statistics, splitting one-cell item lists, typing by hand. Nothing here was run in Excel or Google Sheets.
