# Handoff 2026-11-29: five templates built and tested (LibreOffice Calc 24.2.7.2 only)

Topics: tasks/team-topics-2026-11-29.md. Hooks: tasks/hooks-2026-11-29.md (its "to confirm" values are superseded by the confirmed values below and the posts). Test record: tasks/briefs.md "2026-11-29 템플릿 제작/검증". Nothing was opened in Excel or Google Sheets.

## Topic 1
### (b) Confirmed values (replace every "to confirm" in hooks topic 1)

Layout of the built Customers tab: B1 as-of (11/28/2026), B2 Compare ("Same period" or "Full years", dropdown), B3 Last year through (optional, blank), B4 Last year from `=IF($I$21<>"OK","",DATE(YEAR(B1)-1,1,1))` = 01/01/2025, B5 Last year through `=IF($I$21<>"OK","",IF(B3<>"",B3,IF(B2="Full years",DATE(YEAR(B1)-1,12,31),DATE(YEAR(B1)-1,MONTH(B1),DAY(B1)))))` = 11/28/2025, B6 This year from `=IF($I$21<>"OK","",DATE(YEAR(B1),1,1))` = 01/01/2026, B7 This year through `=IF($I$21<>"OK","",B1)` = 11/28/2026. Customer rows 10 to 109, header row 9. So the hooks' `A7`, `$B$3`, `$B$4` become `A10`, `$B$4`, `$B$5` (last year) and `$B$6`, `$B$7` (this year).

- Opening 1 formula, exact cell B10 (Last year): `=IF(OR(TRIM(A10)="",$I$21<>"OK"),"",ROUND(SUMIFS(Orders!$C$5:$C$2004,Orders!$B$5:$B$2004,A10,Orders!$A$5:$A$2004,">="&$B$4,Orders!$A$5:$A$2004,"<"&$B$5+1),2))`. Short form for the post: `=SUMIFS(Orders!C:C,Orders!B:B,A10,Orders!A:A,">="&$B$4,Orders!A:A,"<"&$B$5+1)`. C10 (This year) is the same with `$B$6` and `$B$7`.
- Bridge numbers (confirmed, LibreOffice): last year $4,550.60 (I10) + New $1,529.99 (I11, 2 customers) + Up $476.25 (I12, 2) - Down $830.75 (I13 shows -$830.75, 2) - Lost $1,545.60 (I14 shows -$1,545.60, 1) + Same $0.00 (0) + No sales $0.00 (1) = $4,180.49 (I17) = sum of customers (I18) = whole log this year (I20); last year whole log (I19) $4,550.60; CHECK (I24) OK. Note: the built bridge shows Down and Lost as negative amounts (each bridge line is `=SUMIF($F$10:$F$109,"<label>",$D$10:$D$109)`), so "minus $830.75" reads as the line "- Down  -$830.75".
- Opening 2 label formula, exact cell F10: `=IF(TRIM(A10)="","",IF($I$21<>"OK","Check: settings (B1:B3)",IF(A10&""<>TRIM(A10),"Check: extra spaces in the name",IF(OR(ISNUMBER(FIND("*",A10)),ISNUMBER(FIND("?",A10)),ISNUMBER(FIND("~",A10)),LEFT(A10,1)="=",LEFT(A10,1)="<",LEFT(A10,1)=">"),"Check: * ? ~ = < or > in the name",IF(SUMPRODUCT(--($A$10:$A$109=A10))>1,"Check: customer listed twice",IF(AND(B10<=0,C10<=0),"No sales",IF(B10<=0,"New",IF(C10<=0,"Lost",IF(C10>B10,"Up",IF(C10<B10,"Down","Same"))))))))))`. Short form for the post (the label part only, columns as built): `=IF(AND(B10<=0,C10<=0),"No sales",IF(B10<=0,"New",IF(C10<=0,"Lost",IF(C10>B10,"Up",IF(C10<B10,"Down","Same")))))` where B is last year and C this year (same as the hook, with row 10 instead of 7). Oak Hollow Farm $1,545.60 to $0.00 Lost; Juniper Co. $0.00 to $649.99 and North Pine Studio $0.00 to $880.00, New, $1,529.99 between them: confirmed.
- Opening 3 (LibreOffice trap, scratch copy of the built file): `"<="&B7` gave $4,116.49, `"<"&B7+1` (the template) $4,180.49; the $64.00 order is Orders row 26: 11/28/2026 3:30 PM, Cedar Lane Cafe, $64.00, Status "This year". Note for writers: in the built sheet the this-year end cell is B7, not B1 (B7 = B1).
- Claims note items: 02/29/2028 as-of -> B5 shows 03/01/2027 and C5 says "02/29 has no match last year, so this is 03/01. Type 02/28 in B3 if you prefer."; typing 02/28/2027 in B3 is accepted (C5 "Typed in B3."). Willow Market: "No sales" with Same period; with B2 = "Full years" B5 = 12/31/2025, Willow Market $500.00 to $0.00 Lost, Maple Street Bakery $915.00 to $330.75 Down (Up in the same-period view), Harbor Gifts $510.00 to $310.00, last year total $5,770.60, CHECK OK.
- % change (E10 `=IF(OR(D10="",N(B10)<=0),"",D10/B10)`, format 0.0%): Birch & Bloom 52.6%, Cedar Lane Cafe -47.4%, Harbor Gifts -27.9%, Maple Street Bakery 20.3%, Oak Hollow Farm -100.0%, blank for New and No sales.
- Orders tab Status D5 (gray): Last year 9 rows, This year 10 rows, Outside both periods 4 rows (11/29/2025 Harbor Gifts $80.00, 12/01/2025 Willow Market $500.00, 12/10/2025 Maple Street Bakery $640.00, 11/29/2026 North Pine Studio $150.00), Check 0, Orders CHECK (G11) OK.

### (c) Example table (fictional candle maker's wholesale accounts, as of 11/28/2026, 01/01 to 11/28 in both years)

| Customer | Last year | This year | Change | Status |
|---|---|---|---|---|
| Birch & Bloom | $800.00 | $1,220.50 | $420.50 | Up |
| Cedar Lane Cafe | $1,500.00 | $789.25 | -$710.75 | Down |
| Harbor Gifts | $430.00 | $310.00 | -$120.00 | Down |
| Juniper Co. | $0.00 | $649.99 | $649.99 | New |
| Maple Street Bakery | $275.00 | $330.75 | $55.75 | Up |
| North Pine Studio | $0.00 | $880.00 | $880.00 | New |
| Oak Hollow Farm | $1,545.60 | $0.00 | -$1,545.60 | Lost |
| Willow Market | $0.00 | $0.00 | $0.00 | No sales |

Bridge: Last year $4,550.60, + New $1,529.99, + Up $476.25, - Down -$830.75, - Lost -$1,545.60 = This year $4,180.49 (CHECK OK).
Orders: 23 rows, 02/10/2025 to 11/29/2026 (sorted by date); Cedar Lane Cafe includes the 11/28/2026 3:30 PM $64.00 order; Harbor Gifts' 11/28/2025 $120.00 is in, 11/29/2025 $80.00 out; Maple Street Bakery's 12/10/2025 $640.00 is outside the window; North Pine Studio's 11/29/2026 $150.00 is after the as-of date.
Key formulas: B10/C10/F10 above; Change D10 `=IF(OR(B10="",C10=""),"",ROUND(C10-B10,2))`; bridge lines I11:I16 `=SUMIF($F$10:$F$109,"New",$D$10:$D$109)` (Up, Down, Lost, Same, No sales likewise); I17 `=ROUND(I10+SUM(I11:I16),2)`; log check I20 `=IF($I$21<>"OK","",ROUND(SUMIFS(Orders!$C$5:$C$2004,Orders!$A$5:$A$2004,">="&B6,Orders!$A$5:$A$2004,"<"&B7+1),2))`; CHECK I24 `=IF($I$21<>"OK","Check",IF(AND(I22=0,I23="OK",ROUND(I17-I18,2)=0,ROUND(I18-I20,2)=0,ROUND(I10-I19,2)=0),"OK","Check"))`.
Behaviors worth a Troubleshooting line (all LibreOffice-observed): a name in the log with a trailing space or not on the list -> Orders Status "Check: customer not on the Customers tab" and CHECK Check; capital letters do not matter; `*` in a list name -> Check (SUMIFS treated it as a wildcard); a refund-only year counts as nothing (New/Lost/No sales) and the bridge still adds up because each line sums the Change column.

### (d) make-heroes.py suggestion
Customers tab, A9:F17 plus the bridge H9:J17 (rows 9 to 17, columns A to J) shows the whole idea: labels colored (Up/New green, Down/Lost orange) next to the bridge from $4,550.60 to $4,180.49. If one range must be narrower, use H9:J17 with A9:F17 side by side at landscape crop; page 1 of the PDF export (LibreOffice Calc 24.2.7.2) already shows exactly this at the top. Caption: "PDF export from LibreOffice Calc 24.2.7.2, fictional sample data".

## Topic 2
### (b) Confirmed values (replace every "to confirm" in hooks topic 2)

- Key formula (Scans C5), exactly as in the file: `=IF(A5="","",IF(ISNUMBER(A5),IFERROR(INDEX(Items!$A$5:$A$504,MATCH(A5,Items!$H$5:$H$504,0)),TEXT(A5,"0")),TRIM(CLEAN(A5))))`. The hook's `=IF(ISNUMBER(A5),TEXT(A5,"000000000000"),TRIM(A5))` is NOT the template formula (see the decision above). For a simple explanation in the post: a scan that became a number is looked up by number value in Items column H (Code as number, H5 `=IF(ISTEXT(A5),IF(LEN(TRIM(A5))>15,"",IFERROR(IF(TEXT(VALUE(TRIM(A5)),REPT("0",LEN(TRIM(A5))))=TRIM(A5),VALUE(TRIM(A5)),""),"")),"")`) and gets the item's own code back; a text scan is cleaned with TRIM(CLEAN()).
- Tally formula (Items D5, Counted): `=IF(A5="","",SUMPRODUCT(--(Scans!$C$5:$C$2004=A5),Scans!$D$5:$D$2004))`. Blank Qty counts as 1 through Units (Scans D5): `=IF(OR(E5="",LEFT(E5,5)="Check"),"",IF(B5="",1,B5))`. (The hook's `SUMPRODUCT((Scans!$C$5:$C$2004=A5)*Scans!$D$5:$D$2004)` is not used: the comma form was chosen so the "" units on blank and Check rows count as 0; the `*` form was not tested.)
- Cedar candle 8 oz (012345678905): 9 units from 4 scans (one scan stored as the number 12345678905, one with Qty 6), expected 14, difference -5, Short. Confirmed.
- Plain COUNTIF on the raw scan column for the cedar candle: 4 in LibreOffice (scratch copy; trap numbers above). Confirmed.
- 3 scans had lost a leading zero (Scans I9 = 3; rows 6, 12, 21: 12345678905, 71234500017, 8800123455). Confirmed.
- 17 scans (Scans I4) come to 57 units (Scans I5); 56 matched to items (I6), 1 scan Not on list (I7, 999999999999, row 20), 1 unit on it (I8), 1 Spaces or Tab removed (I10, row 19 "400012000010 "), CHECK OK (I14, Items K13). Confirmed.
- Wool dryer balls (400012000010): 3 units in 3 scans, expected 9, 6 short (-6). Confirmed.
- Other items: Lavender soap bar (071234500017) 15 units in 4 scans, expected 22, -7; Holiday card 6 pack (008800123455, two leading zeros) 29 units in 5 scans (a typed 24 and a lost-zero scan with Qty 2), expected 30, -1; Gift wrap roll (400012000027) 0 units, expected 5, -5, Not scanned. Items K: 5 items listed, 56 counted, 80 expected, -24, 1 not scanned, 4 short, 0 over.
- Changed from the probe: lavender, card and dryer ball codes (the probe used 036000291452 and 070330506203, which look like real product codes); a fifth item (Gift wrap roll) was added to show Not scanned. No B1 code-length setting.
- TRIM vs Tab/Enter: the template uses CLEAN, and the post should say so (TRIM removes spaces; CLEAN removes Tab, line feed and other control characters).

### (c) Example table and key formulas for the script-writer

Scans (A Scan, B Qty → C Key, D Units, E Status, F Note), fictional:

| Row | Scan | Qty | Key | Units | Status | Note |
|---|---|---|---|---|---|---|
| 5 | 012345678905 | | 012345678905 | 1 | OK | |
| 6 | 12345678905 (number) | | 012345678905 | 1 | OK | Lost leading zero (fixed) |
| 7 | 012345678905 | | 012345678905 | 1 | OK | |
| 8 | 012345678905 | 6 | 012345678905 | 6 | OK | |
| 9 | 071234500017 | | 071234500017 | 1 | OK | |
| 10 | 071234500017 | 12 | 071234500017 | 12 | OK | |
| 11 | 071234500017 | | 071234500017 | 1 | OK | |
| 12 | 71234500017 (number) | | 071234500017 | 1 | OK | Lost leading zero (fixed) |
| 13 | 008800123455 | 24 | 008800123455 | 24 | OK | |
| 14-16 | 008800123455 (x3) | | 008800123455 | 1 each | OK | |
| 17-18 | 400012000010 (x2) | | 400012000010 | 1 each | OK | |
| 19 | "400012000010 " (trailing space) | | 400012000010 | 1 | OK | Spaces or Tab removed |
| 20 | 999999999999 | | 999999999999 | 1 | Not on list | |
| 21 | 8800123455 (number) | 2 | 008800123455 | 2 | OK | Lost leading zero (fixed) |

Items result:

| Code | Item | Expected | Counted | Scans | Difference | Status |
|---|---|---|---|---|---|---|
| 012345678905 | Cedar candle 8 oz | 14 | 9 | 4 | -5 | Short |
| 071234500017 | Lavender soap bar | 22 | 15 | 4 | -7 | Short |
| 008800123455 | Holiday card 6 pack | 30 | 29 | 5 | -1 | Short |
| 400012000010 | Wool dryer balls | 9 | 3 | 3 | -6 | Short |
| 400012000027 | Gift wrap roll | 5 | 0 | 0 | -5 | Not scanned |

Other formulas: Status Scans E5 (OK / Not on list via `ISNUMBER(MATCH(C5,Items!$A$5:$A$504,0))`, Check for Qty text / below 1 / not whole, Qty with no scan, over 15 digits, decimal scan); Note F5 `=IF(E5<>"OK","",IF(ISNUMBER(A5),IF(LEN(C5)>LEN(TEXT(A5,"0")),"Lost leading zero (fixed)",""),IF(LEN(C5)<>LEN(A5),"Spaces or Tab removed","")))`; Items Scans E5 `=IF(A5="","",SUMPRODUCT((Scans!$C$5:$C$2004=A5)*(Scans!$E$5:$E$2004="OK")))`; Difference F5 `=IF(OR(A5="",C5="",NOT(ISNUMBER(C5))),"",D5-C5)`; Status G5 (Check: code stored as a number / no code / spaces in the code / code listed twice / same number as another code / Expected is text / Expected below 0; else Not scanned, Counted, Matches expected, Short, Over); CHECK Scans I14 `=IF(AND(I11=0,I12=0,I13=0,I5=I6+I8),"OK","Check")`. Inputs: Scans A5:B2004, Items A5:C504 (column A formatted as text `@` on both tabs). Functions used: ABS, AND, CLEAN, COUNTA, COUNTIF, IF, IFERROR, INDEX, INT, ISNUMBER, ISTEXT, LEFT, LEN, MATCH, N, NOT, OR, REPT, SUM, SUMIF, SUMPRODUCT, TEXT, TRIM, VALUE.
Help pages to add if the post explains them (not in the topics list): CLEAN, TEXT, INDEX/MATCH, IFERROR; writers check the URLs themselves.

### (d) make-heroes.py suggestion

Best single view: the Items tab A4:G9 (Code, Item, Expected, Counted, Scans, Difference, Status, with the orange Short/Not scanned fills) beside summary J4:K13. Second choice / step image: Scans A4:F21 plus H4:I14 (shows the three Lost leading zero (fixed) notes, the Spaces or Tab removed note, the Not on list row and CHECK OK). Caption source: "PDF export from LibreOffice Calc 24.2" (hero not rendered by me).

## Topic 3
### (b) Confirmed values for tasks/hooks-2026-11-29.md topic 3

- Box formula: the built cell is C7 `=IF(LEFT(E7,6)="Packed",IF(AND(D6>0,ROUND(D6+B7,4)>$B$3),C6+1,C6),C6)`; short form for the opening: `=IF(ROUND(D6+B7,4)>$B$3,C6+1,C6)` is still accurate as the core test (the built cell adds a `D6>0` guard so a too-heavy first item or a box after skipped rows never leaves an empty box, and skips rows that are not packed). C6 is the number 1. Header row 5, items from row 6, B3 usable weight.
- Running weight: D7 `=IF(LEFT(E7,6)="Packed",IF(C7=C6,ROUND(D6+B7,4),B7),D6)`; short form `=IF(C7=C6,D6+B7,B7)` is right in meaning; the built cell rounds to 4 decimals. D6 `=IF(LEFT(E6,6)="Packed",B6,0)`.
- B3 is `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),ROUND(B1-B2,4),"")` (not plain `=B1-B2`), 18.50 in the sample.
- Sample: 10 items pack into **5 boxes**; the first box weighs exactly **20.00 lb packed** (18.50 lb of items, Status At the limit).
- 62.1 lb over 18.5 usable lb: weight-only minimum **4** (Boxes K8 `=IF(OR(Packing!H5<>"OK",K6=0),0,ROUNDUP(K6/Packing!B3,0))`), but list order needs **5**; the 19.00 lb Cast iron doorstop is "Packed alone: too heavy for one box" in box 4 (20.50 lb packed, Over the limit).
- Box 2 holds Soap crate (12 bars), Wool throw and Ceramic vase: 16.50 lb of items, **18.00 lb packed** (Boxes D6).
- Heaviest first: **5 boxes to 4** (checked in LibreOffice on a scratch copy with the rows in descending weight order; total packed 68.10 lb instead of 69.60 lb).
- Unrounded running weight in LibreOffice: same boxes as with ROUND in all 45 edge cases (LibreOffice only; no Excel/Sheets claim).
- Item name note: the sample names are "Mug set (4)", "Soap crate (12 bars)", "Card bundle (25)", "Cast iron doorstop" (the probe said Mug set, Soap crate, Card bundle, Doorstop).

### (c) Example table and key formulas for the script-writer

Settings: Limit per box (B1) 20.00 lb; Box and packing weight (B2) 1.50 lb; Usable weight for items (B3) 18.50 lb.

| Item | Weight (lb) | Box | Weight in box so far (lb) | Status |
|---|---|---|---|---|
| Mug set (4) | 6.40 | 1 | 6.40 | Packed |
| Cast iron trivet | 9.80 | 1 | 16.20 | Packed |
| Candle 3 pack | 2.30 | 1 | 18.50 | Packed |
| Soap crate (12 bars) | 7.40 | 2 | 7.40 | Packed |
| Wool throw | 3.20 | 2 | 10.60 | Packed |
| Ceramic vase | 5.90 | 2 | 16.50 | Packed |
| Jam gift box | 4.40 | 3 | 4.40 | Packed |
| Cast iron doorstop | 19.00 | 4 | 19.00 | Packed alone: too heavy for one box |
| Card bundle (25) | 1.10 | 5 | 1.10 | Packed |
| Tea tin set | 2.60 | 5 | 3.70 | Packed |

Boxes tab:

| Box | Items | Item weight (lb) | Packed weight (lb) | Room left (lb) | Status | First item | Last item |
|---|---|---|---|---|---|---|---|
| 1 | 3 | 18.50 | 20.00 | 0.00 | At the limit | Mug set (4) | Candle 3 pack |
| 2 | 3 | 16.50 | 18.00 | 2.00 | OK | Soap crate (12 bars) | Ceramic vase |
| 3 | 1 | 4.40 | 5.90 | 14.10 | OK | Jam gift box | Jam gift box |
| 4 | 1 | 19.00 | 20.50 | 0.00 | Over the limit: one item too heavy | Cast iron doorstop | Cast iron doorstop |
| 5 | 2 | 3.70 | 5.20 | 14.80 | OK | Card bundle (25) | Tea tin set |

Summary (Boxes K4:K13): Boxes 5; Items packed 10; Total item weight 62.10; Total packed weight 69.60; Weight-only minimum 4; Over the limit 1; At the limit 1; Not listed 0; Packing CHECK OK; CHECK OK. Packing H5:H12: Settings OK, 10, 62.10, 5, 1 too heavy, 0 Check rows, 0 below row 505, CHECK OK.

Key formulas (exact cell text):
- Packing E7 (Status): `=IF(AND(TRIM(A7)="",B7=""),IF(SUMPRODUCT((TRIM(A8:A505)<>"")+(B8:B505<>""))=0,"","Check: blank row inside the list (skipped)"),IF($H$5<>"OK","Check: settings (B1:B2)",IF(TRIM(A7)="","Check: no item name",IF(NOT(ISNUMBER(B7)),"Check: weight blank or text",IF(B7<=0,"Check: weight 0 or below",IF(ROUND(B7-$B$3,4)>0,"Packed alone: too heavy for one box","Packed"))))))`
- Packing H8 (Boxes): `=IF(H6=0,0,MAX(C6:C505))`; H5 settings `=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),IF(AND(B1>0,B2>=0,B2<B1),"OK","Check"),"Check")`
- Boxes B5 (Items): `=IF(A5="","",COUNTIFS(Packing!$C$6:$C$505,A5,Packing!$E$6:$E$505,"Packed*"))`
- Boxes C5 (Item weight): `=IF(A5="","",ROUND(SUMIFS(Packing!$B$6:$B$505,Packing!$C$6:$C$505,A5,Packing!$E$6:$E$505,"Packed*"),4))`
- Boxes D5 (Packed weight): `=IF(A5="","",ROUND(C5+Packing!$B$2,2))`
- Boxes E5 (Room left): `=IF(A5="","",MAX(0,ROUND(Packing!$B$3-C5,2)))`
- Boxes F5: `=IF(A5="","",IF(ROUND(C5-Packing!$B$3,4)>0,"Over the limit: one item too heavy",IF(ROUND(C5-Packing!$B$3,4)=0,"At the limit","OK")))`
- Boxes H5 (Last item): `=IF(A5="","",INDEX(Packing!$A$1:$A$505,SUMPRODUCT(MAX((Packing!$C$6:$C$505=A5)*(LEFT(Packing!$E$6:$E$505,6)="Packed")*ROW(Packing!$C$6:$C$505)))))`; G5 (First item) is the same with `LARGE(...,B5)` in place of `MAX(...)`.
- Troubleshooting material (all observed in LibreOffice): blank row inside the list shows "Check: blank row inside the list (skipped)" and CHECK Check, packing still continues past it; "3.2 lb" typed as text shows "Check: weight blank or text"; B2 at or above B1 shows "Check: settings (B1:B2)" on every row and 0 boxes; more than 200 boxes shows Boxes not listed and CHECK Check; a too-heavy item is packed alone and its box is Over the limit (raise the limit or ship it separately).

### (d) make-heroes.py suggestion

Best single view: the Boxes tab A4:H9 (the five boxes with packed weights, At the limit in green, Over the limit in red, first and last item), optionally with J4:K8 (Boxes 5, weight-only minimum 4). Second choice for a step image: Packing A1:E15 (settings plus the 10 items with Box and running weight, the doorstop row in orange). Source PDF page from LibreOffice: page 1 is Packing (pages 1 to 16), page 17 is the first Boxes page (17 to 21), page 22 How to use of tidy-tabs-pack-boxes-by-weight.pdf (scratch t3/pdf2/).

## Topic 4
### (b) Confirmed values (replace every "to confirm" in hooks topic 4)
Layout changed from the hook draft: the lines are in **B2 (units) and B3 (margin)**, not B1/B2 (row 1 holds the small headers Line used / Median (OK rows) / Your own line (optional)). Ranges run to row 205 (200 products), not 55. The units median reads helper column I (units of OK rows only) so a Check row is not in either median.
- Units line used, Products B2: `=IF(D2="",C2,D2)`; median, C2: `=IF(COUNT($I$6:$I$205)=0,"",MEDIAN($I$6:$I$205))`; I6: `=IF(H6="OK",B6,"")`.
- Margin line used, B3: `=IF(D3="",C3,D3)`; median, C3: `=IF(COUNT($E$6:$E$205)=0,"",MEDIAN($E$6:$E$205))`.
- Margin, E6: `=IF(H6<>"OK","",ROUND((C6-D6)/C6,4))`. Profit, F6: `=IF(H6<>"OK","",ROUND(C6-D6,2))`.
- Group, G6: `=IF(OR(H6<>"OK",$L$10<>"OK"),"",IF(B6>=$B$2,IF(E6>=$B$3,"High units, high margin","High units, low margin"),IF(E6>=$B$3,"Low units, high margin","Low units, low margin")))`. Short form for the opening: `=IF(B6>=$B$2,IF(E6>=$B$3,"High units, high margin","High units, low margin"),IF(E6>=$B$3,"Low units, high margin","Low units, low margin"))`.
- Status, H6: `=IF(AND(TRIM(A6)="",B6="",C6="",D6=""),"",IF(TRIM(A6)="","Check: no product name",IF(NOT(ISNUMBER(B6)),"Check: units blank or text",IF(B6<0,"Check: units below 0",IF(NOT(ISNUMBER(C6)),"Check: revenue blank or text",IF(C6<=0,"Check: revenue 0 or less (no margin)",IF(NOT(ISNUMBER(D6)),"Check: cost blank or text",IF(D6<0,"Check: cost below 0","OK"))))))))`.
- Summary count, B5: `=COUNTIF(Products!$G$6:$G$205,A5)`; profit, F5: `=ROUND(SUMIF(Products!$G$6:$G$205,A5,Products!$F$6:$F$205),2)` (C5 units, D5 revenue, E5 cost the same SUMIF on B, C, D); share G5: `=IF($F$9<=0,"",F5/$F$9)`; group margin H5: `=IF(D5<=0,"",F5/D5)`. A5:A8 hold the four labels in the order High/high, High/low, Low/high, Low/low; row 9 Total.
- Checks: Products L7 on the units line `=COUNTIFS(H6:H205,"OK",B6:B205,$B$2)` (sample 0), L8 on the margin line (sample 0), L10 lines, L11 CHECK; Summary K6 Groups add up `=IF(AND(B9=K4,ROUND(F9-K5,2)=0),"OK","Check")`, K8 CHECK `=IF(AND(K6="OK",K7="OK"),"OK","Check")` (sample OK).
- Sample results (LibreOffice 24.2.7.2): units line **255** (median of 210 and 300), margin line **55.00%** (median of 50.00% and 60.00%); no product sits on either line (L7 0, L8 0); Beeswax wraps and Lip balm at exactly 50.00% fall under the margin line.
  - High units, high margin: **3** (Cedar candle 8 oz 64.96%, Lavender soap bar 70.00%, Holiday card 6 pack 77.78%), 1,447 units, revenue $21,956.00, cost $6,552.80, profit **$15,403.20**, 56.31% of profit, group margin 70.15%.
  - High units, low margin: **2** (Lip balm 820 units 50.00%, Wax melts 300 units 37.50%), 1,120 units, $5,680.00, $3,140.00, profit **$2,540.00**, 9.28%, 44.72%.
  - Low units, high margin: **2** (Wool dryer balls 95 units 71.74%, Gift box set 120 units 60.00%), 215 units, $7,585.00, $2,777.50, profit **$4,807.50**, 17.57%, 63.38%.
  - Low units, low margin: **3** (Beeswax wraps 50.00%, Ceramic mug 43.33%, Linen tea towel 38.89%), 438 units, $10,404.00, $5,798.00, profit **$4,606.00**, 16.84%, 44.27%.
  - Total: 10 products, 3,220 units, revenue $45,625.00, cost $18,268.30, profit **$27,356.70**, margin 59.96%, CHECK OK.
- Hook opening 2: lip balm 820 units at 50.00%, wool dryer balls 95 units at 71.74% (confirmed). Hook opening 3: High units, low margin 2 products $2,540.00; Low units, high margin 2 products $4,807.50 (confirmed); use `=COUNTIF(Products!$G$6:$G$205,A5)` and `=SUMIF(Products!$G$6:$G$205,A5,Products!$F$6:$F$205)` (the file wraps the SUMIF in ROUND(...,2)).
- Typed lines example (scratch copy): D2 300 and D3 50% gives 4 / 1 / 3 / 2 products and $17,043.20 / $900.00 / $6,067.50 / $3,346.00; Wax melts (300) counts High units because of >=. Odd count example: dropping Wax melts (9 products) gives a units line of 210 (Ceramic mug sits on it and is High units).
- Scatter chart: confirmed rendering in the LibreOffice 24.2.7.2 PDF export (Summary page): dots for the OK products, a dashed red vertical units line at 255 and a dashed blue horizontal margin line at 55.00%; line ends in Summary J13:L16 (`=Products!$B$2` with MIN/MAX of the margins, MIN/MAX of the OK units with `=Products!$B$3`). Not checked in Excel or Sheets, so the post can say "the chart in the download (seen in LibreOffice's PDF export)".

### (c) Example table (fictional candle and soap maker, 2026) and key formulas
| Product | Units sold | Revenue | Cost | Margin | Profit | Group |
|---|---|---|---|---|---|---|
| Cedar candle 8 oz | 412 | $9,876.00 | $3,460.80 | 64.96% | $6,415.20 | High units, high margin |
| Lavender soap bar | 655 | $5,240.00 | $1,572.00 | 70.00% | $3,668.00 | High units, high margin |
| Holiday card 6 pack | 380 | $6,840.00 | $1,520.00 | 77.78% | $5,320.00 | High units, high margin |
| Wool dryer balls | 95 | $2,185.00 | $617.50 | 71.74% | $1,567.50 | Low units, high margin |
| Beeswax wraps | 140 | $2,520.00 | $1,260.00 | 50.00% | $1,260.00 | Low units, low margin |
| Ceramic mug | 210 | $6,300.00 | $3,570.00 | 43.33% | $2,730.00 | Low units, low margin |
| Linen tea towel | 88 | $1,584.00 | $968.00 | 38.89% | $616.00 | Low units, low margin |
| Lip balm | 820 | $3,280.00 | $1,640.00 | 50.00% | $1,640.00 | High units, low margin |
| Gift box set | 120 | $5,400.00 | $2,160.00 | 60.00% | $3,240.00 | Low units, high margin |
| Wax melts | 300 | $2,400.00 | $1,500.00 | 37.50% | $900.00 | High units, low margin |

Summary table for the post:
| Group | Products | Profit | Share of profit |
|---|---|---|---|
| High units, high margin | 3 | $15,403.20 | 56.31% |
| High units, low margin | 2 | $2,540.00 | 9.28% |
| Low units, high margin | 2 | $4,807.50 | 17.57% |
| Low units, low margin | 3 | $4,606.00 | 16.84% |
| Total | 10 | $27,356.70 | 100.00% |

Key formulas: B2 `=IF(D2="",C2,D2)`, C2 `=MEDIAN($I$6:$I$205)` (with the COUNT guard), C3 `=MEDIAN($E$6:$E$205)`, E6 margin, G6 group, Summary COUNTIF/SUMIF (all exact text above). Writers: margin is (revenue - cost) / revenue from the reader's own cost; labels describe, they do not advise; medians put about half on each side; LibreOffice Calc only.

### (d) make-heroes.py suggestion
Best result view: **Summary!A1:H30** (lines B1:B2, the four-group table A4:H9 with colored labels, and the scatter chart A11 to about H30), as in the PDF page 6 of the LibreOffice export. If the hero renderer draws cells only (no chart), use **Products!A5:H15** (the 10 products with Margin, Profit and the color-coded Group column) and add **Summary!A4:H9** as a second panel. Caption: "PDF export from LibreOffice Calc 24.2" / sample products fictional.

## Topic 5
### (b) Confirmed values (replace every "to confirm" in hooks topic 5)

- Date, Export C5: `=IF(AND(A5="",B5=""),"",IF(AND(ISNUMBER(A5),N(A5)>=0),INT(A5),"Check"))`, format MM/DD/YYYY. Core: `=INT(A5)`.
- Time, Export D5: `=IF(AND(A5="",B5=""),"",IF(AND(ISNUMBER(A5),N(A5)>=0),ROUND(MOD(A5,1)*86400,0)/86400,"Check"))`, format h:mm AM/PM. Core: `=MOD(A5,1)`, rounded to the second because plain MOD showed 2:29 PM in LibreOffice (see the trap above). **The opening's "and the time is `=MOD(A5,1)` ... splits into 2:30 PM" needs this caveat: in LibreOffice, plain `=MOD(A5,1)` formatted h:mm AM/PM displayed 2:29 PM for a 2:30 PM sale; `=ROUND(MOD(A5,1)*86400,0)/86400` displayed 2:30 PM.**
- Hour, E5: `=IF(ISNUMBER(D5),HOUR(A5),"")`. Minutes after midnight, F5: `=IF(ISNUMBER(D5),INT(ROUND(D5*86400,0)/60),"")` (2:30 PM = 870). Changed from the topic's `ROUND(D5*1440,0)`, which rounds 2:30:45 PM up to 871 and 11:59:59 PM to 1440.
- Daily total, By day B5: `=IF(A5="","",SUMIFS(Export!$B$5:$B$2004,Export!$A$5:$A$2004,">="&A5,Export!$A$5:$A$2004,"<"&A5+1))` (bounded ranges, not B:B/A:A; the post may show the short form `=SUMIFS(B:B,A:A,">="&D5,A:A,"<"&D5+1)` only as a pattern, it was not run with whole columns): Fri 11/27/2026 $85.50, Sat 11/28/2026 $115.00. Confirmed.
- Same total on the split Date column, C5: `=IF(A5="","",SUMIFS(Export!$B$5:$B$2004,Export!$C$5:$C$2004,A5))`: $85.50, $115.00 (F says OK). Orders D5 (COUNTIFS window): 3 and 4.
- Wrong way on the built By day tab, E5: `=IF(A5="","",SUMIFS(Export!$B$5:$B$2004,Export!$A$5:$A$2004,A5))`: **$0.00 for 11/27/2026 and $69.75 for 11/28/2026** (LibreOffice 24.2.7.2; $60.00 + $9.75, the two 12:00 AM rows). Confirmed.
- Time match: `=MOD(A6,1)=TIME(14,30,0)` FALSE on both 2:30 PM rows (scratch copy, LibreOffice); `=F6=870` TRUE on both; the hook's `=ROUND(MOD(A5,1)*1440,0)=870` also gives TRUE on both 2:30 PM rows (scratch copy: 870), but gives 1440 for 11:59:59 PM; the template formula is `INT(ROUND(D5*86400,0)/60)`. By day B3 2:30 PM: I13 870, I14 2 sales, I15 $30.50.
- Text date-time "11/28/2026 4:45 PM" in Export A11: C11/D11 "Check", G11 "Check: text, not a date-time"; Export J6 = 1, J7 = 1; By day I9 1, I10 $20.00.
- Totals: Export J4 7 rows OK, J5 $200.50, J8 11/27/2026, J9 11/28/2026, J10 2 days, J12 Check; By day I5 $200.50, I6 $200.50, I7 $0.00, I8 OK, I11 Check (text row).
- 12:00 AM: counted on its own day (11/28 $60.00 row); 11:59:59 PM counted on 11/27. Confirmed in the sample and edge runs.

### (c) Example table for the script-writer (fictional US sample)

| Date and time (A) | Amount (B) | Date `=INT(A5)` (C) | Time (D) | Minutes (F) | Status (G) |
|---|---|---|---|---|---|
| 11/27/2026 9:15 AM | $42.00 | 11/27/2026 | 9:15 AM | 555 | OK |
| 11/27/2026 2:30 PM | $18.50 | 11/27/2026 | 2:30 PM | 870 | OK |
| 11/27/2026 11:59 PM (11:59:59) | $25.00 | 11/27/2026 | 11:59 PM | 1439 | OK |
| 11/28/2026 12:00 AM | $60.00 | 11/28/2026 | 12:00 AM | 0 | OK |
| 11/28/2026 10:05 AM | $33.25 | 11/28/2026 | 10:05 AM | 605 | OK |
| 11/28/2026 2:30 PM | $12.00 | 11/28/2026 | 2:30 PM | 870 | OK |
| 11/28/2026 4:45 PM (text) | $20.00 | Check | Check | | Check: text, not a date-time |
| 11/28/2026 (no time) | $9.75 | 11/28/2026 | 12:00 AM | 0 | OK |

| Day | Sales (window) | Orders | Matching the date alone (wrong) |
|---|---|---|---|
| Fri 11/27/2026 | $85.50 | 3 | $0.00 |
| Sat 11/28/2026 | $115.00 | 4 | $69.75 |

Key formulas: Export C5 / D5 / F5 and By day B5 / E5 above. Export A column is US_DATETIME (mm/dd/yyyy h:mm AM/PM), so 11:59:59 PM displays as 11:59 PM there and in D.

### (d) make-heroes.py note

Best result view: By day tab A4:F6 with the summary H4:I11 (rows 1 to 16 fit one landscape page; shows $85.50 / $115.00 next to the wrong $0.00 / $69.75 in orange). Second choice: Export A4:G12 (the split columns, the red Check text row). Suggested call: sheet "By day", range A1:I16 (or A4:F6 if the hero should be only the table). Trial PDF page images in scratch: pdf2/p0.png (Export), pdf2/p55.png (By day). Nothing written to images/.
