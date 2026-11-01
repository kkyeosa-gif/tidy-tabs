# Handoff 2026-11-02: five templates built and checked (LibreOffice Calc 24.2.7.2 only)

Status: topics (tasks/team-topics-2026-11-02.md) done; hooks are in tasks/hooks-2026-11-02.md; the posts are the next step. No topic was dropped. Five new functions in `scripts/build-templates.py` (`workshop_registrations`, `price_after_card_fee`, `top_unpaid_invoices`, `number_from_text_units`, `payment_reminders`, added before the `if __name__` block and registered after `gift_set_cost()`). Only these five were run (a small runner that loads the script and calls just them, with `PYTHONDONTWRITEBYTECODE=1`, using `/usr/bin/python3` because the default `python3` has no openpyxl); older .xlsx were not rewritten. Nothing was committed. New in `git status`: the five .xlsx, the edited script, tasks/briefs.md, tasks/hooks-2026-11-02.md, this file.

Exact version string used for `tested_in`: **LibreOffice 24.2.7.2 420(Build:2)** (Calc, Linux, headless). Write "LibreOffice Calc 24.2.7.2". **Nothing was opened in Excel or Google Sheets and nothing may be claimed about them** (including COUNTIF with an array criterion inside SUMPRODUCT, whole-cell COUNTIF matching, VALUE on odd text, TEXT format codes and ISERROR inside SUMPRODUCT).

Method: edited copies made with openpyxl (inputs only), recalculated with `soffice --headless --convert-to xlsx` into a scratch dir, read back with `data_only=True`, compared with independent Python models that read only the typed inputs. 103 workbooks (per template, including the sample copy: 16 workshops, 46 card fee, 19 invoices, 8 weights, 14 reminders), 9,517 cell comparisons. Final run: 0 mismatches and no error value anywhere, including invalid-input copies. The 5 committed files were also recalculated and exported to PDF: no error values; every page is US Letter (workshops: landscape Workshops tab, portrait Signups and How to use; card fee: portrait; invoices: portrait Invoices, landscape Top 3; weights: portrait; reminders: landscape Invoices, portrait Settings and How to use). Bugs found by the tests and fixed (re-run clean): reminders H column had a broken quote (#VALUE! in every row); the weights formula turned `1/2 lb` into the date number 46024 and `12-oz` into -12 (guards added). Sample data fictional, dollars US. Yellow (FFF2CC) = inputs, gray (EDEDED) = formulas. Each file has a "How to use" sheet and one CHECK cell.

Standard `tested_in` wording for all five: `LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets.` Posts must give no advice (no refund or waiver terms, no statement on passing card fees to customers, no collections advice).

Deviations from the topics file sketches (explained per section): workshops summary cells moved (K1..K7, with K6 unmatched and K7 CHECK) and "Check seats" guards added; card fee gets an input check B5 and a CHECK in B14; top 3 uses a tie key G and ranks on it (minus, not plus the row term) and its opening changed; weights typed text is in column B and the formula gained unit guards; reminders gained "there" for a blank name, "1 day", a sign-off cell and Check messages.

---

## 1. Workshop registrations and seats left (`templates/tidy-tabs-workshop-registrations.xlsx`, function `workshop_registrations`)
- Final title: How to Track Workshop Registrations and Seats Left in Excel and Google Sheets
- Slug: `workshop-registrations-seats-left-waitlist-countifs-excel-google-sheets`
- Sheets: Workshops (landscape, header row 4), Signups (header row 4), How to use.
- Inputs (yellow): Workshops A5:C14 (name, date, seats; 4 filled), Signups A5:D204 (name, workshop, email, payment; 28 filled, example.com emails; workshop dropdown and Paid/Unpaid dropdown are warning-type so typed values are allowed).
- Formulas (row 5):
  - D5 `=IF(A5="","",COUNTIFS(Signups!$B$5:$B$204,A5,Signups!$D$5:$D$204,"Paid"))`
  - E5 `=IF(A5="","",COUNTIF(Signups!$B$5:$B$204,A5))`
  - F5 `=IF(A5="","",IF(ISNUMBER(C5),MAX(0,C5-D5),"Check seats"))`
  - G5 `=IF(A5="","",IF(ISNUMBER(C5),MAX(0,E5-C5),"Check seats"))`
  - H5 `=IF(A5="","",IF(NOT(ISNUMBER(F5)),"Check seats",IF(F5=0,"Full",IF(F5<=2,"Almost full","Open"))))`
  - Summary J:K: K1 `=COUNTA(A5:A14)`, K2 `=SUM(C5:C14)`, K3 `=SUM(D5:D14)`, K4 `=SUM(F5:F14)`, K5 `=SUM(G5:G14)`, K6 `=SUMPRODUCT((Signups!B5:B204<>"")*(COUNTIF(A5:A14,Signups!B5:B204)=0))` (worked; no helper column needed), K7 CHECK `=IF(AND(K6=0,COUNT(C5:C14)=K1,SUMPRODUCT((A5:A14<>"")*(D5:D14>C5:C14))=0),"OK","Check")`
- Sample result (every cell equals the model):
  | Workshop | Seats | Paid | All sign-ups | Seats left | Waitlist | Status |
  | Candle Pouring 101 | 8 | 7 | 8 | 1 | 0 | Almost full |
  | Intro to Block Printing | 10 | 5 | 6 | 5 | 0 | Open |
  | Embroidery Basics | 6 | 6 | 7 | 0 | 1 | Full |
  | Wreath Making | 12 | 6 | 7 | 6 | 0 | Open |
  Summary: 4 workshops, 36 seats, 24 paid, 12 seats left, 1 on the waitlist, 0 unmatched, K7 OK.
- Edge findings (16 workbooks, 912 comparisons, 0 mismatches):
  - Sign-ups equal to seats (8 paid on 8 seats): 0 left, Full, waitlist 0, OK. One over (9 paid on 8): 0 left, waitlist 1, K7 Check (paid above seats). 8 paid plus 1 unpaid: Full, waitlist 1, K7 OK. Zero sign-ups: all seats left, Open, OK.
  - Capitalization: `candle pouring 101` and `CANDLE POURING 101` are counted, and `paid` / `PAID` count as Paid. A trailing space on the workshop name (`Wreath Making `) is not counted anywhere: K6 counts them (2 in the test) and K7 shows Check. A misspelling (`Candel Pouring 101`) does the same (K6 1, Check).
  - Payment typed `Paid ` with a trailing space is NOT counted as paid and K7 stays OK (no warning); the post should tell readers to use the dropdown.
  - 0 seats with sign-ups: Full, waitlist equals all sign-ups, K7 Check when someone has paid. Seats typed as text (`8`) or left blank: F:H show "Check seats", K7 Check. Duplicate workshop name listed twice: both rows count every matching sign-up (same D and E), so totals double count, K7 stays unless paid exceeds seats (test: Check because of a 2-seat row). Blank rows in the middle of Signups: skipped. A sign-up with a workshop but no payment: counted in E, not in D.
  - 200 sign-ups across 4 workshops, 10 workshops, rows beyond the slots: 200 and 10 matched; more than 200 sign-ups or 10 workshops were not run.
- Tested_in: `LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets.`
- Hook paragraph: Count paid sign-ups for a workshop with `=COUNTIFS(Signups!$B$5:$B$204,A5,Signups!$D$5:$D$204,"Paid")`, then show the seats left with `=MAX(0,C5-D5)`. Everyone who signs up beyond the seat count goes on the waitlist, which you get with `=MAX(0,E5-C5)` from a separate count of all sign-ups.
- Limits: no refunds, cancellation terms or waivers. COUNTIF treats `*` and `?` in a workshop name as wildcards (not run). Not run: Excel, Sheets.

---

## 2. Price that nets a target after a card fee (`templates/tidy-tabs-price-after-card-fee.xlsx`, function `price_after_card_fee`)
- Final title: How to Set a Price That Still Nets the Amount You Want After a Card Fee in Excel and Google Sheets
- Slug: `price-to-net-target-after-card-fee-excel-google-sheets`
- Sheets: Calculator, How to use.
- Inputs (yellow): B2 `$100.00`, B3 `2.9%` (format 0.0%), B4 `$0.30`, and the five targets A17:A21 ($25.00, $48.00, $100.00, $250.00, $1,250.00). The numbers are examples; the post tells readers to type their own processor's rates (no claim about any provider).
- Formulas:
  - B5 `=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B4)),IF(AND(B2>=0,B3>=0,B3<1,B4>=0),"OK","Percent: 0% up to under 100%. Amounts: 0 or more"),"Type a number in each yellow cell")`
  - B6 `=IF($B$5<>"OK","",ROUND((B2+B4)/(1-B3),2))`; core `=ROUND((B2+B4)/(1-B3),2)`
  - B7 `=IF(B6="","",ROUND(B6*B3+B4,2))`; B8 `=IF(B6="","",ROUND(B6-B7,2))`; B9 `=IF(B6="","",ROUND(B8-B2,2))`
  - B10 `=IF($B$5<>"OK","",ROUND(B2*(1+B3)+B4,2))`; B11 `=IF(B10="","",ROUND(B10-(B10*B3+B4),2))`; B12 `=IF(B10="","",ROUND(B11-B2,2))`
  - B13 `=IF(B5<>"OK","Check inputs",IF(B9>=0,"OK",IF(B9>=-0.01,"1 cent short from rounding","Check")))`
  - B14 CHECK `=IF(AND(B5="OK",ISNUMBER(B6),COUNTIF(E17:E21,"<-0.01")=0,COUNT(E17:E21)=COUNT(A17:A21)),"OK","Check")`
  - Table row 17: B17 `=IF(OR($B$5<>"OK",NOT(ISNUMBER(A17))),"",ROUND((A17+$B$4)/(1-$B$3),2))`, C17 `=IF(B17="","",ROUND(B17*$B$3+$B$4,2))`, D17 `=IF(B17="","",ROUND(B17-C17,2))`, E17 `=IF(B17="","",ROUND(D17-A17,2))`
- Sample result: price $103.30, fee $3.30, you keep $100.00, difference $0.00; shortcut price $103.20 keeps $99.91, short $0.09; status OK, B14 OK. Table: $25.00 -> $26.06 (fee $1.06); $48.00 -> $49.74 ($1.74); $100.00 -> $103.30 ($3.30); $250.00 -> $257.78 ($7.78); $1,250.00 -> $1,287.64 ($37.64); every difference $0.00.
- ROUND vs ROUNDUP (decided in the build): the run kept ROUND. Over 238 checked prices (the 46 workbooks' B2 plus table targets, valid inputs only) LibreOffice's ROUND gave "you keep" equal to the target in all 238 (0 short, 0 over). A Decimal simulation of ROUNDUP on the same inputs overshoots the target by a cent in 45 of the 238. Reason: price rounding moves the result by under half a cent and the fee rounding by under half a cent, so the result lands on the exact target cent.
- Edge findings (46 workbooks, 1,380 comparisons, 0 mismatches):
  - Fee 0%: price = target + flat fee ($100.30). Flat fee $0.00: price $102.99. Fee 99%: price $10,030.00 (valid, still OK). Target $0.01: price $0.32, fee $0.31, keep $0.01. Target $0.00: price $0.31, fee $0.31. Target $1,000,000.00: price $1,029,866.43. Rates 3.5% + $0.25 and 2.95% + $0.49 with other targets: all exact. 30 random inputs (random targets up to $8,000.00, percent 0% to 99%, flat fee up to $1.50): all difference $0.00.
  - Fee 100% or 150%, a percent typed as 2.9 (meaning 2.9%, shown as 290%), a negative flat fee, a blank target, or text in the yellow cells: B5 shows a message, B6:B12 and the table are blank, B13 "Check inputs", B14 Check, no #DIV/0! (fee 100% would divide by zero otherwise) and no price near 0.
  - Table: a text or blank target row stays blank and is skipped; a negative target ($-5.00) still computes ($-4.84 price); CHECK is not affected by those rows.
  - The shortcut shortfall grows with the target: $0.09 at $100.00; $841.01 at $1,000,000.00.
- Tested_in: `LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets.`
- Hook paragraph: To end up with a target amount after a percent fee plus a flat fee, divide instead of multiply: `=ROUND((B2+B4)/(1-B3),2)`, where B2 is the amount you want to keep, B3 the fee percent and B4 the flat fee. Multiplying the target by 1.03 comes up short because the fee is charged on the fee too.
- Limits: arithmetic only; nothing on whether a fee may be passed to customers; the ROUND result was checked on a sample, not proven for every input. Not run: Excel, Sheets, a percent typed into a cell not formatted as percent in Excel (the sheet's B5 message covers the wrong value but the Excel entry behavior was not tested).

---

## 3. Three biggest unpaid invoices (`templates/tidy-tabs-top-unpaid-invoices.xlsx`, function `top_unpaid_invoices`)
- Final title: How to Show Your 3 Biggest Unpaid Invoices in Excel and Google Sheets
- Slug: `three-biggest-unpaid-invoices-large-match-excel-google-sheets`
- Sheets: Invoices (header row 4, portrait), Top 3 (landscape), How to use.
- Inputs (yellow): Invoices A5:E104 (invoice #, client, amount, due date, status; 18 filled, $85.00 to $2,400.00), Top 3 A5:A7 (ranks 1, 2, 3).
- Formulas:
  - Invoices F5 `=IF(OR(A5="",E5="Paid",NOT(ISNUMBER(C5))),"",C5)`; G5 `=IF(F5="","",F5-ROW()/1000000)` (minus, so the invoice higher in the list ranks first; the topics file had plus)
  - Top 3 B5 `=IFERROR(LARGE(Invoices!$G$5:$G$104,A5),"")`; C5 `=IF(B5="","",INDEX(Invoices!$A$5:$A$104,MATCH(B5,Invoices!$G$5:$G$104,0)))`; D5 the same with `$B`; E5 the same with `Invoices!$F$5:$F$104`
  - H4 `=COUNT(Invoices!F5:F104)`; H5 `=SUM(Invoices!F5:F104)`; H6 `=SUM(E5:E7)`; H7 `=IF(H5<=0,"",ROUND(H6/H5,4))`
  - H8 CHECK `=IF(AND(ROUND(SUM(Invoices!F5:F104)-SUMIF(Invoices!E5:E104,"<>Paid",Invoices!C5:C104),2)=0,OR(E6="",E5>=E6),OR(E7="",E6>=E7)),"OK","Check")` (the topics file's `E5>=E6` alone returns FALSE when E6 is an empty string, so it needed the OR guard)
- Sample result: 12 open invoices, total open $9,845.75.
  | Rank | Invoice # | Client | Amount |
  | 1 | INV-2045 | Bluebird Bakery | $2,400.00 |
  | 2 | INV-2042 | Pine & Co. Realty | $1,250.00 |
  | 3 | INV-2047 | Oakline Studio | $1,250.00 |
  Top 3 total $4,900.00, share 49.8% (0.4977), H8 OK. Ranks 2 and 3 are a tie; with `LARGE` on the amount alone MATCH would return INV-2042 twice (not run as a copy, it follows from MATCH finding the first hit; say "the key avoids this", do not say it was tested). The amount shown comes from column F, not from the key, and equals the typed amount.
- Edge findings (19 workbooks, 4,123 comparisons, 0 mismatches):
  - Two open invoices: third row blank, share 100%. One open: rows 2 and 3 blank. All paid or no rows: all blank, count 0, total $0.00, share blank, H8 OK.
  - Tie for first (three $500.00 plus a fourth): INV-1, INV-2, INV-4 listed in sheet order, none repeated. Tie across ranks 3 and 4 ($300.00 x 3 for third place): the first of them (INV-3) is shown.
  - Status `paid` and `PAID` are treated as Paid (comparison ignores case); `Paid ` with a trailing space is open. Blank status is open. Amount typed as text (`640`): left out of the list and the totals, and H8 stays OK (the SUMIF ignores text too), so nothing warns; tell readers to check that amounts are right-aligned numbers. Blank amount: left out, OK.
  - A negative amount (credit) is open: with four open invoices (700, 100, -10 and a -250) the third place can be a credit (-$10.00), and the "share" can exceed 100% (146.3%) because the total is reduced by credits. All open amounts negative: ranked, share blank. Cents (0.10, 0.20, 0.30) rank correctly.
  - An invoice with an amount but no invoice number: left out (F blank) and H8 shows Check because the SUMIF sees it. Blank row in the middle: skipped. 100 random rows (with repeated 640 and 1250 amounts): all matched.
- Tested_in: `LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets.`
- Hook paragraph: Find the largest open amount with `=LARGE(Invoices!$G$5:$G$104,1)`, change the 1 to 2 and 3 for the next two, then bring back the invoice number with `=INDEX(Invoices!$A$5:$A$104,MATCH(B5,Invoices!$G$5:$G$104,0))`. Column G is the open amount minus a tiny row-based number, so two invoices for the same amount are not mixed up.
- Limits: the key adds 0.000001 per row, so amounts that differ by under a cent in 100 rows could mis-order (not run); amounts above about 1e9 not run. Not run: Excel, Sheets.

---

## 4. Number from text like 12 oz (`templates/tidy-tabs-number-from-text-units.xlsx`, function `number_from_text_units`)
- Final title: How to Pull the Number Out of Text Like 12 oz or 3 lb in Excel and Google Sheets
- Slug: `pull-number-out-of-text-12-oz-value-left-find-excel-google-sheets`
- Sheets: Weights (header row 4), How to use.
- Inputs (yellow): A5:B24 (item, typed weight as text; 12 filled, column B formatted as text). The sample has `10 lb`, `12 oz`, `3oz`, `16 OZ`, ` 3 lb ` (outer spaces), `1.5 lb`, `.5 lb`, `12 oz.`, `4  oz` (two spaces), `2lb`, `6 oz`, `1 lb`.
- Formulas (row 5):
  - C5 `=IF(A5="","",IF(AND(ISNUMBER(SEARCH("oz",B5))<>ISNUMBER(SEARCH("lb",B5)),ISERROR(SEARCH("fl",B5)),ISERROR(SEARCH("/",B5)),ISERROR(SEARCH("-",B5)),ISERROR(SEARCH(":",B5))),IFERROR(VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1)),IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""),"lb","")),"Check")),"Check"))`
  - Core space form `=VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1))`; core no-space form `=VALUE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""))` (each run on its own: the first gives 12 for `12 oz`, 3 for ` 3 lb `, 16 for `16 OZ` and #VALUE! for `12oz`; the second gives 12 for `12oz` and `16 OZ` and #VALUE! for ` 3 lb `)
  - D5 `=IF(OR(A5="",NOT(ISNUMBER(C5))),"",IF(ISNUMBER(SEARCH("lb",B5)),"lb","oz"))`; E5 `=IF(ISNUMBER(C5),IF(D5="lb",ROUND(C5*16,2),C5),"")`
  - H4 `=SUM(E5:E24)`; H5 `=INT(ROUND(H4,2)/16)&" lb "&ROUND(MOD(ROUND(H4,2),16),2)&" oz"`; H6 `=COUNTIF(C5:C24,"Check")`; H7 CHECK `=IF(AND(COUNTIF(C5:C24,"Check")=0,COUNT(E5:E24)=COUNTA(A5:A24),COUNTA(B5:B24)=COUNTA(A5:A24)),"OK","Check")`
- Sample result: numbers 10, 12, 3, 16, 3, 1.5, 0.5, 12, 4, 2, 6, 1; ounces 160, 12, 3, 16, 48, 24, 8, 12, 4, 32, 6, 16; total 341 oz = `21 lb 5 oz`; 0 Check rows; H7 OK.
- Edge findings (8 workbooks, 512 comparisons, 0 mismatches against a Python mirror; rows marked "observed" are recorded, not compared):
  - Read correctly: `12 oz`, `12oz`, `12  oz`, `  3 lb `, `16 OZ`, `12 Oz`, `1.5 lb`, `.5 lb`, `12 oz.`, `12oz.`, `oz 12`, `oz12`, `lb 2`, `2 lb.`, `12lb`, `3LB`, `12 lbs`, `+5 oz`, `0 oz`, `100 oz`, `8 oz  `, `10  lb`; also `12 oz 3` (reads 12) and `12 oz oz` (reads 12).
  - Check (by design, no guess): `n/a`, a blank weight, a plain number cell (`12`, `12.5`: no unit), `oz` alone, `12 fl oz` and `fl oz` (fl guard), `1 lb 4 oz` (both units), `12 ounces`, `5 kg`, `twelve oz`, `12.5.5 oz`, `-3 lb`, `1/2 lb`, `12-oz` (the slash, hyphen and colon guard; before the guard LibreOffice returned 46024 and -12), two spaces only.
  - Observed in LibreOffice with US settings: `1,5 lb` showed Check; `1,200 oz` was read as 1200; `1e3 oz` as 1000; `12` + non-breaking space + `oz` was read as 12 (no CHAR(160) step needed in LibreOffice). Excel and Sheets may differ; do not generalize.
  - Item without a weight, or a weight without an item (row ignored by C:E): H6 counts the Check row, H7 Check; a weight with no item name makes the COUNTA test fail, H7 Check.
  - Totals: 0.333 lb x 3 plus 0.999 oz = 16.989 oz shown `1 lb 0.99 oz`; 100 lb + 15.99 oz + 0.01 oz = 1616 oz = `101 lb 0 oz` (H4 is rounded to 2 decimals before the split so 15.999 cannot show as `16 oz`). 20 full rows: matched.
- Tested_in: `LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets.`
- Hook paragraph: For text like `12 oz` with a space, take everything before the space with `=VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1))`. For text like `12oz` with no space, strip the unit instead: `=VALUE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""))`. Either way you get a real number you can add up.
- Limits: units other than oz and lb show Check; a unit word that contains oz or lb (for example `lbs`, `ozs`) is read as that unit. Decimal commas were not tested beyond the observations above. Not run: Excel, Sheets, a non-US language setting.

---

## 5. Payment reminder from invoice data (`templates/tidy-tabs-payment-reminders.xlsx`, function `payment_reminders`)
- Final title: How to Write a Friendly Payment Reminder From Your Invoice List in Excel and Google Sheets
- Slug: `payment-reminder-message-from-invoice-data-text-excel-google-sheets`
- Sheets: Invoices (landscape, header row 4), Settings, How to use.
- Inputs (yellow): Settings B1 as-of date `11/02/2026` (typed, no TODAY), B2 `Jordan`; Invoices A5:E64 (invoice #, first name, amount, due date, Paid; 12 filled, 3 paid).
- Formulas (row 5):
  - Settings B3 `=IF(TRIM(B2)="","","Thanks, "&TRIM(B2))`; C1 `=IF(ISNUMBER(B1),"OK","Type a date like 11/02/2026")`
  - F5 `=IF(OR(A5="",TRIM(E5)="Yes"),"",IF(AND(ISNUMBER(D5),ISNUMBER(Settings!$B$1)),INT(Settings!$B$1)-INT(D5),"Check date"))`
  - G5 `=IF(F5="","",IF(OR(NOT(ISNUMBER(F5)),NOT(ISNUMBER(C5))),"Check",IF(F5<=0,"Upcoming",IF(F5<=14,"Gentle","Follow-up"))))`
  - H5 `=IF(G5="","",IF(G5="Check","Check the amount and due date in this row",TRIM(IF(F5<=0,"Hi "&IF(TRIM(B5)="","there",TRIM(B5))&", "&"just a heads-up that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" is due on "&TEXT(D5,"mm/dd/yyyy")&". ","Hi "&IF(TRIM(B5)="","there",TRIM(B5))&", "&"a quick reminder that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" was due on "&TEXT(D5,"mm/dd/yyyy")&" ("&IF(F5=1,"1 day",F5&" days")&" ago). Let me know if you need another copy. ")&Settings!$B$3)))`
  - Basic form quoted in the hook (run on its own for Maya, INV-2051, $1,250.50, 10/05/2026): `="Hi "&B5&", a quick reminder that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" was due on "&TEXT(D5,"mm/dd/yyyy")&"."` gave `Hi Maya, a quick reminder that invoice INV-2051 for $1,250.50 was due on 10/05/2026.`
  - K4 `=COUNT(F5:F64)-COUNTIF(F5:F64,"<=0")`; K5 `=SUMIF(F5:F64,">0",C5:C64)`; K6 CHECK `=IF(AND(ISNUMBER(Settings!B1),COUNTIF(G5:G64,"Check")=0,SUMPRODUCT(--ISNUMBER(F5:F64),--(H5:H64=""))=0,SUMPRODUCT(--ISERROR(H5:H64))=0),"OK","Check")`
- Sample result (as-of 11/02/2026; the rows with messages):
  | Invoice | Days past due | Tone | Message |
  | INV-2052 | 45 | Follow-up | Hi Devon, a quick reminder that invoice INV-2052 for $85.00 was due on 09/18/2026 (45 days ago). Let me know if you need another copy. Thanks, Jordan |
  | INV-2053 | 3 | Gentle | Hi Priya, a quick reminder that invoice INV-2053 for $640.00 was due on 10/30/2026 (3 days ago). Let me know if you need another copy. Thanks, Jordan |
  | INV-2054 | 14 | Gentle | Hi Sam, ... $320.00 ... 10/19/2026 (14 days ago). ... |
  | INV-2055 | 0 | Upcoming | Hi Alicia, just a heads-up that invoice INV-2055 for $1,200.00 is due on 11/02/2026. Thanks, Jordan |
  | INV-2059 | 16 | Follow-up | Hi Nora, ... $150.00 ... 10/17/2026 (16 days ago). ... |
  Also INV-2056 (-8, Upcoming, $475.25), INV-2057 (-18), INV-2060 (5 days, $960.00), INV-2062 (-3). Paid rows (INV-2051, INV-2058, INV-2061) have no days, tone or message. K4 5 messages to send, K5 $2,155.00, K6 OK. TEXT with `"$#,##0.00"` and `"mm/dd/yyyy"` returned the US format in LibreOffice 24.2.7.2 (en-US).
- Edge findings (14 workbooks, 2,590 comparisons of exact message strings and numbers, 0 mismatches):
  - Days: due = as-of gives 0 (Upcoming, "is due on"); due one day after gives -1 (Upcoming); 1 day past is "(1 day ago)" Gentle; 13, 14 past are Gentle; 15, 16 are Follow-up. A due date with a time part (10/30/2026 3:30 PM) gives 3 days (INT on both). Leap day: due 02/29/2028 with as-of 03/01/2028 gives 1 day, due 02/28/2028 gives 2. Year end: due 12/31/2025, as-of 01/02/2026 formats as 12/31/2025.
  - Amounts: $1,250.50, $85.00, $0.00, $0.50, $1,000.00 and $1,234,567.89 formatted correctly; 19.995 showed $20.00. Amount typed as text (`100`) gives Check and the message "Check the amount and due date in this row"; blank amount: same.
  - Client name blank or only spaces becomes `there`; `  Anna  Lee ` becomes `Anna Lee`. Due date typed as text (`10/01/2026`) or blank: F "Check date", G Check, K6 Check. As-of date typed as text or blank: every unpaid row shows Check date, K6 Check, Settings C1 message.
  - Paid typed `yes` or `Yes ` counts as paid (TRIM and case-insensitive); `Y` or `No` are not paid. A row with a blank invoice # is ignored even if it has a name and amount.
  - Settings name blank: messages end without a sign-off (`... if you need another copy.`). Name `  Jordan   Lee ` becomes `Thanks, Jordan Lee`.
  - Empty list and all paid: K4 0, K5 $0.00, K6 OK. 60 random rows: matched.
- Tested_in: `LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets.` TODAY() was not used and not run.
- Hook paragraph: Build the message from the cells with `="Hi "&B5&", a quick reminder that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" was due on "&TEXT(D5,"mm/dd/yyyy")&"."`. Copy it down and each row gives you a sentence you can paste into an email, with the amount and date formatted the US way.
- Limits: TEXT format codes follow the app's language setting (not run in another language); the post must say so and mention Format > Number. The wording has no fee, deadline or legal text and the post gives no collections advice. Not run: Excel, Sheets, non-US settings.
