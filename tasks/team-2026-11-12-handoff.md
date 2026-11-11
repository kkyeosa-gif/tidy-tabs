# Handoff 2026-11-12: five topics researched, hooks written, templates built and tested, heroes made (LibreOffice Calc 24.2.7.2 only), for the script-writers

Stage 1 of the content team (researcher, hook-writer, builder, designer) was done by one agent. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: tasks/team-topics-2026-11-12.md, tasks/hooks-2026-11-12.md, scripts/build-templates.py (`rank_within_category`, `apply_payment_oldest_first`, `rental_weekend_rate`, `time_numbers_24_hour`, `upc_check_digit`), scripts/make-heroes.py (`heroes_2026_11_12`, run with `--only-2026-11-12`), five new files in templates/, five heroes in images/, results in tasks/briefs.md ("2026-11-12 topics (researcher)" and "2026-11-12 builder results"), five titles appended to tasks/topics-seed.md.

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 83 edited workbooks compared (64,213 comparisons, 0 mismatches against independent Python models), plus 8 probe workbooks read by hand for the short formulas and the failure cases below; no error value in any recalculated workbook. All five files exported to PDF: every page US Letter, no error values or ### in the PDF text. Bugs found and fixed are in tasks/briefs.md (an open amount with fractions of a cent gave Balance after -$0.01 and Partly paid in the payment file; such amounts now show Check).

Repo note: only the five new builder functions were run, so no older template was rewritten; `git status` shows the two scripts, the topics, hooks, handoff, briefs and seed files, five new .xlsx files and five image folders. No older file and no __pycache__ changed.

Post file names: `posts/ready/2026-11-12-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas; a red fill marks Check (and the other flagged results named below), a green fill marks rank 1, Paid in full and OK UPCs. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays. The payment post must not say how a payment must legally or contractually be applied (oldest first is one method; the client's remittance decides). The rental post states the billing rule as a setting, not a rule anyone must follow. The UPC post must not say a code is registered, valid for retail or owned by anyone (arithmetic only; GS1 US assigns company prefixes, link https://www.gs1us.org).

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-12-team-01-rank-products-within-each-category-countifs-excel-google-sheets | tidy-tabs-rank-within-category.xlsx | 21 workbooks, 16,905 comparisons |
| 2 | 2026-11-12-team-02-apply-payment-to-oldest-invoices-first-sumifs-min-excel-google-sheets | tidy-tabs-apply-payment-oldest-first.xlsx | 30 workbooks, 18,300 comparisons |
| 3 | 2026-11-12-team-03-rental-price-weekday-weekend-day-rates-networkdays-excel-google-sheets | tidy-tabs-rental-weekend-rate.xlsx | 16 workbooks, 16,112 comparisons |
| 4 | 2026-11-12-team-04-convert-24-hour-time-numbers-like-1430-to-am-pm-time-int-mod-excel-google-sheets | tidy-tabs-24-hour-time-numbers.xlsx | 8 workbooks, 4,840 comparisons |
| 5 | 2026-11-12-team-05-check-upc-barcode-check-digit-mid-sumproduct-mod-excel-google-sheets | tidy-tabs-upc-check-digit.xlsx | 8 workbooks, 8,056 comparisons |

Hero files (1600x900 PNG, real LibreOffice PDF render of the template with its sample data; alt text must be written per post, 8 to 15 words, not the title):
1. images/2026-11-12-team-01-rank-products-within-each-category-countifs-excel-google-sheets/rank-within-category-template.png (caption: "Fourteen products in four categories with the rank in each category, the overall rank and the Top 3 mark. PDF export from LibreOffice Calc 24.2.")
2. images/2026-11-12-team-02-apply-payment-to-oldest-invoices-first-sumifs-min-excel-google-sheets/apply-payment-oldest-invoices-first-template.png (caption: "A $2,400.00 payment applied by invoice date: two invoices paid in full, one partly paid, four unpaid. PDF export from LibreOffice Calc 24.2.")
3. images/2026-11-12-team-03-rental-price-weekday-weekend-day-rates-networkdays-excel-google-sheets/rental-weekday-weekend-rate-template.png (caption: "Seven rentals priced at $45.00 a weekday and $65.00 a weekend or holiday day. PDF export from LibreOffice Calc 24.2.")
4. images/2026-11-12-team-04-convert-24-hour-time-numbers-like-1430-to-am-pm-time-int-mod-excel-google-sheets/24-hour-time-to-am-pm-template.png (caption: "Exported times such as 1430 and 0745 turned into real times; 1460, 2400 and 14:30 are flagged. PDF export from LibreOffice Calc 24.2.")
5. images/2026-11-12-team-05-check-upc-barcode-check-digit-mid-sumproduct-mod-excel-google-sheets/upc-check-digit-template.png (caption: "Twelve fictional product codes with the computed check digit, the result and a note. PDF export from LibreOffice Calc 24.2.")
The summary block is in the hero for posts 1, 2 and 4; it is hidden in the hero (not in the file) for posts 3 and 5.

AI photo (`image_prompts`, `image_alt`) is still open for every post: one scene each, 2 or 3 specific objects, no screens as the subject, no readable text, no logos, no laptop and coffee:
1. Rank: a craft-fair shelf with candle jars grouped by scent in three rows and soap bars stacked beside them, one small blank ribbon rosette pinned next to the front jar of each group.
2. Payment: a single paper check lying on a short stack of printed invoices in a wire desk tray, the bottom two invoices pulled out and clipped together with a binder clip.
3. Rental: a folded white canopy tent in its carry bag leaning on a stack of folding chairs by an open garage door, with a paper wall calendar whose weekend columns are shaded in pencil (dates unreadable).
4. Times: a long curl of thermal receipt tape on a farmers market booth table next to a gray metal cash box and a small hand-held wind-up kitchen timer (print unreadable).
5. UPC: a roll of blank barcode sticker labels beside a stack of kraft product boxes and a hand-held barcode scanner lying on a packing table (any bars blurred, no digits readable).

Every workbook has a "How to use" tab. Titles and openings: tasks/hooks-2026-11-12.md. Deviations from tasks/team-topics-2026-11-12.md: none in topic or slug; post 4's title was changed to "...to AM and PM Times..." so it has no colon (the 2:30 PM in the earlier draft had one).

---
## 1. rank-products-within-each-category-countifs-excel-google-sheets
Title: How to Rank Products Within Each Category in Excel and Google Sheets
Template: templates/tidy-tabs-rank-within-category.xlsx (tabs: Sales, How to use)

Layout
- Sales: A1 "Mark the top", B1 how many to mark (input, 3), C1 note. Header row 3, data rows 4:203 (14 sample products, not sorted). A Product, B Category, C Sales (inputs, C as $); D Rank in category, E Products in category, F Overall rank, G Top N mark (formulas). Summary I3:J7: Rows ranked (J3), Rows showing Check (J4), Rows ranked 1 (ties count) (J5), Rows marked Top N (J6), CHECK (J7). D is green when 1 and red when Check.

Key formulas (row 4)
- D4 `=IF(AND(TRIM(A4)="",TRIM(B4)="",C4=""),"",IF(OR(TRIM(B4)="",NOT(ISNUMBER(C4)),SUMPRODUCT(--ISNUMBER(FIND({"*","?","~"},B4)))>0),"Check",COUNTIFS($B$4:$B$203,B4,$C$4:$C$203,">"&C4)+1))` (the post teaches the core `=COUNTIFS($B$4:$B$203,B4,$C$4:$C$203,">"&C4)+1`)
- E4 `=IF(ISNUMBER(D4),COUNTIFS($B$4:$B$203,B4,$D$4:$D$203,">0"),"")`
- F4 `=IF(ISNUMBER(D4),COUNTIFS($C$4:$C$203,">"&C4,$D$4:$D$203,">0")+1,"")` (overall rank among the ranked rows, same tie rule)
- G4 `=IF(ISNUMBER(D4),IF(ISNUMBER($B$1),IF(AND($B$1>=1,$B$1=INT($B$1)),IF(D4<=$B$1,"Top "&$B$1,""),"Check"),"Check"),"")`
- J7 `=IF(AND(J4=0,COUNTIF(G4:G203,"Check")=0),"OK","Check")`

Sample values (fictional; B1 = 3; recalculated)
| Product | Category | Sales | Rank in category | Products in category | Overall rank | Top N mark |
|---|---|---|---|---|---|---|
| Lavender soy candle 8 oz | Candles | $1,240.00 | 2 | 4 | 3 | Top 3 |
| Oatmeal goat milk soap | Soap | $720.00 | 2 | 4 | 8 | Top 3 |
| Brass hoop earrings | Jewelry | $1,890.00 | 1 | 3 | 1 | Top 3 |
| Letterpress birthday card | Cards | $310.00 | 3 | 3 | 14 | Top 3 |
| Cedar soy candle 8 oz | Candles | $980.00 | 3 | 4 | 6 | Top 3 |
| Charcoal soap bar | Soap | $720.00 | 2 | 4 | 8 | Top 3 |
| Pearl drop earrings | Jewelry | $1,125.00 | 2 | 3 | 4 | Top 3 |
| Thank you card set | Cards | $492.00 | 2 | 3 | 12 | Top 3 |
| Vanilla bean candle 12 oz | Candles | $1,410.50 | 1 | 4 | 2 | Top 3 |
| Lemon poppy seed soap | Soap | $455.25 | 4 | 4 | 13 | |
| Beaded stack bracelet | Jewelry | $560.00 | 3 | 3 | 11 | Top 3 |
| Holiday card 6 pack | Cards | $1,005.00 | 1 | 3 | 5 | Top 3 |
| Pumpkin spice candle tin | Candles | $615.00 | 4 | 4 | 10 | |
| Lavender soap bar | Soap | $802.00 | 1 | 4 | 7 | Top 3 |
Summary J3:J7: 14, 0, 4, 12, OK. The two soaps at $720.00 are both 2 and the next soap is 4 (no 3 in Soap); both are 8 overall and the next overall is 10.

Pitfalls the post must state
- Ties share a rank and the next rank is skipped (2, 2, 4), like RANK. There is no tie-breaker in the file.
- The list does not need sorting; COUNTIFS looks at the whole range.
- Category matching ignores upper and lower case (Candles, candles, CANDLES are one group) but not spaces: "Candles " with a trailing space is its own group.
- * ? ~ inside a category would act as wildcards in COUNTIFS, so the file shows Check for those rows.
- Sales typed as text ("$900.00" pasted as text) show Check; use the existing convert-text-to-numbers post to fix them.
- Not tested: Excel, Google Sheets. Not tested: decimals with more than 15 significant digits.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Every product in a category shows 1 or the ranks look too low: the category has a trailing space on some rows. Probe: "Candles " with $2,000.00 got rank 1 with Products in category 1, separate from the Candles rows.
- A row shows Check: the sales cell held the text "$900.00" (Check), or the category was blank, or contained a wildcard ("Cand*" showed Check).
- The core formula with a wildcard category counts the wrong rows: in a probe, `=COUNTIFS($B$4:$B$203,B23,$C$4:$C$203,">"&C23)+1` with B23 = "Cand*" and C23 = 10 returned 7 (it counted 6 Candles rows, including "candles" and "Candles ").
- Two equal amounts get the same rank and a number is skipped (sample, Soap: 1, 2, 2, 4).
- Top N mark shows Check on every row: B1 is 0, negative, 2.5, blank or text (tested 0, -1, 2.5, "abc", "3" typed as text and blank).

---
## 2. apply-payment-to-oldest-invoices-first-sumifs-min-excel-google-sheets
Title: How to Apply a Client Payment to the Oldest Invoices First in Excel and Google Sheets
Template: templates/tidy-tabs-apply-payment-oldest-first.xlsx (tabs: Invoices, How to use)

Layout
- Invoices: A1 "Payment received", B1 payment (input, $2,400.00); A2 "Client", B2 client name (input, "Harbor Dental (fictional)"); C1 note. Header row 4, data rows 5:204 (7 sample invoices; the oldest, INV-1039, is typed last). A Invoice no., B Invoice date (MM/DD/YYYY), C Open amount (inputs); D Applied from payment, E Balance after, F Status (formulas). Summary H4:I13: Payment received (I4), Open before the payment (I5), Applied to invoices (I6), Left over (not applied) (I7), Still open after (I8), Paid in full (I9), Partly paid (I10), Unpaid (I11), Rows showing Check (I12), CHECK (I13). F is green for Paid in full, red for Check.

Key formulas (row 5)
- D5 `=IF(AND(TRIM(A5)="",B5="",C5=""),"",IF(OR(NOT(ISNUMBER(B5)),NOT(ISNUMBER(C5))),"Check",IF(OR(C5<0,ROUND(C5,2)<>C5,NOT(ISNUMBER($B$1))),"Check",IF(OR($B$1<0,ROUND($B$1,2)<>$B$1),"Check",ROUND(MAX(0,MIN(C5,$B$1-(SUMIFS($C$5:$C$204,$B$5:$B$204,"<"&B5,$C$5:$C$204,">=0")+SUMIFS($C$4:C4,$B$4:B4,B5,$C$4:C4,">=0")))),2)))))`
- The post teaches the core `=ROUND(MAX(0,MIN(C5,$B$1-(SUMIFS($C$5:$C$204,$B$5:$B$204,"<"&B5,$C$5:$C$204,">=0")+SUMIFS($C$4:C4,$B$4:B4,B5,$C$4:C4,">=0")))),2)`: first SUMIFS = open amounts with an earlier date; second SUMIFS (range grows down, starts at the header row) = same-date invoices on the rows above, so same-day invoices are paid top to bottom. MIN caps at the invoice amount, MAX stops it going below 0.
- E5 `=IF(ISNUMBER(D5),ROUND(C5-D5,2),"")`
- F5 `=IF(D5="","",IF(D5="Check","Check",IF(C5=0,"Nothing due",IF(E5=0,"Paid in full",IF(D5>0,"Partly paid","Unpaid")))))`
- I5 `=SUMIFS(C5:C204,D5:D204,">=0")`; I6 `=SUM(D5:D204)`; I7 `=IF(ISNUMBER(B1),ROUND(B1-I6,2),"")`; I8 `=ROUND(I5-I6,2)`
- I13 `=IF(AND(ISNUMBER(B1),I12=0),IF(AND(B1>=0,ROUND(B1,2)=B1,ROUND(I6+I7-B1,2)=0,I7>=0,OR(I7=0,I8=0)),"OK","Check"),"Check")`

Sample values (fictional; payment $2,400.00; recalculated)
| Invoice no. | Invoice date | Open amount | Applied | Balance after | Status |
|---|---|---|---|---|---|
| INV-1041 | 08/15/2026 | $1,250.00 | $1,250.00 | $0.00 | Paid in full |
| INV-1047 | 09/01/2026 | $840.00 | $775.00 | $65.00 | Partly paid |
| INV-1052 | 09/15/2026 | $1,250.00 | $0.00 | $1,250.00 | Unpaid |
| INV-1058 | 10/01/2026 | $612.50 | $0.00 | $612.50 | Unpaid |
| INV-1063 | 10/15/2026 | $1,250.00 | $0.00 | $1,250.00 | Unpaid |
| INV-1064 | 10/15/2026 | $180.00 | $0.00 | $180.00 | Unpaid |
| INV-1039 | 08/01/2026 | $375.00 | $375.00 | $0.00 | Paid in full |
Summary I4:I13: $2,400.00, $5,757.50, $2,400.00, $0.00, $3,357.50, 2, 1, 4, 0, OK. Order of payment: INV-1039 $375.00, INV-1041 $1,250.00, INV-1047 $775.00 (375 + 1,250 + 775 = 2,400).
With a payment of $7,000.00 (probe): all 7 Paid in full, Left over $1,242.50, Still open $0.00, CHECK OK.

Pitfalls the post must state
- The order comes from the invoice date, not the row; the list does not need sorting. Two invoices on the same date are paid in row order (INV-1063 before INV-1064).
- The shorter "running total" version `=MIN(C5,MAX(0,$B$1-SUM($C$4:C4)))` only works when the rows are sorted oldest first (see Troubleshooting).
- Left over means the payment is more than everything open; the file does not decide what to do with it (refund, credit, next invoice).
- Dates must be real dates; amounts must be numbers of 0 or more in whole cents, and so must the payment. Fix every Check row first: until then the other rows can be off, because a Check row with a date and an amount of 0 or more can still be counted as older.
- Credits and partial refunds are not handled; enter the open amount after them.
- Oldest first is one way to apply a payment. If the client says which invoices a payment covers, follow the remittance. No accounting or legal advice.
- Not tested: Excel, Google Sheets (the SUMIFS with "<"&B5 on dates and the expanding $B$4:B4 range are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- The oldest invoice gets nothing although it is the oldest: the sorted-only formula on the unsorted sample (probe) paid INV-1041 $1,250.00, INV-1047 $840.00, INV-1052 $310.00 and INV-1039 $0.00. The file's SUMIFS version pays INV-1039 first.
- A row shows Check: the date was typed as text ("10/20/2026" entered as a text string), the amount was text or negative, or the amount had fractions of a cent (99.999 and 12.345 showed Check).
- CHECK shows Check although every row looks right: B1 is blank, text, negative or has fractions of a cent (tested blank, "abc", "2400" as text, -1, 2400.005).
- Left over is more than $0.00: the payment is bigger than all open invoices ($7,000.00 left $1,242.50).

---
## 3. rental-price-weekday-weekend-day-rates-networkdays-excel-google-sheets
Title: How to Price a Rental With Weekday and Weekend Day Rates in Excel and Google Sheets
Template: templates/tidy-tabs-rental-weekend-rate.xlsx (tabs: Rentals, Holidays, How to use)

Layout
- Rentals: A1 "Weekday day rate", B1 (input, $45.00); A2 "Weekend or holiday day rate", B2 (input, $65.00); C1 note. Header row 4, data rows 5:204 (7 sample event-equipment rentals). A Customer, B Item, C Pickup date, D Return date (inputs, MM/DD/YYYY); E Days billed, F Weekend and holiday days, G Weekday days, H Price, I Status (formulas). Summary K4:L10: Rentals listed (L4), Days billed (L5), Weekend and holiday days (L6), Weekday days (L7), Total price (L8), Rows showing Check (L9), CHECK (L10). I is red for Check.
- Holidays: A1 note, header row 3, A4:A23 dates (inputs) and B4:B23 names. Sample: 11/26/2026 Thanksgiving Day, 12/25/2026 Christmas Day, 01/01/2027 New Year's Day.

Key formulas (row 5)
- E5 `=IF(AND(TRIM(A5)="",TRIM(B5)="",C5="",D5=""),"",IF(OR(NOT(ISNUMBER(C5)),NOT(ISNUMBER(D5))),"Check",IF(OR(C5<>INT(C5),D5<>INT(D5),D5<C5),"Check",MAX(1,D5-C5))))` (core `=MAX(1,D5-C5)`)
- F5 `=IF(ISNUMBER(E5),E5-NETWORKDAYS(C5,C5+E5-1,Holidays!$A$4:$A$23),"")` (core `=E5-NETWORKDAYS(C5,C5+E5-1,Holidays!$A$4:$A$23)`; C5+E5-1 is the last billed day)
- G5 `=IF(ISNUMBER(E5),E5-F5,"")`
- H5 `=IF(ISNUMBER(E5),IF(AND(ISNUMBER($B$1),ISNUMBER($B$2)),IF(AND($B$1>=0,$B$2>=0),ROUND(G5*$B$1+F5*$B$2,2),"Check"),"Check"),"")` (core `=ROUND(G5*$B$1+F5*$B$2,2)`)
- I5 `=IF(E5="","",IF(OR(E5="Check",H5="Check"),"Check",IF(D5=C5,"Same-day return","OK")))`
- L10 `=IF(AND(L9=0,L5=L6+L7,COUNT(B1:B2)=2),"OK","Check")`

Sample values (fictional; $45.00 weekday, $65.00 weekend or holiday; recalculated)
| Customer | Item | Pickup | Return | Days billed | Weekend and holiday | Weekday | Price | Status |
|---|---|---|---|---|---|---|---|---|
| Lopez family | 10x10 canopy tent | 11/06/2026 (Fri) | 11/09/2026 (Mon) | 3 | 2 | 1 | $175.00 | OK |
| Harbor Dental | Folding tables, set of 6 | 11/10/2026 (Tue) | 11/12/2026 (Thu) | 2 | 0 | 2 | $90.00 | OK |
| Jordan Reed | Photo booth backdrop | 11/14/2026 (Sat) | 11/14/2026 (Sat) | 1 | 1 | 0 | $65.00 | Same-day return |
| Cedar Print Shop | String lights, 100 ft | 11/16/2026 (Mon) | 11/23/2026 (Mon) | 7 | 2 | 5 | $355.00 | OK |
| Lakeside Yoga | Patio heater | 11/25/2026 (Wed) | 11/30/2026 (Mon) | 5 | 3 (Thanksgiving, Sat, Sun) | 2 | $285.00 | OK |
| Maria Lopez | Folding chairs, set of 20 | 12/05/2026 (Sat) | 12/08/2026 (Tue) | 3 | 2 | 1 | $175.00 | OK |
| Pine & Co. Realty | 10x10 canopy tent | 12/24/2026 (Thu) | 12/28/2026 (Mon) | 4 | 3 (Christmas Fri, Sat, Sun) | 1 | $240.00 | OK |
Summary L4:L10: 7, 25, 13, 12, $1,385.00, 0, OK. (The weekday names in parentheses are for the post; the file shows dates only.)

Pitfalls the post must state
- Billing rule in the file: the pickup day is billed, the return day is not; a same-day return is 1 day. Friday to Monday = Friday, Saturday, Sunday = 3 days. To bill the return day too, use `=D5-C5+1` (probe: 4 for the Friday to Monday rental) and change the NETWORKDAYS end to match.
- NETWORKDAYS counts Monday to Friday minus listed holidays; the weekend is Saturday and Sunday only. A holiday on a Saturday or Sunday is not counted twice (tested 07/04/2026 and 11/07/2026, both Saturdays).
- Holidays must be real dates. In LibreOffice a holiday typed as text ("11/27/2026" as text) and a word ("Black Friday") were both ignored with no error, so the day was billed at the weekday rate. Excel and Sheets behavior with text in the holiday range was not tested; say so.
- Dates with a time in them, text dates and a return before pickup show Check.
- The rates and the holiday list are the reader's own settings; the file decides nothing about what to charge.
- Not tested: Excel, Google Sheets. NETWORKDAYS.INTL (other weekends) is not used; link Microsoft's NETWORKDAYS.INTL help if mentioned.

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- Status shows Check: return date before the pickup (11/20/2026 to 11/18/2026), pickup typed as text ("11/20/2026" as a text string), pickup with a time (11/20/2026 9:30 AM), or no return date.
- A holiday is billed at the weekday rate: the holiday was typed as text in the Holidays list (probe: "11/27/2026" as text was ignored; the Lakeside Yoga rental stayed at 3 weekend and holiday days).
- One day too few or too many: the return day is not billed in this file (Friday to Monday is 3); `=D5-C5+1` gives 4 if you bill the return day.
- Price shows Check on every row: a rate is blank, text or negative (tested blank, "abc", -1).

---
## 4. convert-24-hour-time-numbers-like-1430-to-am-pm-time-int-mod-excel-google-sheets
Title: How to Convert 24-Hour Times Like 1430 to AM and PM Times in Excel and Google Sheets
Template: templates/tidy-tabs-24-hour-time-numbers.xlsx (tabs: Times, How to use)

Layout
- Times: A1 note. Header row 3, data rows 4:203 (13 sample sales tickets). A Ticket, B Time as exported (inputs; numbers or text); C Time (formula, real time formatted h:mm AM/PM), D As text (12-hour) (formula), E Status (OK or Check, red for Check). Summary G3:H7: Times listed (H3), Converted (H4), Rows showing Check (H5), Earliest time (H6), Latest time (H7). Portrait page (the other four files are landscape).

Key formulas (row 4; `p` = RIGHT("000"&TRIM(B4),4))
- C4 `=IF(TRIM(B4)="","",IF(LEN(TRIM(B4))>4,"Check",IF(SUMPRODUCT(--ISNUMBER(FIND(MID(RIGHT("000"&TRIM(B4),4),{1,2,3,4},1),"0123456789")))<4,"Check",IF(OR(--LEFT(RIGHT("000"&TRIM(B4),4),2)>23,--RIGHT(RIGHT("000"&TRIM(B4),4),2)>59),"Check",TIME(--LEFT(RIGHT("000"&TRIM(B4),4),2),--RIGHT(RIGHT("000"&TRIM(B4),4),2),0)))))`
- The post teaches the core `=TIME(INT(B4/100),MOD(B4,100),0)` (INT(B4/100) = hours, MOD(B4,100) = minutes) plus the cell format h:mm AM/PM, then explains why the file pads and checks.
- D4 `=IF(ISNUMBER(C4),TEXT(C4,"h:mm AM/PM"),"")`; E4 `=IF(C4="","",IF(ISNUMBER(C4),"OK","Check"))`
- H6 `=IF(H4=0,"",MIN(C4:C203))`; H7 `=IF(H4=0,"",MAX(C4:C203))`

Sample values (fictional tickets; recalculated)
| Ticket | Time as exported | Time | As text | Status |
|---|---|---|---|---|
| T-101 | 0745 (text) | 7:45 AM | 7:45 AM | OK |
| T-102 | 905 | 9:05 AM | 9:05 AM | OK |
| T-103 | 1200 | 12:00 PM | 12:00 PM | OK |
| T-104 | 1430 | 2:30 PM | 2:30 PM | OK |
| T-105 | 0 | 12:00 AM | 12:00 AM | OK |
| T-106 | 2359 | 11:59 PM | 11:59 PM | OK |
| T-107 | 1205 | 12:05 PM | 12:05 PM | OK |
| T-108 | " 815 " (text with spaces) | 8:15 AM | 8:15 AM | OK |
| T-109 | 30 | 12:30 AM | 12:30 AM | OK |
| T-110 | 1745 | 5:45 PM | 5:45 PM | OK |
| T-111 | 1460 | Check | | Check |
| T-112 | 2400 | Check | | Check |
| T-113 | 14:30 (text) | Check | | Check |
Summary H3:H7: 13, 10, 3, 12:00 AM, 11:59 PM.

Pitfalls the post must state
- TIME returns a fraction of a day (the recalculated value of 2:30 PM is 0.6041666...). The file formats column C as h:mm AM/PM. In the LibreOffice probe a new cell with the core formula was shown as a time on its own; if a cell shows a decimal instead, format it as h:mm AM/PM (not tested in Excel or Sheets).
- The core formula does not check anything: 1460 becomes 3:00 PM (TIME carries 60 minutes into the next hour) and 2400 becomes 12:00 AM (TIME wraps 24 hours to 0). The file shows Check for both. If an export uses 2400 for end of day, type 0 or decide yourself.
- Numbers lose leading zeros (0745 becomes 745) but that does not change the time: 745 and 0745 both give 7:45 AM. 30 is 12:30 AM, 5 is 12:05 AM.
- Values with a colon (14:30), seconds (143015), decimals (905.5), minus signs, letters and more than 4 characters show Check. A cell that already holds a real time also shows Check (it is not a 4-digit number); it does not need converting.
- Column D is text for labels and messages; use column C for sorting and math.
- Not tested: Excel, Google Sheets (TEXT with "h:mm AM/PM" in a non-US locale, and B4/100 on the text "0745" in the core formula, are the open questions).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- 1460 shows 3:00 PM with the core formula (probe); the file shows Check. The minutes part was 60.
- 2400 shows 12:00 AM with the core formula (probe); the file shows Check.
- A pasted 14:30 shows 12:00 AM with the core formula (probe: the text "14:30" divided by 100 gave a time of 0:00:36); the file shows Check. Text that already has a colon is not an HHMM number.
- A time with spaces around it (" 815 ") works in the file because of TRIM (8:15 AM).

---
## 5. check-upc-barcode-check-digit-mid-sumproduct-mod-excel-google-sheets
Title: How to Check a UPC Barcode Number's Check Digit in Excel and Google Sheets
Template: templates/tidy-tabs-upc-check-digit.xlsx (tabs: Barcodes, How to use)

Layout
- Barcodes: A1 note. Header row 3, data rows 4:203 (12 fictional products and codes). A Product, B UPC as typed (inputs; B formatted as text); C Digits only, D Length, E Check digit from first 11, F Result, G Note (formulas). Summary I3:J9: UPCs listed, OK, Wrong check digit, 11 digits, Check, Blank UPC, CHECK. F is green for OK, red for Check, Wrong check digit, 11 digits and Blank UPC.

Key formulas (row 4)
- C4 `=SUBSTITUTE(SUBSTITUTE(TRIM(B4)," ",""),"-","")`; D4 `=IF(C4="","",LEN(C4))`
- E4 `=IF(OR(D4=11,D4=12),IF(SUMPRODUCT(--ISNUMBER(FIND(MID(C4,{1,2,3,4,5,6,7,8,9,10,11},1),"0123456789")))=11,MOD(10-MOD(SUMPRODUCT(--MID(C4,{1,2,3,4,5,6,7,8,9,10,11},1),{3,1,3,1,3,1,3,1,3,1,3}),10),10),""),"")` (the post teaches the core `=MOD(10-MOD(SUMPRODUCT(--MID(C4,{1,2,3,4,5,6,7,8,9,10,11},1),{3,1,3,1,3,1,3,1,3,1,3}),10),10)`)
- F4 `=IF(C4="",IF(TRIM(A4)="","","Blank UPC"),IF(E4="","Check",IF(D4=12,IF(ISNUMBER(FIND(RIGHT(C4,1),"0123456789")),IF(--RIGHT(C4,1)=E4,"OK","Wrong check digit"),"Check"),"11 digits")))`
- G4 `=IF(F4="Wrong check digit","Last digit is "&RIGHT(C4,1)&", the first 11 give "&E4,IF(F4="11 digits",IF(MOD(10-MOD(SUMPRODUCT(--MID(C4,{1,2,3,4,5,6,7,8,9,10},1),{1,3,1,3,1,3,1,3,1,3}),10),10)=--RIGHT(C4,1),"0"&C4&" passes with a 0 in front; or add "&E4&" at the end","Add "&E4&" at the end: "&C4&E4),""))`
- J9 `=IF(J3=J4+J5+J6+J7+J8,"OK","Check")`
- How the math works (for the post): digits 1, 3, 5, 7, 9, 11 count 3 times, digits 2, 4, 6, 8, 10 once; the check digit brings the total up to the next multiple of 10 (0 if it already is one). Worked example matching the file: first 11 digits of 084512300179 are 0 8 4 5 1 2 3 0 0 1 7; odd positions 0+4+1+3+0+7 = 15, times 3 = 45; even positions 8+5+2+0+1 = 16; total 61; 70 - 61 = 9, and the file shows 9 and OK.

Sample values (fictional codes; recalculated)
| Product | UPC as typed | Digits only | Length | Check digit | Result | Note |
|---|---|---|---|---|---|---|
| Lavender soy candle 8 oz | 084512300179 | 084512300179 | 12 | 9 | OK | |
| Cedar soy candle 8 oz | 0 84512 30018 6 | 084512300186 | 12 | 6 | OK | |
| Oatmeal goat milk soap | 0-84512-30025-4 | 084512300254 | 12 | 4 | OK | |
| Charcoal soap bar | 084512300328 | 084512300328 | 12 | 2 | Wrong check digit | Last digit is 8, the first 11 give 2 |
| Brass hoop earrings | 84512300407 (stored as a number) | 84512300407 | 11 | 4 | 11 digits | 084512300407 passes with a 0 in front; or add 4 at the end |
| Letterpress card set | 08451230050 | 08451230050 | 11 | 6 | 11 digits | Add 6 at the end: 084512300506 |
| Ceramic mug 12 oz | 0845123006120 | 0845123006120 | 13 | | Check | |
| Macrame plant hanger | O84512300759 (letter O) | O84512300759 | 12 | | Check | |
| Leather key fob | (blank) | | | | Blank UPC | |
| Canvas tote bag | 084512300896 | 084512300896 | 12 | 6 | OK | |
| Kraft gift box | 0845123009 | 0845123009 | 10 | | Check | |
| Beeswax wrap set | 084512 300933 | 084512300933 | 12 | 3 | OK | |
Summary J3:J9: 12, 5, 1, 2, 3, 1, OK.

Pitfalls the post must state
- Store UPCs as text (format column B as Plain text before pasting). A UPC stored as a number loses its leading 0 and shows 11 digits (Brass hoop earrings). Link the existing keep-leading-zeros post.
- 11 digits is ambiguous: a dropped leading 0 or a missing check digit. The note says which one passes; only the reader knows which is right. When both are possible (like the earrings), the note says both.
- Wrong check digit means at least one digit is wrong somewhere, not always the last one.
- The check does not catch every mistake: swapping two neighboring digits that differ by 5 still passes (probe: 084512300162 and 084512300612 are both OK). Swaps such as 8 and 4 were caught.
- Spaces and hyphens are removed first; letters (O for 0), full width digits and other characters show Check. 13-digit EAN/GTIN-13 codes show Check (not covered by this file).
- Arithmetic only: it does not say a code is registered or belongs to the reader. GS1 US assigns company prefixes (link https://www.gs1us.org); do not tell readers to make up codes for products they sell.
- Not tested: Excel, Google Sheets (SUMPRODUCT with array constants inside MID is the open question).

Troubleshooting cases actually reproduced in LibreOffice 24.2.7.2
- A good code shows 11 digits: it was stored as a number and lost its leading 0 (84512300407; the note says 084512300407 passes).
- Check on a code that looks right: a letter O instead of a zero (O84512300759), a full width digit, or 13 digits (0845123006120), or 10 digits (0845123009).
- Wrong check digit: the last digit was typed wrong (084512300328, first 11 give 2).
- A code with a typo still shows OK: neighbors 1 and 6 swapped (084512300162 vs 084512300612, both OK).
- Spaces or hyphens from a printed label (0 84512 30018 6, 0-84512-30025-4) are fine; the file removes them.
