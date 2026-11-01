# Team topics - 2026-11-02 queue (5 new post topics)

Posts dated 2026-11-02 PT. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc 24.2, a download that gives a reason to visit, no tax, wage-law or legal judgment). Each was checked by grep against slugs and titles in posts/ready and posts/published, tasks/topics-seed.md and every tasks/team-topics-*.md. Grepped for: unpaid, largest, biggest, LARGE, reminder, workshop, class, seat, capacity, ticket, card fee, processing, gross up, net target, number from text, oz, units, appointment, time slot, YYYYMMDD, price in effect, handmade, overhead. Three first-draft ideas were dropped as duplicates: text dates to real dates (convert-yyyymmdd-number-to-date-date-left-mid), price in effect on a date (price-in-effect-on-order-date-price-history-lookup), time-slot schedule with a lunch break (appointment-schedule-end-times-gaps-buffer). Also dropped: pricing a handmade item with labor and overhead (too close to batch-cost-per-unit and markup-vs-margin), late fee calculator (legal flavor), refundable deposit rules (legal flavor).

Closest existing posts and how these differ: marketplace-fee-net-payout-calculator (forward: sale price in, net out; topic 2 is the reverse, net wanted in, price out); rank-top-customers (RANK on summed spend; topic 3 lists the largest open invoices with LARGE and breaks ties); receivables-aging-report and highlight-overdue-invoices (buckets and colors; topic 3 and topic 5 produce a short list and a ready-to-paste sentence); client-retainer-hours-remaining and break-even-calculator (hours or units used against a limit; topic 1 counts paid and waitlist registrations against seats); extract-domain-from-email and split-sku-code (split text on a character; topic 4 returns a number with a unit stripped so it can be summed).

Function families vary from the last 3 days (ROUNDUP, COUNTIF/INDEX/MATCH, MOD, MAXIFS, SUMIFS on 10-30; SUMIF, SUMIFS, INDEX/MATCH, SUMPRODUCT on 10-31; COUNTIF, SUMIF, SUMPRODUCT, ROUND, TODAY on 11-01): COUNTIFS with MAX/MIN, ROUND with a reverse fee formula, LARGE with MATCH/INDEX, VALUE/LEFT/FIND/SUBSTITUTE, and TEXT with string joining. No FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TODAY, NOW, dynamic arrays or WEBSERVICE. No tax, wage, contract or financial-planning advice: arithmetic on fictional numbers only. Mix: three templates with sample data (1, 3, 5), two formula how-tos (2, 4).

Formulas were NOT run in LibreOffice while picking these topics. Nothing below is a recorded test. Each topic's Verify line says what the build must run; only claim LibreOffice in `tested_in` after that run. Excel and Google Sheets are not tested for any topic; link the Microsoft and Google help pages for the functions.

Ranked best first. Titles are final picks; opening paragraphs are the answer-first paragraph for the script-writer to open the post with. The openings state formulas that the build must confirm before publishing.

---

## 1. Track workshop registrations and seats left

- Title: How to Track Workshop Registrations and Seats Left in Excel and Google Sheets
- Slug: `workshop-registrations-seats-left-waitlist-countifs-excel-google-sheets`
- Opening paragraph: Count paid sign-ups for a workshop with `=COUNTIFS(Signups!$B$5:$B$204,A5,Signups!$D$5:$D$204,"Paid")`, then show the seats left with `=MAX(0,C5-D5)`. Anyone who signs up after the class is full goes on the waitlist, which you get with `=MAX(0,E5-C5)` from a separate count of all sign-ups.
- Other title options considered: How to Count Seats Left for a Class or Craft Workshop in Google Sheets; How to Keep a Sign-Up Sheet and Waitlist for Events in Excel
- Search wording people type: "workshop registration spreadsheet", "class sign up sheet google sheets", "seats remaining formula excel", "countifs by class name and status", "waitlist spreadsheet template"
- Template spec: `tidy-tabs-workshop-registrations.xlsx` (Workshops, Signups, How to use). All people and workshops are fictional; the post says so once.
  - Workshops (yellow input A4:C4 headers, rows 5-14): Workshop (`Candle Pouring 101`, `Intro to Block Printing`, `Embroidery Basics`, `Wreath Making`), Date (`11/14/2026`, `11/21/2026`, `12/05/2026`, `12/12/2026`), Seats (`8`, `10`, `6`, `12`).
    - D Paid `=IF(A5="","",COUNTIFS(Signups!$B$5:$B$204,A5,Signups!$D$5:$D$204,"Paid"))`
    - E All sign-ups `=IF(A5="","",COUNTIF(Signups!$B$5:$B$204,A5))`
    - F Seats left `=IF(A5="","",MAX(0,C5-D5))`
    - G Waitlist `=IF(A5="","",MAX(0,E5-C5))` (all sign-ups beyond the seat count, paid or not; the post states this plainly)
    - H Status `=IF(A5="","",IF(F5=0,"Full",IF(F5<=2,"Almost full","Open")))`
  - Signups (yellow input A:D, header row 4, rows 5-204): Name, Workshop, Email (fictional `@example.com`), Payment (`Paid`, `Unpaid`). 28 fictional rows; one workshop must be full (6 paid on Embroidery Basics), one must have a waitlist (a seventh sign-up), one must show Almost full.
  - Summary (Workshops J1:K6): workshops count `=COUNTA(A5:A14)`, total seats `=SUM(C5:C14)`, total paid `=SUM(D5:D14)`, total seats left `=SUM(F5:F14)`, sign-ups with a workshop name not on the list `=SUMPRODUCT((Signups!B5:B204<>"")*(COUNTIF(A5:A14,Signups!B5:B204)=0))` (if that form fails in the build, use a helper column on Signups).
  - CHECK cell: `=IF(AND(K3<=K2,K6=0),"OK","Check")` (paid never exceeds seats in the sample, and no sign-up points at a misspelled workshop).
  - The post does not say anything about refunds, cancellation terms or liability waivers.
- Verify plan: Python model for the sample and 10 edited copies (sign-ups exactly equal to seats, one over, zero sign-ups, a workshop name with different capitalization (COUNTIFS ignores case), a trailing space in the workshop name (must be caught by the unmatched count, not silently dropped), Payment typed "paid " with a trailing space (not counted, show the effect), 0 seats, blank rows in the middle, 200 sign-ups). LibreOffice headless to recalculate every copy, read with `data_only`, compare every D:H cell and the summary, confirm no error values. Not run: Excel, Sheets.
- Page-1 competitors angle: Page 1 will likely be Eventbrite and Jotform style sign-up pages, Vertex42 and Smartsheet templates, and forum threads. Angle: a plain sheet for the maker who runs four classes a year and does not want a booking app, with seats left, a waitlist count and an "almost full" flag.

---

## 2. Set a price that nets the amount you want after a card fee

- Title: How to Set a Price That Still Nets the Amount You Want After a Card Fee in Excel and Google Sheets
- Slug: `price-to-net-target-after-card-fee-excel-google-sheets`
- Opening paragraph: To end up with a target amount after a percent fee plus a flat fee, divide instead of multiply: `=ROUND((B2+B4)/(1-B3),2)`, where B2 is the amount you want to keep, B3 the fee percent and B4 the flat fee. Multiplying the target by 1.03 comes up short because the fee is charged on the fee too.
- Other title options considered: How to Work Backward From the Amount You Want to Keep to the Price to Charge in Google Sheets; How to Add a Card Fee to a Price Without Losing Money in Excel
- Search wording people type: "price to cover credit card fee formula", "gross up price for processing fee excel", "how to calculate price to net amount after fee", "add 2.9% + 30 cents to price spreadsheet"
- Template spec: `tidy-tabs-price-after-card-fee.xlsx` (Calculator, How to use). All numbers are fictional examples; the post says the fee rates are typed by the reader and tells them to take current rates from their own processor, with no claim about any provider's actual rates.
  - Calculator (yellow input B2:B4): B2 Amount you want to keep `$100.00`, B3 Percent fee `2.9%` (formatted 0.0%), B4 Flat fee `$0.30`.
    - B6 Price to charge `=ROUND((B2+B4)/(1-B3),2)`
    - B7 Fee taken `=ROUND(B6*B3+B4,2)`
    - B8 You keep `=ROUND(B6-B7,2)`
    - B9 Difference from target `=ROUND(B8-B2,2)`
    - B10 Wrong shortcut price `=ROUND(B2*(1+B3)+B4,2)`, B11 what you keep with it `=ROUND(B10-(B10*B3+B4),2)`, B12 shortfall `=ROUND(B11-B2,2)` (shows why multiplying fails).
    - B13 Status `=IF(B9>=0,"OK",IF(B9>=-0.01,"1 cent short from rounding","Check"))`
  - Table (A16:E22): five fictional target amounts ($25.00, $48.00, $100.00, $250.00, $1,250.00) with the same formulas in columns B:E (Price to charge, Fee, You keep, Difference). Totals are not needed.
  - CHECK cell: `=IF(AND(B6>0,B3<1,B3>=0,COUNTIF(E17:E21,"<-0.01")=0),"OK","Check")` (fee percent below 100%, no row more than one cent short).
  - The post is arithmetic only. It does not say whether a business may pass a fee to customers; it tells the reader to check their processor's and state's rules and links no legal claim.
- Verify plan: Python `Decimal` ROUND_HALF_UP for the five rows and 12 edited inputs (fee 0%, flat fee $0.00, fee 100% must not divide by zero and must show Check, a target of $0.01, a half-cent case that leaves the result 1 cent short, large targets, a percent typed as 2.9 instead of 2.9% (must not give a price near 0; add a guard that the percent is below 1), text instead of numbers). LibreOffice headless to recalculate and compare B6:B13 and the table. Decide in the build whether to switch to ROUNDUP for the price; the verify run shows how many of the edited copies end up one cent short with ROUND. Not run: Excel, Sheets.
- Page-1 competitors angle: Page 1 will likely be payment-processor blog posts and online fee calculators. Angle: the algebra in one cell, a side-by-side that shows what the wrong multiply-by-1.03 shortcut costs, and a sheet the reader can keep for their own rates.

---

## 3. Show your 3 biggest unpaid invoices

- Title: How to Show Your 3 Biggest Unpaid Invoices in Excel and Google Sheets
- Slug: `three-biggest-unpaid-invoices-large-match-excel-google-sheets`
- Opening paragraph: Find the largest open amount with `=LARGE($F$5:$F$104,1)` and change the 1 to 2 and 3 for the next two. To get the invoice number beside each amount, match it back with `=INDEX($A$5:$A$104,MATCH(B5,$F$5:$F$104,0))`. Paid invoices stay out because column F is blank for them.
- Other title options considered: How to List the Top 3 Open Invoices by Amount in Google Sheets; How to Find the Largest Amounts Still Owed to You in Excel
- Search wording people type: "largest unpaid invoices excel", "top 3 values with name excel large index match", "list top 5 outstanding invoices google sheets", "large function ties duplicate values"
- Template spec: `tidy-tabs-top-unpaid-invoices.xlsx` (Invoices, Top 3, How to use). All clients are fictional; the post says so once.
  - Invoices (yellow input A:E, header row 4, rows 5-104): Invoice # (`INV-2041`), Client, Amount, Due date, Status (`Paid`, `Sent`, `Overdue`). 18 fictional rows with amounts from $85.00 to $2,400.00; include two open invoices with the same amount (for example $640.00) to show ties.
    - F Open amount `=IF(OR(A5="",E5="Paid",NOT(ISNUMBER(C5))),"",C5)`
    - G Tie-safe key `=IF(F5="","",F5+ROW()/1000000)` (adds a tiny row-based amount so two equal invoices get different keys; the post says the key is never shown to clients, hide the column or grey it).
  - Top 3 sheet (A4:D7): Rank 1, 2, 3 in A5:A7.
    - B Key `=IFERROR(LARGE(Invoices!$G$5:$G$104,A5),"")`
    - C Invoice # `=IF(B5="","",INDEX(Invoices!$A$5:$A$104,MATCH(B5,Invoices!$G$5:$G$104,0)))`
    - D Client `=IF(B5="","",INDEX(Invoices!$B$5:$B$104,MATCH(B5,Invoices!$G$5:$G$104,0)))`
    - E Amount `=IF(B5="","",INDEX(Invoices!$F$5:$F$104,MATCH(B5,Invoices!$G$5:$G$104,0)))`
  - Summary: open invoices `=COUNT(Invoices!F5:F104)`, total open `=SUM(Invoices!F5:F104)`, top 3 total `=SUM(E5:E7)`, top 3 share of open `=IF(total=0,"",ROUND(top3/total,4))`.
  - CHECK cell: `=IF(AND(ROUND(SUM(Invoices!F5:F104)-SUMIF(Invoices!E5:E104,"<>Paid",Invoices!C5:C104),2)=0,E5>=E6,E6>=E7),"OK","Check")`.
- Verify plan: Python model for the sample and 10 edited copies (fewer than 3 open invoices so the rows go blank, all paid, a tie for first place, a tie across ranks 3 and 4, Status typed "paid" lowercase (comparison with "Paid" is case-insensitive in Calc; show the result), amount as text "640", a blank amount, negative amount (credit), 100 rows). LibreOffice headless to recalculate and compare Top 3 rows and the summary; confirm no error values. Note whether the tie key keeps the displayed amount unchanged (the amount comes from column F, not G). Not run: Excel, Sheets.
- Page-1 competitors angle: Page 1 will likely be Microsoft's LARGE help page, ExcelJet and forum answers that return the value without the name. Angle: a ready invoice sheet, ties handled with a visible helper column, and the name and client beside each amount, for a freelancer deciding whom to follow up with first.

---

## 4. Pull the number out of text like 12 oz or 3 lb

- Title: How to Pull the Number Out of Text Like 12 oz or 3 lb in Excel and Google Sheets
- Slug: `pull-number-out-of-text-12-oz-value-left-find-excel-google-sheets`
- Opening paragraph: For text like `12 oz` with a space, take everything before the space with `=VALUE(LEFT(TRIM(A5),FIND(" ",TRIM(A5)&" ")-1))`. For text like `12oz` with no space, strip the unit instead: `=VALUE(SUBSTITUTE(LOWER(TRIM(A5)),"oz",""))`. Either way you get a real number you can add up.
- Other title options considered: How to Turn Weights Typed as Text Into Numbers You Can Add in Google Sheets; How to Sum Cells That Contain Units Like oz or lb in Excel
- Search wording people type: "extract number from text with units excel", "sum cells with oz in them", "remove unit from cell and keep number", "excel value left find space", "convert 12 oz text to number google sheets"
- Template spec: `tidy-tabs-number-from-text-units.xlsx` (Weights, How to use). All items are fictional; the post says so once.
  - Weights (yellow input A4:B4 headers, rows 5-24): Item (`Soy wax`, `Fragrance oil`, `Tins, 8 oz`, `Shipping box`), Typed weight (as text: `12 oz`, `12oz`, ` 3 lb `, `16 OZ`, `1.5 lb`, `n/a`).
    - C Number `=IF(A5="","",IFERROR(VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1)),IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""),"lb","")),"Check")))`
    - D Unit `=IF(OR(A5="",C5="Check"),"",IF(ISNUMBER(SEARCH("lb",B5)),"lb","oz"))`
    - E Ounces `=IF(ISNUMBER(C5),IF(D5="lb",ROUND(C5*16,2),C5),"")`
  - Summary: total ounces `=SUM(E5:E24)`, pounds and ounces shown as `=INT(total/16)&" lb "&ROUND(MOD(total,16),2)&" oz"`, count of rows marked Check `=COUNTIF(C5:C24,"Check")`.
  - CHECK cell: `=IF(AND(COUNTIF(C5:C24,"Check")=0,COUNT(E5:E24)=COUNTA(A5:A24)),"OK","Check")` (the committed sample keeps one `n/a` row in a separate edited copy only, so the committed file shows OK).
  - Post notes that numbers copied with a non-breaking space need the SUBSTITUTE(CHAR(160)) step from the earlier non-breaking-space post, and that decimal commas depend on the computer's region setting.
- Verify plan: Python parser mirroring the formula for the sample and 14 edited copies (`12 oz`, `12oz`, `12  oz` double space, ` 3 lb ` with outer spaces, `16 OZ`, `1.5 lb`, `1,5 lb` (shows Check or a wrong value in a US locale, document which), `n/a`, empty string, a plain number cell `12` (VALUE of a number, TRIM turns it into text), `12 oz.` with a period, `12 fl oz` (must not be read as 12 oz silently; decide and document), `oz 12`, `.5 lb`). LibreOffice headless to recalculate and compare C:E. Note the `lb` text containing "oz" or "oz" containing "lb" cannot overlap, but a unit like `ounce` would break the SUBSTITUTE path; document it. Not run: Excel, Sheets.
- Page-1 competitors angle: Page 1 will likely be generic "extract number from string" tutorials with array or LAMBDA solutions, plus forum threads. Angle: two short formulas for the exact `12 oz` shapes a seller types, no array formulas, and a unit conversion to total weight for a shipping box.

---

## 5. Write a friendly payment reminder from invoice data

- Title: How to Write a Friendly Payment Reminder From Your Invoice List in Excel and Google Sheets
- Slug: `payment-reminder-message-from-invoice-data-text-excel-google-sheets`
- Opening paragraph: Build the message from the cells with `="Hi "&B5&", a quick reminder that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" was due on "&TEXT(D5,"mm/dd/yyyy")&"."`. Copy it down and each row gives you a sentence you can paste into an email, with the amount and date formatted the US way.
- Other title options considered: How to Make a Payment Reminder Email Text From a Spreadsheet in Google Sheets; How to Join Text and Formatted Numbers Into One Sentence in Excel
- Search wording people type: "payment reminder template excel", "combine text and date in one cell excel", "text function dollar amount in sentence", "invoice reminder message google sheets", "concatenate date shows number"
- Template spec: `tidy-tabs-payment-reminders.xlsx` (Invoices, Settings, How to use). Fictional clients; the post says so once.
  - Settings (yellow): B1 As-of date `11/02/2026` (typed, not TODAY, so messages do not change when the file is opened later), B2 Your name `Jordan` (fictional).
  - Invoices (yellow input A:E, header row 4, rows 5-64): Invoice # (`INV-2051`), Client first name, Amount, Due date, Paid (`Yes` or blank). 12 fictional rows; some paid, some due in the future, some past due by 3, 14 and 45 days.
    - F Days past due `=IF(OR(A5="",E5="Yes",NOT(ISNUMBER(D5))),"",Settings!$B$1-D5)`
    - G Tone `=IF(F5="","",IF(F5<=0,"Upcoming",IF(F5<=14,"Gentle","Follow-up")))`
    - H Message `=IF(F5="","",IF(F5<=0,"Hi "&B5&", just a heads-up that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" is due on "&TEXT(D5,"mm/dd/yyyy")&". Thanks, "&Settings!$B$2,"Hi "&B5&", a quick reminder that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" was due on "&TEXT(D5,"mm/dd/yyyy")&" ("&F5&" days ago). Let me know if you need another copy. Thanks, "&Settings!$B$2))`
  - Summary: messages to send `=COUNT(F5:F64)-COUNTIF(F5:F64,"<=0")`, total past due `=SUMIF(F5:F64,">0",C5:C64)`.
  - CHECK cell: `=IF(AND(COUNTIF(H5:H64,"*#*")=0,COUNTIF(H5:H64,"*NaN*")=0,SUMPRODUCT((F5:F64<>"")*(H5:H64=""))=0),"OK","Check")` (every row with a day count has a message and none shows an error or a raw serial number).
  - The message is plain and polite: no late fee, no deadline threat, no legal wording. The post says it does not give advice on collections and links nothing legal.
- Verify plan: Python string model with `datetime` for the sample and 10 edited copies (due date equal to the as-of date, one day after, 14 and 15 days past due, amount with cents like $1,250.50 and a whole $85, amount as text, blank client name (message must not read "Hi ,"; add a guard), a date typed as text, Paid typed "yes" lowercase (equality is case-insensitive), a leap-day date). Compare the exact message strings after LibreOffice headless recalculation, including that TEXT with `"$#,##0.00"` and `"mm/dd/yyyy"` returns the US format in LibreOffice en-US. State in the post that TEXT format codes follow the app's language setting and give the `Format > Number` alternative. Not run: Excel, Sheets, TODAY().
- Page-1 competitors angle: Page 1 will likely be invoicing-software blogs with email wording templates and Microsoft's TEXT function help page. Angle: wording generated from the sheet itself, three tones chosen by days past due, and the TEXT codes explained for the reader whose date shows as 46327.
