# Handoff 2026-11-07: five topics researched, hooks written, templates built and tested (LibreOffice Calc 24.2.7.2 only), for the script-writers

Stage 1 of the content team (researcher, hook-writer, builder) was done by one agent. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: tasks/team-topics-2026-11-07.md, tasks/hooks-2026-11-07.md, scripts/build-templates.py (`split_file_name_extension`, `return_deadline`, `lookup_by_part_of_name`, `customers_by_order_count`, `wholesale_from_retail`), five new files in templates/, results in tasks/briefs.md ("2026-11-07 builder results"), titles appended to tasks/topics-seed.md.

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 99 workbooks, 53,239 comparisons, 0 mismatches against independent Python models, no error value in any recalculated workbook. All five files exported to PDF: every page US Letter, no error values in the PDF text. Two bugs found and fixed (AND without short circuit giving #VALUE! on text; a test model that forgot wildcard escaping), plus removed test cases, in tasks/briefs.md.

Repo note for the lead: running the build script rewrote all older templates (metadata only; 165 tracked files show as modified). `git checkout -- templates` was refused by the permission system in this session; run it before committing so only the five new .xlsx files change.

Post file names: `posts/ready/2026-11-07-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-07-team-01-split-file-name-and-extension-last-dot-substitute-excel-google-sheets | tidy-tabs-split-file-name-extension.xlsx | 11 workbooks, 6,754 comparisons |
| 2 | 2026-11-07-team-02-return-by-date-from-delivery-date-days-left-excel-google-sheets | tidy-tabs-return-deadline.xlsx | 26 workbooks, 13,208 comparisons |
| 3 | 2026-11-07-team-03-look-up-price-by-part-of-product-name-wildcard-match-excel-google-sheets | tidy-tabs-lookup-by-part-of-name.xlsx | 16 workbooks, 2,336 comparisons |
| 4 | 2026-11-07-team-04-customers-by-number-of-orders-frequency-table-countif-excel-google-sheets | tidy-tabs-customers-by-number-of-orders.xlsx | 19 workbooks, 17,765 comparisons |
| 5 | 2026-11-07-team-05-wholesale-price-from-retail-price-margin-check-excel-google-sheets | tidy-tabs-wholesale-from-retail-price.xlsx | 27 workbooks, 13,176 comparisons |

Every workbook count includes the committed sample. "Comparisons" are cell-by-cell checks of the recalculated workbook against the Python model. Formulas use no TODAY, FILTER, LET, XLOOKUP, SORT, UNIQUE or dynamic arrays and no _xlfn functions.

---

## 1. How to Split a File Name Into Name and Extension in Excel and Google Sheets
Hook: Find the last dot with `=FIND(CHAR(1),SUBSTITUTE(B7,".",CHAR(1),LEN(B7)-LEN(SUBSTITUTE(B7,".",""))))`, then take the extension with `=MID(B7,C7+1,500)` and the name with `=LEFT(B7,C7-1)`. For "Invoice.2026.11.docx" the last dot is at position 16, the extension is docx and the name is Invoice.2026.11.

Template: templates/tidy-tabs-split-file-name-extension.xlsx. Sheets: **Files**, **How to use**. Landscape on the Files tab.

Layout (Files): A1 title, A2 note. Header row 4, input rows 5 to 104 (100 names, column A yellow). B "File name, spaces trimmed", C "Last dot at position", D "Extension", E "Name without extension", F "Status", G "Extension in lowercase" (all gray). Summary I4:J9: J4 names listed, J5 OK, J6 No extension, J7 starts or ends with a dot, J8 space in the extension, J9 CHECK. Count table: I12:J12 headers, yellow extensions I13:I20 (sample png, pdf, svg, jpg, zip; three empty slots), J13:J20 count of files with that lowercase extension.

Key formulas (row 5):
- B5 `=IF(A5="","",TRIM(A5))`
- C5 `=IF(B5="","",IF((LEN(B5)-LEN(SUBSTITUTE(B5,".","")))=0,0,FIND(CHAR(1),SUBSTITUTE(B5,".",CHAR(1),(LEN(B5)-LEN(SUBSTITUTE(B5,".","")))))))`
- D5 `=IF(B5="","",IF(OR(C5=0,C5=1,C5=LEN(B5)),"",MID(B5,C5+1,500)))`
- E5 `=IF(B5="","",IF(OR(C5=0,AND(C5=1,LEN(B5)>1)),B5,LEFT(B5,C5-1)))`
- F5 `=IF(B5="","",IF(C5=0,"No extension",IF(C5=LEN(B5),"Ends with dot",IF(C5=1,"Starts with dot",IF(ISNUMBER(FIND(" ",D5)),"Check extension","OK")))))`
- G5 `=IF(B5="","",LOWER(D5))`
- J13 `=IF(SUBSTITUTE(TRIM(I13),".","")="","",COUNTIF($G$5:$G$104,SUBSTITUTE(TRIM(I13),".","")))`
- J9 `=IF(AND(SUM(J5:J8)=J4,SUMPRODUCT(--ISERROR(B5:G104))=0,SUMPRODUCT((F5:F104="OK")*(E5:E104&"."&D5:D104<>B5:B104))=0),"OK","Check")`

Sample (12 names, rows 5 to 16): J4 12, J5 10, J6 1, J7 0, J8 1, J9 OK; counts png 2, pdf 2, svg 1, jpg 1, zip 1.
- A5 Mug-Wrap-Design_v2.png: C5 19, D5 png, E5 Mug-Wrap-Design_v2.
- A7 Invoice.2026.11.docx: C7 16, D7 docx, E7 Invoice.2026.11.
- A8 wedding-sign.final.SVG: D8 SVG, E8 wedding-sign.final, G8 svg.
- A10 Logo.tar.gz: D10 gz, E10 Logo.tar.
- A12 Packing list: C12 0, D12 blank, E12 Packing list, F12 No extension.
- A15 Dr. Smith flyer: C15 3, D15 " Smith flyer", E15 Dr, F15 Check extension.

Caveats the post must state: the extension is whatever follows the LAST dot, so a name like "Dr. Smith flyer" is flagged Check extension rather than guessed; ".gitignore" and "report." get their own statuses and an empty extension; only ordinary spaces are trimmed (TRIM does not remove a non-breaking space; not tested); paste names, not folder paths (a dot in a folder name would be taken as the last dot; not tested); CHAR(1) is used as an invisible marker and a name that already contains it (not tested) breaks the position; 500 characters is the cap after the dot; a plain number typed as a name (123) is treated as text with no extension, a decimal such as 12.5 was NOT tested. Not tested: Excel, Google Sheets.

---

## 2. How to Calculate a Return-By Date From the Delivery Date in Excel and Google Sheets
Hook: Add the return window to the delivery date with `=INT(B11)+D11`, then count the days left with `=E11-$B$3`, where B3 is an as-of date you type. A parcel delivered 10/20/2026 with a 30 day window is returnable until Thursday 11/19/2026, which is 4 days left as of 11/15/2026 (your own shop policy, not a legal rule).

Template: templates/tidy-tabs-return-deadline.xlsx. Sheets: **Orders**, **How to use**. Landscape on Orders.

Layout (Orders): A2 "Default return window (days)", **B2** 30 (yellow). A3 as-of date, **B3** 11/15/2026 (yellow, typed on purpose, no TODAY). A4 "Call it Closing soon when this many days or fewer are left", **B4** 7 (yellow). A5 "Settings check", **B5**. Header row 7, input rows 8 to 107 (100 orders). A order # (text), B "Delivered on" (yellow date), C "Window for this order (days, optional)" (yellow), D window used, E return-by date, F return-by day (ddd), G days left as of date, H status (Open, Closing soon, Last day, Closed, Check). Summary J7:K13: K7 delivered orders listed, K8 Open, K9 Closing soon, K10 Last day, K11 Closed, K12 rows showing Check, K13 CHECK.

Key formulas (row 8):
- B5 `=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B4)),IF(AND(B2>=0,B2=INT(B2),B3>=1,B3=INT(B3),B4>=0,B4=INT(B4)),"OK","Check"),"Check")`
- D8 `=IF(B8="","",IF($B$5<>"OK","Check",IF(C8="",$B$2,IF(ISNUMBER(C8),IF(AND(C8>=0,C8=INT(C8)),C8,"Check"),"Check"))))`
- E8 `=IF(B8="","",IF(AND(ISNUMBER(B8),ISNUMBER(D8)),IF(AND(B8>=1,INT(B8)+D8<=2958465),INT(B8)+D8,"Check"),"Check"))`
- F8 `=IF(ISNUMBER(E8),TEXT(E8,"ddd"),E8)`
- G8 `=IF(ISNUMBER(E8),E8-$B$3,E8)`
- H8 `=IF(ISNUMBER(G8),IF(G8<0,"Closed",IF(G8=0,"Last day",IF(G8<=$B$4,"Closing soon","Open"))),G8)`
- K13 `=IF(AND(B5="OK",K12=0,SUM(K8:K11)=K7),"OK","Check")`

Sample (12 rows, rows 8 to 19; row 19, order 1052, has no delivery date and stays blank): K7 11, K8 5, K9 2, K10 1, K11 3, K12 0, K13 OK.
- Row 8 (1041) delivered 10/09/2026, 30 days: return-by 11/08/2026 Sun, -7, Closed.
- Row 11 (1044) delivered 10/20/2026: 11/19/2026 Thu, 4, Closing soon.
- Row 13 (1046) delivered 11/02/2026, override 14: 11/16/2026 Mon, 1, Closing soon.
- Row 14 (1047) delivered 11/04/2026: 12/04/2026 Fri, 19, Open.
- Row 16 (1049) delivered 11/13/2026, override 60: 01/12/2027 Tue, 58, Open.
- Row 18 (1051) delivered 11/15/2026, override 0: 11/15/2026 Sun, 0, Last day.

Caveats the post must state: the window is the shop's own policy in calendar days, not a legal rule, and the post gives no advice on what a buyer is entitled to; the delivery date counts as day 0 (return-by = delivery date + window) and a window of 0 makes the delivery day the last day; a time of day on the delivery date is dropped (INT); the as-of date is typed because TODAY() changes every time the file opens (so a printed copy cannot be matched to a date); a text date, a time-only value, a decimal or negative window shows Check; dates after 12/31/9999 (serial 2,958,465) show Check; carrier-promised delivery dates are not looked up. Not tested: Excel, Google Sheets, TODAY(), a non-English ddd weekday name, serials below 61.

---

## 3. How to Look Up a Price From Part of a Product Name in Excel and Google Sheets
Hook: Put a * on each side of the word you typed, so `=MATCH("*"&B9&"*",Products!$B$5:$B$104,0)` returns the first product whose name contains it, and `=INDEX(Products!$C$5:$C$104,D9)` returns its price. Typing tote finds "Canvas tote bag, 12 oz canvas" at $22.00 (fictional products).

Template: templates/tidy-tabs-lookup-by-part-of-name.xlsx. Sheets: **Search**, **Products**, **How to use**. Landscape on Search and Products.

Layout (Search): A1 title, A2 note. Header row 4, input rows 5 to 24 (20 search words, column A yellow). B "Search text used", C "Matches", D "First match is list row", E "First matching product", F "SKU", G "Price", H "Status" (One match, Several matches, No match, Too long). Summary J4:K9: K4 search words typed, K5 one match, K6 several, K7 no match, K8 too long (over 200 characters), K9 CHECK.
Layout (Products): A1 title, A2 note, header row 4 (SKU, Product name, Price), input rows 5 to 104 (100 products, yellow).

Key formulas (row 5 of Search):
- B5 `=IF(TRIM(A5)="","",IF(LEN(TRIM(A5))>200,"",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A5),"~","~~"),"*","~*"),"?","~?")))`
- C5 `=IF(B5="","",COUNTIF(Products!$B$5:$B$104,"*"&B5&"*"))`
- D5 `=IF(B5="","",IF(C5=0,"",MATCH("*"&B5&"*",Products!$B$5:$B$104,0)))`
- E5 `=IF(ISNUMBER(D5),INDEX(Products!$B$5:$B$104,D5),"")`; F5 the same on column A with &""; G5 the same on column C with "No price" when the price is not a number
- H5 `=IF(TRIM(A5)="","",IF(B5="","Too long",IF(C5=0,"No match",IF(C5=1,"One match","Several matches"))))`
- K9 `=IF(AND(K8=0,SUM(K5:K8)=K4,SUMPRODUCT(--ISERROR(B5:H24))=0),"OK","Check")`

Sample (12 products, 9 words): K4 9, K5 1, K6 5, K7 3, K8 0, K9 OK.
- A5 lavender: C5 2, D5 1, E5 Soy candle, Lavender, 8 oz, F5 CND-LAV8, G5 $18.00, Several matches.
- A7 "  mug  ": B7 mug, C7 2, D7 7, Enamel mug, Blue, 12 oz, MUG-BLU, $16.00.
- A9 tote: C9 1, D9 9, Canvas tote bag, 12 oz canvas, TOT-CAN, $22.00, One match.
- A10 gift, A12 zzz, A13 50% off: No match.

Caveats the post must state: only the FIRST matching row is returned, so Several matches is a warning to type more of the name; matching ignores capitals and trims spaces at the ends of the term, but double spaces INSIDE a product name are not collapsed; * ? ~ typed in the term are searched as plain characters (escaped with ~); terms over 200 characters show Too long (MATCH cannot search over 255); keep product names as text (in LibreOffice a name stored as the number 8 was counted, Microsoft documents COUNTIF wildcards for text only; not tested in Excel); wildcard use in MATCH and COUNTIF is documented by Microsoft (link the help page) but was NOT tested in Google Sheets; say so and tell readers to try one search before relying on it there. Not tested: Excel, Google Sheets.

---

## 4. How to Count Customers by Number of Orders in Excel and Google Sheets
Hook: Give each customer's first row a 1 with `=IF(SUMPRODUCT(--($C$5:C5=C5))=1,1,0)`, count their orders with `=SUMPRODUCT(--($C$5:$C$304=C5))`, then count customers who ordered exactly twice with `=COUNTIFS($E$5:$E$304,1,$D$5:$D$304,2)`. In the sample, 4 of 14 customers ordered twice.

Template: templates/tidy-tabs-customers-by-number-of-orders.xlsx. Sheets: **Orders**, **How to use**. Landscape on Orders.

Layout (Orders): A1 title, A2 note. Header row 4, input rows 5 to 304 (300 orders): A customer, B order amount (yellow); C customer with spaces trimmed, D orders by this customer, E first row of this customer (1 = yes) (gray). Table G4:L10: G "Orders placed" (1, 2, 3, 4, 5 or more), H customers, I share of customers, J orders, K sales, L share of sales, row 10 Total. Summary G12:H18: H12 orders listed, H13 customers, H14 average orders per customer, H15 most orders by one customer, H16 rows with an amount that is not a number, H17 amounts with no customer name, H18 CHECK.

Key formulas (row 5; table row 6 is "2 orders"):
- C5 `=IF(TRIM(A5)="","",TRIM(A5))`
- D5 `=IF(C5="","",SUMPRODUCT(--($C$5:$C$304=C5)))`
- E5 `=IF(C5="","",IF(SUMPRODUCT(--($C$5:C5=C5))=1,1,0))`
- H6 `=COUNTIFS($E$5:$E$304,1,$D$5:$D$304,2)`; H9 (5 or more) uses `">=5"`
- J6 `=COUNTIFS($D$5:$D$304,2)`; K6 `=SUMIFS($B$5:$B$304,$D$5:$D$304,2)`
- I6 `=IF($H$10=0,"",H6/$H$10)`; L6 `=IF($K$10=0,"",K6/$K$10)`
- H18 `=IF(AND(H10=H13,J10=H12,H16=0,H17=0,ROUND(K10-SUM(B5:B304),2)=0),"OK","Check")`

Sample (40 orders, 14 customers): H12 40, H13 14, H14 2.86, H15 7 (Maria Lopez), H18 OK.
- 1 order: 3 customers, 3 orders, $130.00, 21.4% of customers, 9.2% of sales.
- 2 orders: 4 customers, 8 orders, $200.00, 28.6%, 14.1%.
- 3 orders: 3 customers, 9 orders, $355.75, 21.4%, 25.1%.
- 4 orders: 2 customers, 8 orders, $303.50, 14.3%, 21.4%.
- 5 or more: 2 customers, 12 orders, $429.25, 14.3%, 30.3%.
- Total: 14 customers, 40 orders, $1,418.50.

Caveats the post must state: names are matched after TRIM and without regard to capitals, so Maria Lopez and maria lopez with a trailing space are one customer; two spellings (Jo Ann and JoAnn) are two customers, clean the names first (link the existing clean-customer-list post); the SUMPRODUCT comparison uses = so * and ? in a name are not wildcards (tested); a row with an amount but no name, or a name but no numeric amount, makes CHECK show Check; the formulas are recalculated over 300 rows and each row scans the whole list (300 x 300), fine at this size, not tested beyond it; there is no date column, so this counts orders ever listed, not a period. Not tested: Excel, Google Sheets.

---

## 5. How to Set a Wholesale Price From the Retail Price and Check Your Margin in Excel and Google Sheets
Hook: Get the wholesale price with `=ROUND(C7*$B$2,2)` and your margin at that price with `=E7/D7`, where E7 is `=ROUND(D7-B7,2)`. At 50% of an $18.00 retail price the wholesale price is $9.00, and a $4.10 unit cost leaves a 54.4% margin (made-up numbers; the percent is yours to choose).

Template: templates/tidy-tabs-wholesale-from-retail-price.xlsx. Sheets: **Prices**, **How to use**. Landscape on Prices.

Layout (Prices): A2 "Wholesale price as a percent of retail", **B2** 50% (yellow). A3 "Your target margin on wholesale sales", **B3** 40% (yellow). A4 settings check, **B4**. Header row 6, input rows 7 to 66 (60 products): A product, B your unit cost, C retail price (yellow); D wholesale price, E profit per unit at wholesale, F margin at wholesale, G margin at retail, H shop's margin at retail, I status (OK, Below target, Loss, Check), J wholesale price needed for target, K retail price needed for target (gray). Summary M6:N12: N6 products listed, N7 at or above target, N8 below target, N9 selling at a loss, N10 rows showing Check, N11 lowest margin at wholesale, N12 CHECK.

Key formulas (row 7):
- B4 `=IF(AND(ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B2>0,B2<=1,B3>=0,B3<1),"OK","Check"),"Check")`
- D7 `=IF(AND(B7="",C7=""),"",IF(AND(ISNUMBER(B7),ISNUMBER(C7),$B$4="OK"),IF(AND(B7>=0,C7>0),IF(ROUND(C7*$B$2,2)>0,ROUND(C7*$B$2,2),"Check"),"Check"),"Check"))`
- E7 `=IF(ISNUMBER(D7),ROUND(D7-B7,2),D7)`; F7 `=IF(ISNUMBER(E7),E7/D7,E7)`; G7 `=IF(ISNUMBER(D7),(C7-B7)/C7,D7)`; H7 `=IF(ISNUMBER(D7),(C7-D7)/C7,D7)`
- I7 `=IF(ISNUMBER(F7),IF(E7<0,"Loss",IF(ROUND(F7,6)<ROUND($B$3,6),"Below target","OK")),F7)`
- J7 `=IF(ISNUMBER(D7),ROUNDUP(ROUND(B7/(1-$B$3),6),2),D7)`; K7 `=IF(ISNUMBER(J7),ROUNDUP(ROUND(J7/$B$2,6),2),J7)`
- N12 `=IF(AND(B4="OK",N10=0,SUM(N7:N10)=N6),"OK","Check")`

Sample (50%, target 40%, 11 products, rows 7 to 17): N6 11, N7 6, N8 4, N9 1, N10 0, N11 -10.0%, N12 OK.
- Row 7 Soy candle, Lavender, 8 oz: cost $4.10, retail $18.00, D $9.00, E $4.90, F 54.4%, G 77.2%, H 50.0%, OK, J $6.84, K $13.68.
- Row 13 Enamel mug, Blue, 12 oz: cost $7.80, retail $16.00, D $8.00, E $0.20, F 2.5%, Below target, J $13.00, K $26.00.
- Row 17 Bath bomb gift box: cost $11.00, retail $19.99, D $10.00, E -$1.00, F -10.0%, Loss, J $18.34, K $36.68.

Caveats the post must state: margin is profit divided by the price, markup is profit divided by cost (link the markup-vs-margin post); the 50% and the 40% are example inputs, the post must not say what a wholesale percent or margin should be; everything is arithmetic on fictional costs, it ignores shipping, fees, minimum orders and taxes; the wholesale price is rounded to the cent before the margin is computed, and the status compares margins rounded to 6 decimals so a margin exactly on target counts as OK; a percent typed as 50 (not 50% or 0.5), a retail price of 0, a negative cost or text shows Check; J and K are the lowest prices in whole cents that reach the target under that rounding. Not tested: Excel, Google Sheets.
