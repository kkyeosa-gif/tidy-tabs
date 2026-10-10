# Handoff for posts dated 2026-10-11

Status: topics (researcher), titles and openings (hook-writer) and templates (template-builder) done. Templates are built by `scripts/build-templates.py` and were recalculated in LibreOffice Calc 24.2 (Linux, headless: `soffice --headless --convert-to xlsx` into a temp dir, read back with openpyxl `data_only`) on 10/10/2026. Every result below was compared with hand or Python math. Edge cases listed were run by editing a copy of the template with openpyxl, recalculating it in LibreOffice and reading the values. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. No tasks/search-stats.md exists: no search data. Sample data is fictional (say so once per post). Each template tab prints on one US Letter page wide; the PDF export from LibreOffice gave one page per sheet.

Titles carry "Excel and Google Sheets", name the task and function, and have no colon, dash or hyphen. Openings give the exact formula first. Cell addresses below are the real ones in the built files.

---

## 1. 2026-10-11-team-01-quarterly-sales-totals-sumifs-excel-google-sheets
- Title: How to Total Sales by Quarter in Excel and Google Sheets
- Opening: Turn each date into a quarter label with `=YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0)`, then total a quarter with `=SUMIFS($B$2:$B$200,$C$2:$C$200,"2026 Q2")`. The year in the label keeps 2025 Q4 and 2026 Q4 apart.
- Template: `templates/tidy-tabs-quarterly-sales-totals.xlsx`. Sheets: Sales, By quarter, How to use.
- Formulas:
  - Sales!C2: `=YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0)` (filled to C17)
  - By quarter!A2: `=YEAR(B2)&" Q"&ROUNDUP(MONTH(B2)/3,0)`; C2 `=SUMIFS(Sales!$B$2:$B$200,Sales!$C$2:$C$200,A2)`; D2 `=COUNTIFS(Sales!$C$2:$C$200,A2)`
  - By quarter!E2 (no helper column): `=SUMIFS(Sales!$B$2:$B$200,Sales!$A$2:$A$200,">="&B2,Sales!$A$2:$A$200,"<"&EDATE(B2,3))`; F2 `=IF(ROUND(C2-E2,2)=0,"Yes","CHECK")`
- Verified values (LibreOffice Calc 24.2 only):
  - Sales A1:C17: 16 fictional sales, dates in A (real dates), amounts in B, 12/20/2025 to 12/31/2026.
  - By quarter C2:C6 (and E2:E6, same): 2025 Q4 $90.00 (1 order), 2026 Q1 $405.50 (3), 2026 Q2 $630.25 (4), 2026 Q3 $600.50 (4), 2026 Q4 $779.99 (4). F2:F6 all Yes. C7 total $2,506.24, D7 16, C8 (SUM of Sales tab) $2,506.24. All match Python.
  - Boundary days: 03/31/2026 lands in 2026 Q1 (Sales!C5), 04/01/2026 in 2026 Q2 (C6), 09/30 in Q3, 10/01 in Q4, 12/31/2026 in Q4.
  - Edge: a date typed as text ("05/09/2026" and "May 9, 2026" in Sales!A8): LibreOffice still converted it inside YEAR/MONTH, so C8 shows "2026 Q2" and the helper total (C4) stays $630.25, but the date-range total E4 drops to $320.25 and F4 shows CHECK. So a text date can make the two methods disagree. How Excel or Sheets treat text dates was not tested.
- Limits: calendar quarters only (January to March is Q1). Fiscal or tax-filing quarters are not covered and no tax advice is given. Ranges stop at row 200.

---

## 2. 2026-10-11-team-02-find-missing-invoice-numbers-countif-excel-google-sheets
- Title: How to Find Missing Invoice Numbers in a Sequence in Excel and Google Sheets
- Opening: List every number you should have, then flag the gaps with `=IF(COUNTIF($A$2:$A$200,D2)=0,"Missing","OK")`, where D2 is one expected number and A2:A200 holds the invoice numbers you actually issued. Any number that returns Missing is a gap to look into.
- Template: `templates/tidy-tabs-missing-invoice-numbers.xlsx`. Sheets: Invoices, Check, How to use. Note the opening uses plain refs; in the template the formulas read `Invoices!$A$2:$A$200` on the Check tab.
- Formulas (Check tab):
  - A2: `=MIN(Invoices!$A$2:$A$200)+ROW()-2` (expected number, A2:A13)
  - B2: `=COUNTIF(Invoices!$A$2:$A$200,A2)`
  - C2: `=IF(B2=0,"Missing",IF(B2>1,"Duplicate","OK"))` (red fill for Missing, yellow for Duplicate)
  - Summary: F2 `=MIN(Invoices!A2:A200)`, F3 `=MAX(...)`, F4 `=F3-F2+1`, F5 `=COUNT(...)`, F6 `=COUNTIF(C2:C13,"Missing")`, F7 `=COUNTIF(C2:C13,"Duplicate")`, F8 `=F5-(SUMPRODUCT((B2:B13>1)*(B2:B13-1)))+F6`
- Verified values (LibreOffice Calc 24.2 only):
  - Invoices A2:A11: 1001, 1002, 1003, 1005, 1006, 1008, 1009, 1009, 1011, 1012 (10 rows).
  - Check: A2:A13 = 1001 to 1012. Missing: 1004 (C5), 1007 (C8), 1010 (C11). Duplicate: 1009 (B10 = 2, C10). Others OK. F2 1001, F3 1012, F4 12, F5 10, F6 3, F7 1, F8 12 (equals F4). Matches hand check.
  - Edge: invoice number typed as text "1005" in Invoices!A5: COUNTIF did not match it (B6 = 0, C6 Missing, F5 drops to 9). Number with a prefix "INV-1005": same result, and MIN/MAX ignore it. So text numbers show as false gaps; keep the number in its own numeric column.
  - Edge: a number far above the list (1020 in Invoices!A12): F3 1020, F4 20, F6 3, F8 13, which no longer equals F4, so the mismatch shows the Check tab needs more rows. A number below the list (1000) moves the start to 1000 and pushes 1012 off the 12 rows.
- Limits: only finds gaps between the lowest and highest number. Does not say why a number is missing (a voided invoice is possible); no accounting or legal conclusion. Check tab lists 12 numbers; fill down for longer runs.

---

## 3. 2026-10-11-team-03-sum-if-cell-contains-text-wildcard-sumif-excel-google-sheets
- Title: How to Sum Cells That Contain Certain Text With SUMIF Wildcards in Excel and Google Sheets
- Opening: Wrap the word in asterisks with `=SUMIF(B2:B13,"*"&E1&"*",C2:C13)` to total every amount whose description contains the word in E1. The `*` stands for any text, and the match ignores capital letters.
- Template: `templates/tidy-tabs-sumif-contains-text.xlsx`. Sheets: Expenses, Totals, How to use. In the template the search words are Totals!A2:A7 and ranges read `Expenses!$B$2:$B$200` and `Expenses!$C$2:$C$200`.
- Formulas (Totals tab, row 2 shown):
  - B2 contains: `=SUMIF(Expenses!$B$2:$B$200,"*"&A2&"*",Expenses!$C$2:$C$200)`
  - C2 count: `=COUNTIF(Expenses!$B$2:$B$200,"*"&A2&"*")`
  - D2 starts with: `=SUMIF(Expenses!$B$2:$B$200,A2&"*",Expenses!$C$2:$C$200)`; E2 ends with: `=SUMIF(Expenses!$B$2:$B$200,"*"&A2,Expenses!$C$2:$C$200)`
- Verified values (LibreOffice Calc 24.2 only; wildcards worked in the xlsx as built):
  - Expenses A1:C13: 12 fictional September expenses. Total B9 $216.59.
  - "Uber" (row 2): contains $68.75, count 3 (18.40 + 14.25 + 36.10, including "UBER EATS lunch", so case is ignored), starts with $68.75, ends with $0.00.
  - "Etsy" (row 3): $31.00, count 2. "ink" (row 4): contains $78.24, count 4 (Printer ink 42.99, Pink ribbon 6.50, Packing tape and ink pen 9.75, Ink refill kit 19.00), starts with $19.00. "Ink pen" (row 5): $9.75, count 1, ends with $9.75. "Stamps~*" (row 6): $11.60, count 1 (the ~ makes the * literal; the description is "Stamps*"). "Zoom" (row 7): $0.00, count 0. All match hand math.
  - Edge: the substring trap is real: "ink" also matched "Pink ribbon spool".
  - Edge: blank search cell (Totals!A7 emptied): pattern becomes `**`, B7 = $216.59 and C7 = 12, every row with text. "Ub?r" matched the 3 Uber rows ($68.75), so `?` works as one character.
  - Edge: a number typed in the description column (B2 = 2024) is not matched by a `*` pattern (Uber total drops to $50.35, count 2). An amount typed as text ("18.40" in C2): COUNTIF still counted the row (3) but SUMIF left it out of the total ($50.35).
- Limits: letters, not whole words. Pattern characters `*`, `?`, `~` in the search word need `~`. Behavior in Excel and Google Sheets was not tested; link their SUMIF help pages for that.

---

## 4. 2026-10-11-team-04-split-invoice-into-installments-excel-google-sheets
- Title: How to Split an Invoice Into Equal Monthly Payments in Excel and Google Sheets
- Opening: Divide with `=ROUND(B1/B2,2)` for the regular payment, then make the last payment `=B1-B4*(B2-1)` so the cents left over land there and the payments add up to the invoice exactly. A $2,500.00 invoice in 6 payments is five payments of $416.67 and a last one of $416.65.
- Template: `templates/tidy-tabs-invoice-installments.xlsx`. Sheets: Payment plan, How to use.
- Inputs (yellow): B1 invoice total, B2 number of payments (data validation whole number 1 to 12), B3 first payment date.
- Formulas:
  - B4: `=ROUND(B1/B2,2)`; B5: `=B1-B4*(B2-1)`
  - Rows 8 to 19 (payment #1 to 12): B `=IF(A8>$B$2,"",EDATE($B$3,A8-1))`; C `=IF(A8>$B$2,"",IF(A8=$B$2,$B$5,$B$4))`; D `=IF(A8>$B$2,"",$B$1-SUM($C$8:C8))`
  - C21 `=SUM(C8:C19)`; C22 `=ROUND(C21-B1,2)`
- Verified values (LibreOffice Calc 24.2 only):
  - Default: $2,500.00, 6 payments, first date 11/01/2026: B4 $416.67, B5 $416.65; C8:C12 $416.67, C13 $416.65; due dates 11/01/2026, 12/01/2026, 01/01/2027, 02/01/2027, 03/01/2027, 04/01/2027; D13 $0.00; C21 $2,500.00; C22 $0.00. Matches Python.
  - $1,000.00 in 3 payments: $333.33, $333.33, $333.34, total $1,000.00. $1,000.00 in 12: $83.33 x 11 and a last $83.37, balance ends $0.00. $100.00 in 7: $14.29 x 6 and $14.26. $750.00 in 1: one payment of $750.00, balance $0.00.
  - Month end: first date 01/31/2026, 4 payments: 01/31/2026, 02/28/2026, 03/31/2026, 04/30/2026.
  - Edge: typing 13 payments into B2 (done in a copy by script, which bypasses the dropdown validation rule) breaks the plan: C21 $1,200.00 and C22 -$100.00 because only 12 rows exist. The validation rule is in the file but was not tried by typing in the LibreOffice interface.
- Limits: arithmetic only; no interest, fees or late charges; follow the written agreement with the client. Rows past the payment count show blank (the # column still shows 7 to 12).

---

## 5. 2026-10-11-team-05-marketplace-fee-net-payout-calculator-excel-google-sheets
- Title: How to Calculate Net Payout After Marketplace Fees in Excel and Google Sheets
- Opening: Subtract the fees from the price with `=B2-ROUND(B2*($H$2+$H$3)+$H$4,2)`, where H2 and H3 are the percentage fees and H4 is the fixed fee per order. The sample uses made up rates; type the rates from your own seller account.
- Template: `templates/tidy-tabs-marketplace-net-payout.xlsx`. Sheets: Net payout, How to use (landscape).
- Inputs (yellow): H2 marketplace fee 6.5%, H3 payment processing 3.0%, H4 fixed fee $0.25 (all fictional), H7 target net payout $15.00.
- Formulas:
  - C2 fees: `=ROUND(B2*($H$2+$H$3)+$H$4,2)`; D2 net: `=B2-C2`; E2 `=D2/B2`
  - H8 price for target: `=ROUNDUP((H7+H4)/(1-H2-H3),2)`; H9 check: `=H8-ROUND(H8*(H2+H3)+H4,2)`
- Verified values (LibreOffice Calc 24.2 only):
  - Orders A2:B7 and results (price, fees, net, net %): Soy candle $18.00, $1.96, $16.04, 89.1%; Cedar soap bar $8.50, $1.06, $7.44, 87.5%; Brass hoop earrings $24.00, $2.53, $21.47, 89.5%; Gift tag set $1.00, $0.35, $0.65, 65.0%; Candle gift box $60.00, $5.95, $54.05, 90.1%; Sticker sample $0.25, $0.27, -$0.02, -8.0%. Totals row 8: $111.75, $12.12, $99.63, 89.2%. All match Python (fee rounding half up at cents; $8.50 gives 0.8075 + 0.25 = 1.0575, shown $1.06).
  - Reverse: H8 = $16.86 for a $15.00 target; H9 = $15.01. One cent above target because fees round to cents (a $16.85 price would net exactly $15.00 by hand). The notes tab says this. $20.00 target gives $22.38 and a check of $20.00.
  - Edge: fixed fee $0 gives sticker fee $0.02, net $0.23, and H9 exactly $15.00. Percent fees adding to 100% (set both to 50%) make H8 and H9 show #DIV/0!.
- Limits: only the fees typed in. No shipping, sales tax collected, ads, refunds or income tax. Rates are fictional and marketplaces change fees, so the reader must use current rates from their own account.
