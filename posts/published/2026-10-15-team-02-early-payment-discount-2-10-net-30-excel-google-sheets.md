---
threads_url: https://www.threads.com/@tin_ylab/post/Dehxa5vjzft
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-210-net-30-early.html
title: How to Calculate a 2/10 Net 30 Early Payment Discount in Excel and Google Sheets
labels: excel-formulas, small-business, invoicing
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python, including a payment on the deadline day, one day late, an amount typed as text, different terms and an empty sheet. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: 2/10 net 30 discount: =ROUND(D2*$M$2,2). A $1,250.00 invoice saves $25.00 if paid by 09/11/2026. Free .xlsx with deadlines and status.
image_prompts: Stack of paper invoices with a metal binder clip, a brass desk calendar block and a pen resting on a worn wooden desk, soft side light, no screens, no readable text, no logos
image_alt: Paper invoices held by a binder clip beside a brass calendar block and pen
threads: A 2/10 net 30 invoice means 2% off if paid in 10 days.\n=ROUND(D2*$M$2,2) gives the discount, =C2+$M$3 gives the last day to take it.\nOn $1,250.00 dated 09/01/2026, that is $25.00 off by 09/11/2026.\nPay on 09/11 and you still get it. 09/12 and you do not.
---
Find the discount with `=ROUND(D2*$M$2,2)` and the last day to take it with `=C2+$M$3`, where M2 is 2% and M3 is 10 days. On a $1,250.00 invoice dated 09/01/2026, the client saves $25.00 by paying $1,225.00 on or before 09/11/2026.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Invoice table with discount amounts, deadlines and status labels beside discount totals](images/2026-10-15-team-02-early-payment-discount-2-10-net-30-excel-google-sheets/early-payment-discount-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

"2/10 net 30" means the client gets 2% off if they pay within 10 days. Otherwise the full amount is due in 30 days. Take the three numbers from your own invoice or contract. This post only shows the arithmetic. It is not tax or legal advice, and your contract decides what a client owes.

1. In row 1, type these headers: Invoice, Client, Invoice date, Amount, Paid date, Discount deadline, Discount amount, Pay by deadline, Full amount due date, Status.
2. Type each invoice date as MM/DD/YYYY in column **C**. Type the amount as a plain number in column **D**, with no text and no commas typed as characters. Leave **E** empty until the client pays.
3. In **M2**, type `2%`. In **M3**, type `10`. In **M4**, type `30`. These are the terms.
4. In **F2**, type `=IF(C2="","",C2+$M$3)`. This is the last day to take the discount. Format the column as a date.
5. In **G2**, type `=IF(OR(C2="",NOT(ISNUMBER(D2))),"",ROUND(D2*$M$2,2))`. ROUND keeps the discount to whole cents.
6. In **H2**, type `=IF(G2="","",D2-G2)`. This is what the client pays if they hit the deadline.
7. In **I2**, type `=IF(C2="","",C2+$M$4)`. This is the full due date.
8. In **J2**, type `=IF(OR(C2="",G2=""),"",IF(E2="","Open",IF(E2<=F2,"Discount earned","Full amount")))`.
9. Select **F2:J2** and drag the fill handle down as far as you have invoices.
10. Add totals. In **M6**, type `=SUMIF(J2:J25,"Discount earned",G2:G25)`. Repeat with "Full amount" in **M7** and "Open" in **M8**. In **M9**, type `=COUNTIF(J2:J25,"Open")`.

The formulas stop at row 25. Extend them if you have more invoices.

### In Excel

The formulas work as typed. Microsoft describes the rounding on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) page and the totals on its [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) page. Format the date columns from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Format the date columns with **Format > Number > Date**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. The terms are 2% in 10 days, net 30.

| Invoice | Invoice date | Amount | Paid date | Deadline | Discount | Pay by deadline | Status |
|---|---|---|---|---|---|---|---|
| INV-3001 | 09/01/2026 | $1,250.00 | 09/09/2026 | 09/11/2026 | $25.00 | $1,225.00 | Discount earned |
| INV-3002 | 09/03/2026 | $480.00 | 09/20/2026 | 09/13/2026 | $9.60 | $470.40 | Full amount |
| INV-3003 | 09/08/2026 | $2,200.00 | 09/18/2026 | 09/18/2026 | $44.00 | $2,156.00 | Discount earned |
| INV-3008 | 10/09/2026 | $725.40 | (blank) | 10/19/2026 | $14.51 | $710.89 | Open |

INV-3003 was paid on the deadline day, and that still earns the discount. The test is "paid date is on or before the deadline".

Across all eight sample invoices, the discounts earned total $88.20, the discounts missed total $26.31, and the three open invoices carry $50.36 in possible discounts. INV-3008 shows how rounding works: 2% of $725.40 is $14.508, which ROUND turns into $14.51.

## Troubleshooting

### A payment one day late still shows Discount earned
Check the paid date. In the recalculated copy, INV-3001 paid on 09/11/2026 showed Discount earned. Paid on 09/12/2026 it switched to Full amount, and the missed total rose to $51.31.

### The discount cell is blank
The amount is probably text. In the recalculated copy, "1,250.00" typed as text left the Discount, Pay by deadline and Status cells blank, so the row was ignored. Retype it as a plain number.

### The deadline does not match my terms
The deadline reads M3 and the full due date reads M4. Change them there. With 1%, 15 days and net 45, INV-3001 got a 09/16/2026 deadline, a $12.50 discount, $1,237.50 to pay and a 10/16/2026 due date.

### An open invoice still says Open after the deadline
That is a known limit. The sheet does not compare the deadline with today's date, so an unpaid invoice stays Open. Compare the Discount deadline column with the calendar yourself.

### The totals show $0.00
If every row is empty, M6 through M9 show 0. Add an invoice date and a numeric amount.

## Limits

This template counts calendar days, not business days. It handles terms of the form discount percent, discount days and net days, and one set of terms applies to the whole sheet. It does not handle end of month terms.

## Template

[Download the .xlsx](templates/tidy-tabs-early-payment-discount.xlsx)

The file has two tabs. **Discounts** holds 8 sample invoices, 3 of them unpaid, with room for 25 rows, plus the yellow terms cells and the totals. **How to use** has short fill-in notes. The Discounts tab prints on one US Letter page.

The names, dates and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
