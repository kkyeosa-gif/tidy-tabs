---
title: How to Apply a Client Payment to the Oldest Invoices First in Excel and Google Sheets
labels: excel-formulas, small-business, invoicing, freelancers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 30 workbooks and 18,300 comparisons with 0 mismatches. Checked: payments of $0.00, $375.00, $374.99, $375.01, $1,625.00, $5,757.50, $5,757.49, $6,000.00 and $100,000.00, a payment that is blank, text, negative or 2400.005, two invoices on one date, a $0.00 invoice, text dates and amounts, negative amounts, blank cells, a date with a time, amounts of 12.345, and 8 random lists of up to 200 invoices. Excel and Google Sheets were not used. Not run: credits, refunds, more rows than the template holds, a non-US language setting.
search_description: Spread one client payment over open invoices by invoice date with SUMIFS, MIN and MAX so the oldest is paid first. Free template for Excel and Sheets.
image_prompts: A single paper check lying on a short stack of printed invoices in a wire desk tray, the bottom two invoices pulled out and clipped together with a black binder clip, small home office desk, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Paper check resting on printed invoices in a wire desk tray, two clipped together
threads: One check, seven open invoices, not a word about which ones it covers.\nPut the payment in B1. For each invoice, SUMIFS adds up everything dated earlier, and MIN and MAX keep its share between $0.00 and the invoice amount.\nMade-up sample: $2,400.00 pays two invoices in full and $775.00 of a third, even with the oldest invoice typed on the last row.
---
Apply a payment oldest first by giving each invoice the payment minus everything older, capped at the invoice amount and never below zero: `=ROUND(MAX(0,MIN(C5,$B$1-(SUMIFS($C$5:$C$204,$B$5:$B$204,"<"&B5,$C$5:$C$204,">=0")+SUMIFS($C$4:C4,$B$4:B4,B5,$C$4:C4,">=0")))),2)`, where **B1** holds the payment, column **B** the invoice dates and column **C** the open amounts. In the made-up sample, a $2,400.00 payment pays two invoices in full and $775.00 of a third, even though the oldest invoice is typed on the last row. The template adds a Check for text dates, negative amounts and fractions of a cent.

Oldest first is one way to apply a payment, not a rule. If the client's remittance says which invoices a payment covers, follow the remittance. This post is not accounting or legal advice.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (30 workbooks, 18,300 comparisons). The formulas use only IF, AND, OR, NOT, ISNUMBER, TRIM, SUMIFS, SUM, COUNTIF, MIN, MAX and ROUND, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Seven sample invoices showing applied amounts, balances and paid, partly paid or unpaid status](images/2026-11-12-team-02-apply-payment-to-oldest-invoices-first-sumifs-min-excel-google-sheets/apply-payment-oldest-invoices-first-template.png) *A $2,400.00 payment applied by invoice date: two invoices paid in full, one partly paid, four unpaid. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has an **Invoices** tab and a **How to use** tab.

1. In **B1**, type the payment, for example `2400`.
2. List the client's open invoices from row 5 down: invoice number in **A**, invoice date in **B**, open amount in **C**. The template holds rows 5 to 204, in any order.
3. In **D5**, type the formula from the top of this page.
4. In **E5**, type `=ROUND(C5-D5,2)` for the balance after the payment.
5. Fill **D5:E5** down to the last invoice.

How it works: the first `SUMIFS` adds the open amounts of every invoice with an earlier date than this one. The second `SUMIFS` starts at the header row and grows as you fill down (`$C$4:C4`, then `$C$4:C5`), so it adds invoices with the same date on the rows above. Subtract both from the payment and you have what is left when this invoice's turn comes.

`MIN(C5,...)` stops an invoice from getting more than it owes. `MAX(0,...)` stops it from going below zero once the payment runs out. `ROUND(...,2)` keeps the result in whole cents.

Because the order comes from the invoice date, not the row, you do not need to sort the list. Two invoices on the same date are paid top to bottom.

You may see a shorter running total version, `=MIN(C5,MAX(0,$B$1-SUM($C$4:C4)))`. It only works when the rows are sorted oldest first. See Troubleshooting for what it does to an unsorted list.

In the template, column **D** wraps the core formula so blank rows stay blank and bad entries show Check. Column **F** shows Paid in full, Partly paid, Unpaid, Nothing due (open amount $0.00) or Check.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [MIN function](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152). They describe the functions, not our test. The two open questions for Excel and Google Sheets are the `"<"&B5` date criterion and the growing `$B$4:B4` range; both were checked in LibreOffice Calc only.

## Example

Sample data is fictional: seven invoices for a made-up client, Harbor Dental, with a $2,400.00 partial payment. The oldest invoice, INV-1039, is typed last.

| Invoice no. | Invoice date | Open amount | Applied | Balance after | Status |
|---|---|---|---|---|---|
| INV-1041 | 08/15/2026 | $1,250.00 | $1,250.00 | $0.00 | Paid in full |
| INV-1047 | 09/01/2026 | $840.00 | $775.00 | $65.00 | Partly paid |
| INV-1052 | 09/15/2026 | $1,250.00 | $0.00 | $1,250.00 | Unpaid |
| INV-1058 | 10/01/2026 | $612.50 | $0.00 | $612.50 | Unpaid |
| INV-1063 | 10/15/2026 | $1,250.00 | $0.00 | $1,250.00 | Unpaid |
| INV-1064 | 10/15/2026 | $180.00 | $0.00 | $180.00 | Unpaid |
| INV-1039 | 08/01/2026 | $375.00 | $375.00 | $0.00 | Paid in full |

The payment goes to INV-1039 ($375.00), then INV-1041 ($1,250.00), then INV-1047 ($775.00). That is $375.00 + $1,250.00 + $775.00 = $2,400.00, so nothing reaches the October invoices. INV-1063 and INV-1064 share a date, so INV-1063 would be paid first.

The summary block shows Payment received $2,400.00, Open before the payment $5,757.50, Applied to invoices $2,400.00, Left over $0.00, Still open after $3,357.50, 2 paid in full, 1 partly paid, 4 unpaid, 0 rows showing Check and a CHECK cell that says OK.

## Troubleshooting

### The oldest invoice gets nothing although it is the oldest
You are using the sorted-only running total formula on an unsorted list. On the sample, `=MIN(C5,MAX(0,$B$1-SUM($C$4:C4)))` paid INV-1041 $1,250.00, INV-1047 $840.00 and INV-1052 $310.00, and left INV-1039 at $0.00. The SUMIFS version reads the dates and pays INV-1039 first.

### A row shows Check
The date is text (for example "10/20/2026" stored as a text string), the amount is text or negative, or the amount has fractions of a cent (99.999 and 12.345 both show Check). Retype the date as a real date and the amount as a number in whole cents. Fix every Check row first, because until then the other rows can be off.

### CHECK shows Check although every row looks right
The payment in **B1** is blank, text, negative or has fractions of a cent. Blank, "abc", "2400" typed as text, -1 and 2400.005 all gave Check. Type the payment as a plain number such as `2400`.

### Left over is more than $0.00
The payment is bigger than everything open. With a payment of $7,000.00, all seven sample invoices show Paid in full and Left over is $1,242.50. The file does not decide what happens to that money (refund, credit or a future invoice); that is between you and the client.

## Template

[Download the .xlsx](templates/tidy-tabs-apply-payment-oldest-first.xlsx)

The file has two tabs. **Invoices** has the payment in **B1**, the client name in **B2**, invoices in **A5:C204**, the formula columns **D:F**, and a summary with a CHECK cell in **H4:I13**. **How to use** explains each column. Credits and partial refunds are not handled; enter the open amount after them.

Yellow cells are inputs and gray cells are formulas. Paid in full is green and Check is red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
