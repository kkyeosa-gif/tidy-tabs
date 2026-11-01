---
title: How to Write a Friendly Payment Reminder From Your Invoice List in Excel and Google Sheets
labels: excel-formulas, small-business, invoicing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there in 14 edited copies and 2,590 comparisons of exact message text and numbers against an independent Python model, with 0 mismatches. Checked: due dates equal to, one day after and 1, 13, 14, 15 and 16 days before the as-of date, a due date with a time, leap day and year end, amounts from $0.00 to $1,234,567.89, an amount typed as text, blank and padded client names, a due date or as-of date typed as text, and Paid typed yes, Yes, Y and No. TEXT returned the US format with the en-US setting. TODAY() was not used. Excel and Google Sheets were not used.
search_description: Build a payment reminder from invoice data with TEXT in Excel and Google Sheets. Free template with days past due and a tone for each invoice.
image_prompts: A stack of paper invoices held with a binder clip on a wooden desk next to a brass desk calendar and a mug of pens, soft window light, no screens, no readable text, no logos
image_alt: Invoices clipped together on a desk beside a brass calendar and a pen mug
threads: Need to nudge a client about invoice INV-2052?\nJoin the cells into one sentence and wrap the amount in TEXT(C5,"$#,##0.00") so it reads $85.00, not 85.\nThe template picks a gentle or follow-up wording from days past due.
---
Build the message from the cells with `="Hi "&B5&", a quick reminder that invoice "&A5&" for "&TEXT(C5,"$#,##0.00")&" was due on "&TEXT(D5,"mm/dd/yyyy")&"."`. Copy it down and each row gives you a sentence you can paste into an email, with the amount and date formatted the US way.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every message below was recalculated and compared with a Python model. `&` and `TEXT` exist in every current version of Excel and Google Sheets, but neither app was opened for this post.

![Invoice list with days past due, tone and ready-to-paste reminder messages for unpaid invoices](images/2026-11-02-team-05-payment-reminder-message-from-invoice-data-text-excel-google-sheets/payment-reminder-message-list.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. Put one invoice per row, starting in row 5: invoice number in A, client first name in B, amount in C, due date in D.

1. Type the invoice details in **A5:D5**, such as INV-2051, Maya, $1,250.50 and 10/05/2026. Enter the date as a real date, not as text.
2. In **E5**, type `Yes` when the invoice is paid. Leave it blank when it is not.
3. In any empty cell, type the formula from the top of this post.
4. Copy the cell down, one row per invoice.
5. Copy the message, paste it into an email, and change anything you want before sending.

Run on that row, the formula gives:

`Hi Maya, a quick reminder that invoice INV-2051 for $1,250.50 was due on 10/05/2026.`

The `&` joins pieces of text. `TEXT` turns a number or date into text in the format you choose, so the amount keeps its dollar sign and cents, and the date keeps its slashes. Without `TEXT`, a date joins in as its stored number.

The `TEXT` format codes follow your app's language setting. They returned the US format in the test, which used the en-US setting. They were not run in another language. If your codes do not work, format the cells with **Format > Number** and join the cells without `TEXT`, or check your language setting.

The template goes further than that one formula. It counts days past due, picks a tone, and writes a different sentence for each:

- **F5**, days past due: `=IF(OR(A5="",TRIM(E5)="Yes"),"",IF(AND(ISNUMBER(D5),ISNUMBER(Settings!$B$1)),INT(Settings!$B$1)-INT(D5),"Check date"))`
- **G5**, tone: `=IF(F5="","",IF(OR(NOT(ISNUMBER(F5)),NOT(ISNUMBER(C5))),"Check",IF(F5<=0,"Upcoming",IF(F5<=14,"Gentle","Follow-up"))))`

The message in **H5** is the basic formula plus guards. A blank client name becomes "there", "1 day ago" is not "1 days ago", and a bad amount or date shows a plain note instead of a broken sentence. Your name comes from **Settings**.

The as-of date is typed in **Settings B1**, not set with `TODAY()`. That way a message does not change when you open the file next week.

Help pages: Microsoft's [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test.

## Example

Sample data is fictional: made up clients, invoice numbers and amounts. The as-of date is 11/02/2026 and the sign-off name is Jordan.

| Invoice | Days past due | Tone | Message |
|---|---|---|---|
| INV-2052 | 45 | Follow-up | Hi Devon, a quick reminder that invoice INV-2052 for $85.00 was due on 09/18/2026 (45 days ago). Let me know if you need another copy. Thanks, Jordan |
| INV-2053 | 3 | Gentle | Hi Priya, a quick reminder that invoice INV-2053 for $640.00 was due on 10/30/2026 (3 days ago). Let me know if you need another copy. Thanks, Jordan |
| INV-2055 | 0 | Upcoming | Hi Alicia, just a heads-up that invoice INV-2055 for $1,200.00 is due on 11/02/2026. Thanks, Jordan |

Of the 12 sample invoices, 3 are paid and get no message. Gentle means 1 to 14 days past due, and Follow-up means 15 or more. A due date of today or later is Upcoming. In the test, 14 days past due was Gentle and 15 was Follow-up.

The summary on the right says 5 messages to send, with $2,155.00 past due.

The wording has no late fee, deadline or legal text. This is a spreadsheet exercise, not collections advice.

## Troubleshooting

### A row shows Check date
The due date is blank or typed as text, such as `10/01/2026` stored as plain text. Retype it as a real date. In the test, the row showed Check date and the **CHECK** cell turned to Check.

### Every unpaid row shows Check date
The as-of date in **Settings B1** is blank or text. Type a date like 11/02/2026. **Settings C1** shows a message when it is wrong.

### The message says Check the amount and due date in this row
The amount is text or blank. In the test, an amount of `100` typed as text gave this note instead of a sentence. Retype it as a number.

### A paid invoice still gets a message
Only `Yes` counts as paid. The test treated `yes` and `Yes ` with a trailing space as paid, but `Y` and `No` were not. Type `Yes`.

### The message ends without Thanks
The name in **Settings B2** is blank. Messages then end after "if you need another copy."

## Template

[Download the .xlsx](templates/tidy-tabs-payment-reminders.xlsx)

The file has three tabs. **Invoices** has 12 fictional rows in yellow cells **A5:E64** (60 slots), with days past due, tone and message columns, and a summary in **J4:K6**. **Settings** holds the as-of date in **B1** and your name in **B2**. **How to use** has notes on each formula.

The **CHECK** cell, **K6**, says OK, or Check if the as-of date is not a date, any row says Check, or a row with a day count has no message. Fix every Check before you send anything.

Not tested: Excel, Google Sheets and non-US settings. **Invoices** prints landscape on US Letter. To open the file in Google Sheets, go to **File > Import > Upload** and select it.
