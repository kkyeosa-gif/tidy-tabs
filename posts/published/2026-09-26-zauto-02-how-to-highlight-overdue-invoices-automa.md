---
threads_url: https://www.threads.com/@tin_ylab/post/DdwT-bSlGUh
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-highlight-overdue-invoices.html
title: How to Highlight Overdue Invoices Automatically in Excel and Google Sheets
labels: conditional-formatting, invoicing
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small desk with a laptop showing a spreadsheet with colored rows, a stack of paper invoices, and a wall calendar nearby, all screens and text unreadable
image_alt: Desk with laptop, printed invoices, and calendar showing overdue billing tracking
search_description: Set up one conditional formatting rule in Excel or Google Sheets that turns invoice rows red automatically when they pass their due date and are still unpaid.
threads: Stop scrolling your invoice tracker looking for who owes you money.\nOne conditional formatting rule turns a row red the day it's overdue and still marked unpaid: =AND($D2<>"Paid",$C2<TODAY())\nWorks the same in Excel and Google Sheets.
---
Highlight overdue invoices by adding a conditional formatting rule with the formula `=AND($D2<>"Paid", $C2<TODAY())`, pointed at your Due Date and Status columns. Rows turn red automatically the day they pass due, and turn back to normal the moment you mark them Paid.

> Works in: Excel (Microsoft 365 and Excel 2019+) and Google Sheets. Menu paths below come from Microsoft and Google's own help pages.

## Steps

Set up your sheet with at least these columns: Invoice #, Client, Due Date, Status, Amount. The formula below assumes Due Date is column C and Status is column D, both starting on row 2.

### In Excel

1. Select the data range you want to format, for example `A2:E100`.
2. Go to **Home > Conditional Formatting > New Rule**.
3. Choose **Use a formula to determine which cells to format**.
4. Enter `=AND($D2<>"Paid", $C2<TODAY())`.
5. Click **Format**, pick a red fill, then click **OK** twice.

### In Google Sheets

1. Select the data range, for example `A2:E100`.
2. Go to **Format > Conditional formatting**.
3. In the **Format cells if...** dropdown, choose **Custom formula is**.
4. Enter `=AND($D2<>"Paid", $C2<TODAY())`.
5. Pick a red fill under **Formatting style**, then click **Done**.

Both apps apply the rule to every row in the range, checking each row's own Due Date and Status.

## Example

Sample data below is fictional.

| Invoice # | Client | Due Date | Status | Amount | Row color |
|---|---|---|---|---|---|
| 1001 | Maple Street Bakery | 09/10/2026 | Unpaid | $450.00 | Red (overdue) |
| 1002 | Riverside Yoga Studio | 09/28/2026 | Unpaid | $200.00 | Normal (not due yet) |
| 1003 | Oak & Co. Woodworks | 09/05/2026 | Paid | $780.00 | Normal (paid) |
| 1004 | Bloom Farmers Market | 09/15/2026 | Unpaid | $125.00 | Red (overdue) |

## Troubleshooting

### Every row turns red, even ones that aren't overdue
Your formula probably uses absolute references like `$D$2` instead of `$D2`. Locking the row number stops the rule from checking each row against its own data.

### Rows don't turn red on the due date itself
The formula uses `<` for Due Date, so a row turns red the day *after* it's due, not on the due date. Change `<` to `<=` if you want the row to flag on the due date itself.

### Dates are stored as text, so the rule never triggers
If Due Date shows as text (left-aligned instead of right-aligned), `TODAY()` comparisons fail silently. Reformat the column as a date or convert the values to real dates — in Excel, **Data > Text to Columns** can often do this. In Google Sheets, retyping the dates or reformatting the cells as dates is usually the more reliable fix.

### Paid invoices still show red for a moment after you update Status
Some versions may not recalculate conditional formatting instantly. Clicking any other cell can force a refresh.

### The rule disappears after you copy rows to a new sheet
Conditional formatting rules don't always carry over with a plain copy-paste. Use **Paste special > Paste format only** (Google Sheets, under the Edit menu) or **Paste Special > Formats** (Excel, under Home > Paste) to bring the rule along.

## Copy-paste setup

Headers (row 1): `Invoice #`, `Client`, `Due Date`, `Status`, `Amount`

Conditional formatting rule (applies to `A2:E100`):
- Rule type: Custom formula
- Formula: `=AND($D2<>"Paid", $C2<TODAY())`
- Format: red fill (or any color you want for overdue rows)

Optional second rule for invoices due soon but not yet late:
- Formula: `=AND($D2<>"Paid", $C2>=TODAY(), $C2<=TODAY()+3)`
- Format: yellow fill

Type Status as plain text values (`Unpaid`, `Paid`) so the formula's `<>"Paid"` comparison matches exactly.

More on how conditional formatting rules work: [Add, change, find, or clear conditional formats in Excel](https://support.microsoft.com/en-us/office/add-change-find-or-clear-conditional-formats-99fbd928-7be5-4324-8fda-6e5f92f91385), [TODAY function](https://support.microsoft.com/en-us/office/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9), and [Add conditional formatting rules in Google Sheets](https://support.google.com/docs/answer/78413).
