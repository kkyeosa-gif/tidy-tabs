---
threads_url: https://www.threads.com/@tin_ylab/post/DeMU79DGx5Y
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-auto-number-invoices-in-excel.html
title: How to Auto Number Invoices in Excel and Google Sheets
labels: formulas, invoices, freelancers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps follow Microsoft Support (Google Sheets not tested here).
search_description: Auto number invoices like INV-2026-0007 with TEXT and ROWS. The year follows the invoice date. Free .xlsx template with six fictional invoices.
image_prompts: A stack of blank carbon-copy invoice pads with a rubber number stamp and an ink pad on a wooden counter, a few paper invoices fanned out with blurred unreadable lines, a ballpoint pen, soft window light, no screens, no logos
image_alt: Invoice pads beside a number stamp and ink pad on a counter
threads: Typing invoice numbers by hand means one day you skip INV-0007.\nLet the sheet count: ="INV-"&TEXT(B2,"yyyy")&"-"&TEXT(ROWS($A$2:A2),"0000").\nThe year comes from the invoice date, and the count pads to four digits.
---
Build numbers like INV-2026-0007 with `="INV-"&TEXT(B2,"yyyy")&"-"&TEXT(ROWS($A$2:A2),"0000")`. The year follows the invoice date in B2, and the count pads with zeros to four digits.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and the numbers below were recalculated and read back. The Excel and Google Sheets steps use the same functions and link to Microsoft Support. Neither Excel nor Google Sheets was opened for this post.

![Invoices sheet listing six clients with invoice dates, generated invoice numbers, and amounts](images/2026-10-06-team-04-invoice-number-generator-text-excel-google-sheets/invoice-number-generator-template.png) *LibreOffice Calc 24.2 PDF export of the Invoices sheet with sample data (fake).*

## Steps

The formula needs no macro and no saved counter. It counts rows, so each new row gets the next number.

1. In row 1, type these headers: Client, Invoice date, Invoice number (TEXT), Invoice number (YEAR), Amount.
2. In column A, type the client. In column B, type the invoice date as MM/DD/YYYY, such as 09/28/2026. It must be a real date, not text.
3. In **C2**, type `="INV-"&TEXT(B2,"yyyy")&"-"&TEXT(ROWS($A$2:A2),"0000")` and press Enter.
4. Drag the fill handle in the corner of C2 down to the last row. Each row now shows the next number.
5. In **D2**, type `="INV-"&YEAR(B2)&"-"&TEXT(ROWS($A$2:A2),"0000")` and fill it down. This version takes the year with YEAR instead of a format code.
6. Format column E as currency.

Here is how the parts work:

- `"INV-"&` adds the prefix.
- `TEXT(B2,"yyyy")` pulls the four-digit year out of the invoice date.
- `ROWS($A$2:A2)` counts the rows in the range. The first cell is locked with dollar signs, so the range grows by one row each time you fill down: 1 in row 2, 2 in row 3, and so on.
- `TEXT(...,"0000")` pads that count to four digits, so 7 becomes 0007.

### In Excel

The formulas work as typed. Microsoft documents the functions on its [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c) and [ROWS function](https://support.microsoft.com/en-us/office/rows-function-b592593e-3fc2-47f2-bec1-bda493811597) pages.

### In Google Sheets

TEXT, ROWS, and YEAR are Sheets functions too, so the same formulas should work as typed. No Google help page is linked here, and this was not tested in Sheets.

## Example

Sample data below is fictional. It is six invoices from a freelancer's client list.

| Client | Invoice date | Invoice number (TEXT) | Invoice number (YEAR) | Amount |
|---|---|---|---|---|
| Harbor Coffee Co. | 09/28/2026 | INV-2026-0001 | INV-2026-0001 | $450.00 |
| Maple Street Bakery | 09/30/2026 | INV-2026-0002 | INV-2026-0002 | $1,200.00 |
| Riverside Yoga | 10/02/2026 | INV-2026-0003 | INV-2026-0003 | $315.00 |
| Juniper Florist | 10/05/2026 | INV-2026-0004 | INV-2026-0004 | $780.00 |
| Bluebird Books | 10/09/2026 | INV-2026-0005 | INV-2026-0005 | $95.00 |
| Cedar Hardware | 10/12/2026 | INV-2026-0006 | INV-2026-0006 | $640.00 |

Both formula columns gave identical results in LibreOffice. The numbers run from INV-2026-0001 to INV-2026-0006 in the order the rows appear.

The result is text, not a number. That is why the leading zeros stay.

## Troubleshooting

### The invoice numbers change when I sort or delete rows
The formula uses the row position, not a saved counter. Sort the table or delete a row and every number below it shifts. Once an invoice is sent, copy the number cells and use **Paste Special > Values only** so the numbers stay fixed. This is how the template's notes tell you to handle it.

### The count does not restart in January
The year comes from the invoice date, but the count comes from the row position. A row dated 01/05/2027 in row 40 would read INV-2027-0039. Start a new sheet or tab each year so the count begins at 0001 again.

### The year shows as text like "yyyy" or comes out wrong
A date format code such as `yyyy` can depend on the language or locale of the app. This was not reproduced here, because the template was only checked in LibreOffice. If it happens, use the column D formula, which takes the year with `YEAR(B2)` and has no format code in the year part.

### The year is wrong or the formula returns an error
`TEXT` and `YEAR` read the value in B2 as a date. If the date was typed or pasted as text, retype it as MM/DD/YYYY and format the cell as a Date. This error was not reproduced in the template, so check that cell first.

## Template

[Download the .xlsx](templates/tidy-tabs-invoice-number-generator.xlsx)

The file has two tabs. **Invoices** holds the six sample rows with both formula columns, the TEXT version in C and the YEAR version in D. **How to use** has short notes on the formulas and on pasting values once an invoice is sent.

The clients and amounts are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
