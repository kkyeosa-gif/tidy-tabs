---
threads_url: https://www.threads.com/@tin_ylab/post/Dehc1CFGixO
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-average-days-to-get.html
title: How to Calculate Average Days to Get Paid in Excel and Google Sheets
labels: excel-formulas, small-business, invoicing
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python, including an unpaid invoice that gets a paid date, a sheet with no paid dates, a paid date before the invoice date, a paid date typed as text and a paid date with no invoice date. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Average days to get paid: =D2-C2 per invoice, then =AVERAGE(E2:E30). Eight paid sample invoices average 25.5 days. Free .xlsx template.
image_prompts: Small stack of paper invoices held by a brass binder clip next to a rubber date stamp and an ink pad on a wooden desk, soft window light, no screens, no readable text, no logos
image_alt: Paper invoices under a binder clip beside a date stamp and ink pad
threads: Want to know how long clients really take to pay?\n=D2-C2 gives days per invoice and =AVERAGE(E2:E30) gives the mean.\nThe eight paid sample invoices average 25.5 days, fastest 7, slowest 45.\nUnpaid invoices stay out of the average, and a text date can break it.
---
Subtract the invoice date from the paid date with `=D2-C2`, then average the results with `=AVERAGE(E2:E30)`. For the eight paid sample invoices, clients took 25.5 days on average, from 7 days at the fastest to 45 at the slowest.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same formulas as typed. Neither Excel nor Google Sheets was opened for this post.

![Invoice table with days to pay per client beside an average of 25.5 days](images/2026-10-15-team-01-average-days-to-get-paid-excel-google-sheets/days-to-get-paid-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: a date is a number in a spreadsheet, so subtracting two dates returns the days between them. Averaging those day counts gives your typical wait. A blank result is skipped by AVERAGE, which keeps unpaid invoices out.

1. On a tab named **Invoices**, type these headers in row 1: Invoice, Client, Invoice date, Paid date, Days to pay, Status.
2. Enter each invoice in rows 2 and down: number in **A**, client in **B**, invoice date in **C**, and the date the money arrived in **D**. Leave **D** empty while an invoice is unpaid.
3. In **E2**, type `=IF(OR(C2="",D2=""),"",D2-C2)`. This returns the days, or nothing if either date is missing. Copy it down.
4. In **F2**, type `=IF(C2="","",IF(D2="","Unpaid","Paid"))`. Copy it down.
5. Pick a cell for the average, such as **J4**, and type `=IFERROR(AVERAGE(E2:E30),"")`. The IFERROR keeps it blank instead of showing an error when nothing is paid yet.
6. Add the fastest and slowest: `=MIN(E2:E30)` and `=MAX(E2:E30)`.
7. Count the paid invoices with `=COUNT(E2:E30)` and the unpaid ones with `=COUNTIF(F2:F30,"Unpaid")`.
8. Put your payment terms, such as 30, in **J2**. Count the late ones with `=COUNTIF(E2:E30,">"&J2)`.

Enter real dates, such as 09/14/2026, not text. A date typed as text can behave differently by app and locale. In the recalculated copy, LibreOffice converted "07/21/2026" and returned the right number of days, but another app may leave it as text and give an error or a blank.
### In Excel

The formulas work as typed. Microsoft describes the functions on its [AVERAGE function](https://support.microsoft.com/en-us/office/average-function-047bac88-d466-426c-a32b-8f33eb960cf6) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) pages. If column **E** shows a date instead of a number, select it and choose **Home > Number Format > General**.

### In Google Sheets

The same formulas should work as typed. If column **E** shows a date, select it and choose **Format > Number > Automatic**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. Terms are 30 days, and two invoices are still unpaid.

| Invoice | Days to pay | Status |
|---|---|---|
| INV-2001 | 20 | Paid |
| INV-2002 | 35 | Paid |
| INV-2003 | 14 | Paid |
| INV-2004 | 35 | Paid |
| INV-2005 | 7 | Paid |
| INV-2006 | 45 | Paid |
| INV-2007 | 20 | Paid |
| INV-2008 | 28 | Paid |
| INV-2009 | (blank) | Unpaid |
| INV-2010 | (blank) | Unpaid |

| Result | Cell | Value |
|---|---|---|
| Average days to pay | J4 | 25.5 |
| Fastest | J5 | 7 |
| Slowest | J6 | 45 |
| Paid invoices | J7 | 8 |
| Unpaid invoices | J8 | 2 |
| Paid later than terms | J9 | 3 |
| Check | J10 | OK |

The eight paid invoices add up to 204 days, and 204 divided by 8 is 25.5. Three of them (35, 35 and 45) passed the 30-day terms.

## Troubleshooting

### Days to pay shows a date, not a number
Column **E** inherited a date format. Reset it to General in Excel or Automatic in Google Sheets.

### The average changed after I entered a paid date
That is expected. In the recalculated copy, INV-2009 got a paid date of 09/20/2026, which gave 4 days. The average moved from 25.5 to 23.1111, the paid count went to 9 and the unpaid count to 1.

### The result is negative
The paid date is earlier than the invoice date, usually a typo. In the recalculated copy, a paid date of 07/01/2026 against an invoice date of 07/08/2026 gave -7 and pulled the average down to 20.25. The template does not flag it, so scan the fastest value for negatives.

### The check cell says CHECK
Some row has a paid date but no invoice date. In the recalculated copy, that row stayed blank in **E** and **F**, so the paid and unpaid counts no longer matched the invoice list. Fill in the invoice date.

### The average is blank
No row has both dates yet. In the recalculated copy, with every paid date cleared, the average, fastest and slowest were blank, the paid count was 0 and all 10 invoices counted as unpaid.

## Limits

This looks back at invoices that were already paid. It is not an aging report, so an invoice that is 60 days late and still open does not raise the average. It counts calendar days, not business days. It gives one average for all clients, with no split per client.

## Template

[Download the .xlsx](templates/tidy-tabs-days-to-get-paid.xlsx)

The file has two tabs. **Invoices** holds 10 sample rows (2 unpaid) with room through row 30, yellow input cells, the formulas above, a payment terms cell and the summary. Days over your terms turn red. **How to use** has short fill-in notes. Formulas read rows 2 to 30.

The invoices, clients and dates are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
