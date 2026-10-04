---
threads_url: https://www.threads.com/@tin_ylab/post/DeFIERUm8JB
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-build-accounts-receivable-aging.html
title: How to Build an Accounts Receivable Aging Report in Excel and Google Sheets
labels: excel-formulas, invoicing, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Build an accounts receivable aging report: days past due with MAX, Current to 90+ buckets, and SUMIFS totals per bucket. Free .xlsx template.
image_prompts: A small-business back-office counter with a stack of paper invoices held by binder clips, a beige desktop adding machine with a paper tape, and a red ink stamp pad; some invoices are pulled out of the stack, all text is blurred and unreadable, no logos, natural window light
image_alt: Stack of paper invoices beside a desktop calculator and a red ink stamp
threads: Which invoices are actually late, and by how much?\nMAX(0,AsOfDate-DueDate) gives days past due, and SUMIFS totals each bucket: Current, 1-30, 31-60, 61-90, 90+.\nIn the sample file, 8 unpaid invoices add up to $4,615.00, and $450.00 of that is 90+ days late.
---
Sort unpaid invoices by how late they are with `=MAX(0,AsOfDate-DueDate)`, label each one Current, 1-30, 31-60, 61-90, or 90+, and total each bucket with SUMIFS.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every formula and bucket total below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Aging summary with dollar totals for Current, 1-30, 31-60, 61-90, and 90+ day buckets](images/2026-09-30-team-01-receivables-aging-report-excel-google-sheets/receivables-aging-report-template.png) *LibreOffice Calc 24.2 PDF export of the Aging sheet with sample data (fake).*

## Steps

Set up two tabs: **Invoices** (one row per invoice) and **Aging** (the summary). The As of date lives in cell **B1** of the Aging tab.

1. On the Aging tab, type your report date in **B1**, for example 09/30/2026. Type `=TODAY()` instead if you want the report to always age as of today.
2. On the Invoices tab, use these headers in row 1: Invoice #, Client, Invoice date, Terms (days), Due date, Amount, Paid on, Balance, Days past due, Bucket.
3. In **E2** (Due date), type `=C2+D2`. Invoice date plus terms in days is the due date.
4. In **H2** (Balance), type `=IF(OR(F2="",G2<>""),0,F2)`. An invoice with a Paid on date gets a balance of 0.
5. In **I2** (Days past due), type `=IF(H2=0,"",MAX(0,Aging!$B$1-E2))`. `MAX(0, ...)` keeps invoices that are not late yet at 0 instead of a negative number.
6. In **J2** (Bucket), type `=IF(H2=0,"Paid",IF(I2=0,"Current",IF(I2<=30,"1-30",IF(I2<=60,"31-60",IF(I2<=90,"61-90","90+")))))`.
7. Fill E2, H2, I2, and J2 down for every invoice row.
8. On the Aging tab, list the buckets in **A4:A8**: Current, 1-30, 31-60, 61-90, 90+.
9. In **B4**, type `=SUMIFS(Invoices!$H$2:$H$500,Invoices!$J$2:$J$500,A4)` and fill it down to **B8**. Put `=COUNTIFS(Invoices!$J$2:$J$500,A4)` in **C4** and fill down to count invoices per bucket.
10. In row 9, add **Total outstanding** with `=SUM(B4:B8)` and `=SUM(C4:C8)`.

### In Excel

The formulas above work as typed. Format E, G, and the As of date as dates and F, H, and B4:B9 as currency. For argument details see Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) pages.

### In Google Sheets

The same formulas work as typed. Format the date columns as Date and the dollar columns as Currency. Google explains number formats on its [Format numbers in a spreadsheet page](https://support.google.com/docs/answer/56470) and documents the arguments on its [SUMIFS function](https://support.google.com/docs/answer/3238496) page.

## Example

Sample data below is fictional. The As of date is 09/30/2026.

| Invoice # | Client | Due date | Balance | Days past due | Bucket |
|---|---|---|---|---|---|
| 1001 | Harbor Dental | 07/31/2026 | $0.00 (paid 08/05/2026) | blank | Paid |
| 1002 | Pine & Co. Realty | 08/14/2026 | $1,200.00 | 47 | 31-60 |
| 1003 | Lakeside Yoga | 09/09/2026 | $380.00 | 21 | 1-30 |
| 1004 | Main St. Gift Shop | 09/09/2026 | $300.00 | 21 | 1-30 |
| 1005 | Northside Print Co. | 10/05/2026 | $875.00 | 0 | Current |
| 1006 | Desert Bloom Candles | 06/19/2026 | $450.00 | 103 | 90+ |
| 1007 | Wasatch Bike Repair | 07/30/2026 | $220.00 | 62 | 61-90 |
| 1008 | Gateway Bakery | 10/20/2026 | $540.00 | 0 | Current |
| 1009 | Harbor Dental | 10/15/2026 | $650.00 | 0 | Current |

The Aging tab then totals each bucket:

| Bucket (days past due) | Balance | Invoices |
|---|---|---|
| Current | $2,065.00 | 3 |
| 1-30 | $680.00 | 2 |
| 31-60 | $1,200.00 | 1 |
| 61-90 | $220.00 | 1 |
| 90+ | $450.00 | 1 |
| Total outstanding | $4,615.00 | 8 |

Invoice 1001 is paid, so it drops out of every bucket. Invoice 1004 uses 15-day terms, which is why it came due on the same day as 1003 despite a later invoice date.

## Troubleshooting

### Days past due shows a date like 02/16/1900 instead of 47
The cell inherited a date format. Select column I and change the format to Number or General.

### Every bucket shows $0.00
The SUMIFS criteria cell does not match the label text. Bucket names in **A4:A8** must match the Bucket column exactly, including the hyphen in `1-30`. Spreadsheet apps can also read a typed `1-30` as a date, so format A4:A8 as plain text before typing the buckets.

### A paid invoice still shows a balance
The Paid on cell holds a space or text, or the balance formula was not filled down. Clear the cell and retype the date, then check that H has the formula on every row.

### New invoices are missing from the totals
The SUMIFS ranges stop at row 500. Invoices past row 500 need the ranges extended, for example `$H$2:$H$1000`.

### Days past due does not change from day to day
The As of date is a typed date. Replace Aging!**B1** with `=TODAY()` to age as of today. Use a fixed date when you want a month-end report that stays put.

## Template

[Download the .xlsx](templates/tidy-tabs-receivables-aging-report.xlsx)

The file has three tabs. **Invoices** holds the nine sample invoices with the formulas from the steps above, and shades the formula columns (Due date, Balance, Days past due, Bucket). Invoices in the 61-90 and 90+ buckets are highlighted. **Aging** has the As of date in B1, the bucket totals, and the total row. **How to use** has short fill-in notes.

The clients and amounts are fictional. The file only groups invoices by how late they are; it makes no decision about collection or write-offs. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
