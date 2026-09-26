---
title: How to Turn an Etsy Sales CSV Into a Monthly Summary With SUMIFS in Excel and Google Sheets
labels: excel-formulas, etsy-sellers, sales-reports
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small home craft workspace with a laptop open to a blurred spreadsheet, a stack of printed order slips, and handmade product samples nearby, all screen text and labels unreadable.
image_alt: Craft seller's desk with laptop, printed orders, and handmade products near a spreadsheet
search_description: Turn your Etsy order export into a monthly sales summary using SUMIFS in Excel or Google Sheets, with copy-paste formulas and a troubleshooting guide.
threads: Etsy's CSV export dumps every order into one long list, no monthly totals anywhere.\nFix: add a summary tab and use SUMIFS with a date range, like =SUMIFS(Orders!D:D, Orders!A:A, ">="&DATE(2026,9,1), Orders!A:A, "<"&DATE(2026,10,1)).\nWorks the same in Excel and Google Sheets once the date column is real dates, not text.
---
Add a summary sheet next to your Etsy order export and use `SUMIFS` with a date range for each month, like `=SUMIFS(Orders!D:D, Orders!A:A, ">="&DATE(2026,9,1), Orders!A:A, "<"&DATE(2026,10,1))`. This adds up the "Item Total" column for every order dated in September 2026.

> Works in: Excel (Microsoft 365, Excel 2019+) and Google Sheets. Menu names for importing and formatting may vary slightly across Excel for Windows, Mac, and the web. Steps come from Microsoft and Google support pages, not hands-on tested.

## Steps

1. Download your order export from Etsy: **Shop Manager > Settings > Options > Download Data**, choose Orders and a date range, then export as CSV. (Etsy's menu wording can change, so check the current option if this doesn't match what you see.)
2. Open a new spreadsheet and import the CSV as its own sheet, named something like `Orders`.
3. Add a second sheet named `Summary` for your monthly totals.
4. On the `Orders` sheet, confirm the date column (usually "Sale Date") is formatted as a date, not text. Text-formatted dates make `SUMIFS` return 0.
5. On `Summary`, list one month per row, for example `09/01/2026` in column A.
6. In column B, enter the SUMIFS formula for that month's total sales, using the first day of the month and the first day of the next month as your range.
7. Copy the formula down for each month, adjusting the two dates in each row.

### In Excel (Windows, Microsoft 365)

- To import the CSV, use **Data > Get Data > From File > From Text/CSV**, then choose **Load To** and select a new sheet. Mac and web versions of Excel offer similar CSV import options, but the exact menu path may differ.
- If the date column imports as text, select it and use **Data > Text to Columns**, then choose **Date** in the wizard to convert it.

### In Google Sheets

- Use **File > Import > Upload**, choose the CSV, and select **Insert new sheet(s)**.
- If dates land as text, select the column and use **Format > Number > Date** to fix it.

## Example

Sample data below is fictional, made up for this post.

**Orders sheet**

| Sale Date | Order Number | Buyer | Item Total |
|---|---|---|---|
| 09/03/2026 | 1001 | J. Rivera | $18.00 |
| 09/14/2026 | 1002 | K. Tran | $42.50 |
| 09/22/2026 | 1003 | M. Osei | $27.00 |
| 10/02/2026 | 1004 | S. Doyle | $35.00 |

**Summary sheet**

| Month | Total Sales |
|---|---|
| 09/01/2026 | $87.50 |
| 10/01/2026 | $35.00 |

The September row uses `=SUMIFS(Orders!D:D, Orders!A:A, ">="&A2, Orders!A:A, "<"&DATE(YEAR(A2),MONTH(A2)+1,1))`, where A2 holds `09/01/2026`. This adds every order dated in September without hardcoding the end date.

## Troubleshooting

### SUMIFS returns 0 for every month

Your date column probably imported as text instead of real dates. Select the column, convert it with **Text to Columns** (Excel) or **Format > Number > Date** (Sheets), then re-check the formula.

### The total is higher than expected

Etsy's export sometimes includes a refund or cancelled order as a separate row with a negative amount. Check your `Item Total` column for negative values before summing, or add a filter for order status if your export includes one.

### Some rows have the wrong month

If your CSV has a "Sale Date" and a separate "Ship Date," confirm your SUMIFS formula points at the sale date column, not the ship date.

### Currency symbols break the sum

If `Item Total` imported with a "$" as plain text, SUMIFS will skip those cells. Reformat the column as **Currency** or **Number**, or use **Find & Replace** to strip the "$" first.

### New months don't show up automatically

SUMIFS only reads the range you give it. Add a new row to the Summary sheet at the start of each month with that month's first date, and copy the formula down.

## Copy-paste setup

**Orders sheet headers** (row 1): `Sale Date`, `Order Number`, `Buyer`, `Item Total`

**Summary sheet headers** (row 1): `Month`, `Total Sales`

**Summary sheet formula** (row 2, column B, with the month's first date in A2):

```
=SUMIFS(Orders!D:D, Orders!A:A, ">="&A2, Orders!A:A, "<"&DATE(YEAR(A2),MONTH(A2)+1,1))
```

Copy this formula down for each new month row, keeping the column A reference relative so it updates automatically.

Reference pages for the formula and import steps:
- [SUMIFS function - Microsoft Support](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b)
- [SUMIFS - Google Docs Editors Help](https://support.google.com/docs/answer/3093301)
- [Import CSV files into Google Sheets](https://support.google.com/docs/answer/40608)
