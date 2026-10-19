---
threads_url: https://www.threads.com/@tin_ylab/post/DerOeEdG2Ih
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-your-on-time-shipping.html
title: How to Calculate Your On Time Shipping Rate in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 30 sample orders were checked against a Python date comparison, plus edge cases: shipped before the order date, a shipped date typed as text, an order date typed as text, all orders late, no orders shipped and a blank row inside the list. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Calculate an on time shipping rate with =COUNTIF(F2:F31,"On time")/COUNT(D2:D31) in Excel and Google Sheets. Free order tracker template.
image_prompts: Three taped cardboard shipping boxes stacked beside a roll of packing tape and a postal scale on a packing table, soft daylight from a side window, no screens, no readable text, no logos
image_alt: Taped cardboard shipping boxes stacked next to a packing tape roll and postal scale
threads: Shipped on the ship-by date still counts as on time.\n=IF(D2<=C2,"On time","Late") marks each order, then on time / (on time + late) gives the rate.\nOrders with no shipped date stay out of both counts, so the rate is not dragged down by boxes still on the shelf.
---
Mark each shipped order On time when the shipped date is on or before the ship by date, then divide the on time count by the shipped count with `=COUNTIF(F2:F31,"On time")/COUNT(D2:D31)`. Orders with no shipped date are left out of both counts.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Orders table with late rows highlighted beside a summary showing 82.1% on time](images/2026-10-18-team-04-on-time-shipping-rate-promised-vs-shipped-date-excel-google-sheets/on-time-shipping-rate-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The **Orders** tab has one order per row: Order ID in **A**, Order date in **B**, Ship by date in **C**, Shipped date in **D**, Days late in **E** and Status in **F**. Type dates as MM/DD/YYYY and leave **D** blank for orders that have not shipped.

1. On a tab named **Orders**, type the headers Order ID, Order date, Ship-by date, Shipped date, Days late and Status in **A1:F1**.
2. From row 2 down, type each order. Format **B:D** as dates so they are real dates, not text.
3. In **F2**, type `=IF(A2="","",IF(OR(NOT(ISNUMBER(B2)),NOT(ISNUMBER(C2))),"Check",IF(D2="","Not shipped",IF(NOT(ISNUMBER(D2)),"Check",IF(OR(D2<B2,C2<B2),"Check",IF(D2<=C2,"On time","Late"))))))`.
4. The part that does the work is `IF(D2<=C2,"On time","Late")`. The rest shows Check for a date typed as text or a date before the order date, and Not shipped when **D2** is blank.
5. In **E2**, type `=IF(F2="Late",D2-C2,IF(F2="On time",0,""))`. This is the number of days past the ship by date.
6. Select **E2:F2** and drag the fill handle down past your last order. The template fills to row 200.
7. On a tab named **Summary**, count each status with `=COUNTIF(Orders!$F$2:$F$200,"On time")`, and the same with "Late", "Not shipped" and "Check". Put them in **B4:B7**.
8. In **B8**, type `=IF(B4+B5=0,"",B4/(B4+B5))` and format it as a percentage with one decimal. This is on time divided by on time plus late.

The formula in the first paragraph divides by `COUNT(D2:D31)`, the number of shipped dates. The template divides by on time plus late instead, so a shipped date sitting on a Check row is not counted. With no Check rows the two give the same rate.

### In Excel

The formulas work as typed. Microsoft describes the functions on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) pages.

### In Google Sheets

The same formulas should work as typed. Google documents the counting function on its [COUNTIF help page](https://support.google.com/docs/answer/3093480). Nothing here was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It shows 6 of the 30 orders in the template, each with a 3 day ship by window.

| Order ID | Order date | Ship by date | Shipped date | Days late | Status |
|---|---|---|---|---|---|
| #2401 | 09/01/2026 | 09/04/2026 | 09/02/2026 | 0 | On time |
| #2403 | 09/03/2026 | 09/06/2026 | 09/06/2026 | 0 | On time |
| #2405 | 09/05/2026 | 09/08/2026 | 09/09/2026 | 1 | Late |
| #2409 | 09/09/2026 | 09/12/2026 | 09/14/2026 | 2 | Late |
| #2414 | 09/14/2026 | 09/17/2026 | 09/21/2026 | 4 | Late |
| #2429 | 09/29/2026 | 10/02/2026 | | | Not shipped |

Order #2403 shipped on its ship by date, so it is On time. Across all 30 orders, 28 have a shipped date, 23 are on time and 5 are late (#2405, #2409, #2414, #2419 and #2425). The 2 unshipped orders are left out, so **B8** shows 23 / 28 = 82.1%. Late orders average 2.2 days late and the worst is 4. These matched a Python date comparison.

## Troubleshooting

### One order shows Check
A date is typed as text, or the shipped date is earlier than the order date. Fix the date and the status updates. In the recalculated copy, shipping #2402 on 08/30/2026, before its order date, made that row Check and moved the rate from 82.1% to 22 / 27 = 81.5%.

### The rate is lower than the Late count suggests
A shipped date typed as text, such as 09/06/2026 left-aligned in **D4**, is skipped by `COUNT`. That row shows Check and stays out of the rate. Retype the date so it is right-aligned.

### The rate cell is blank
No order has a shipped date yet, or none are on time or late. The formula returns an empty cell instead of #DIV/0!. In the recalculated copy with all 30 shipped dates cleared, the rate stayed blank and all 30 orders showed Not shipped.

### The check cell says CHECK
**B11** shows OK only when on time plus late equals the shipped count and no row needs checking. Look for a Check status in column **F**.

### An order shipped on the ship by date shows Late
Both cells probably have a time of day, not just a date. The template compares whole dates only, so remove the times. This case was not run.

## Limits

This compares dates you type and nothing else. It ignores times of day, weekends and holidays. Marketplaces set and measure their own shipping targets, so use your seller help pages for the official numbers. The formulas run to row 200 and the sheet does not use TODAY.

## Template

[Download the .xlsx](templates/tidy-tabs-on-time-shipping-rate.xlsx)

The file has three tabs. **Orders** holds yellow input cells in **A2:D60**, with the Days late and Status formulas ready to row 200, and late rows are colored. **Summary** has the counts, the on time rate, average and longest delay, and an OK or CHECK cell. **How to use** has short notes. The Orders tab prints the first 40 orders on US Letter.

The 30 orders are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
