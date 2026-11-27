---
title: How to Calculate Order Processing Time in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 225 workbooks and 44,764 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Count business days from order to shipment with NETWORKDAYS, then promise the 90th percentile rounded up. Free template with a coverage table.
image_prompts: A small home packing table with a row of handmade ceramic mugs half wrapped in brown kraft paper, a short stack of flat kraft mailer boxes, and a paper desk calendar with several weekdays crossed off in pencil, soft window daylight, realistic photo, no screens, no readable text, no logos
image_alt: Kraft-wrapped ceramic mugs, flat mailer boxes and a pencil-marked desk calendar
threads: The average is the wrong number to promise for processing time.\nIn the sample it was 2.72 business days, but 3 days covered only 67.65% of orders.\n=ROUNDUP(PERCENTILE(days,0.9),0) turned a 4.3 into a 5 business day promise, met by 65 of 68 orders.
---
Count the business days from order to shipment with `=NETWORKDAYS(B5,C5,Holidays)-1`, then promise the 90th percentile rounded up to a whole day: `=ROUNDUP(PERCENTILE(Orders!$D$5:$D$504,0.9),0)`. In the made-up shop's 68 shipped fall orders the 90th percentile is 4.3, so the promise is 5 business days, which 65 of the 68 orders met.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (225 workbooks, 44,764 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Shipping stats and coverage table showing a five-business-day promise met by 95.59% of orders](images/2026-11-28-team-04-set-order-processing-time-from-90th-percentile-business-days-networkdays-percentile-excel-google-sheets/processing-time-percentile-template.png) *Promise tab of a made-up shop's 68 shipped fall orders: average 2.72 and median 3.0 business days, a 90th percentile of 4.3 rounded up to a 5 business day promise in green, and the coverage table where 65 orders (95.59%) shipped within 5 business days. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

Processing time here means your own handling time: order date to ship date, in business days, skipping the days you are closed. Carrier transit is not included. Your marketplace or shop settings decide where the number goes, and past orders do not guarantee how December will go.

The steps are the same in Excel and Google Sheets.

1. On the **Orders** tab, type your closed days (holidays, vacation days) in **G5:G34**. The template names that range `Holidays`.
2. Type one order per row from row 5: Order # in **A**, Order date in **B**, Ship date in **C**. Leave **C** blank for an order that has not shipped yet.
3. Check column **E**. It shows OK, "Not shipped yet", "Outside the date range (not counted)", or a Check message such as "Check: ship date before order date".
4. Read the business days in column **D**: `=IF(E5<>"OK","",MAX(0,NETWORKDAYS(INT(B5),INT(C5),Holidays)-1))`. NETWORKDAYS counts both the order day and the ship day, so the -1 makes a next-business-day shipment 1 and a same-day shipment 0. INT drops any time of day, and MAX keeps a Saturday order shipped Sunday at 0 instead of -1. See the NETWORKDAYS help from [Microsoft](https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7) or [Google](https://support.google.com/docs/answer/3092979).
5. On the **Promise** tab, type the percentile in **B1** as 90% or 0.9, not 90. Optional: type a first and last order date in **B2** and **B3** to count only part of your history.
6. Read the summary in **B6:B11**: shipped orders counted, not shipped yet, average, median, fastest and slowest.
7. Read the percentile in **B12**: `=PERCENTILE(Orders!$D$5:$D$504,B1)` inside an IF guard. This is the inclusive percentile, the same as PERCENTILE.INC, not PERCENTILE.EXC. Blank cells in column D are skipped. See PERCENTILE and PERCENTILE.INC from Microsoft ([PERCENTILE](https://support.microsoft.com/en-us/office/percentile-function-91b43a53-543c-4708-93de-d626debdddca), [PERCENTILE.INC](https://support.microsoft.com/en-us/office/percentile-inc-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed)) or [Google](https://support.google.com/docs/answer/3094093).
8. Read the processing time to promise in **B13**: the same PERCENTILE wrapped in `ROUNDUP(...,0)`, shown as "5 business days".
9. Read how many past orders that promise covers: **B14** is `=COUNTIF(Orders!$D$5:$D$504,"<="&B13)` and **B15** divides it by the shipped count. **B16** shows the 95th percentile for comparison.
10. Use the coverage table in **D5:F16** to compare choices. Each row from 0 to 10 business days counts the orders shipped within that many days with `=COUNTIF(Orders!$D$5:$D$504,"<="&D6)` and shows the share. The promise row is green.

**CHECK** in **B23** is OK only when the settings (**B19**) and the closed-day list (**B20**) are valid, the Orders CHECK is OK, and a promise was calculated.

To turn a promise into a ship date for each new order, see [How to Calculate a Ship Date With a Daily Cutoff Time](https://tidytabs.blogspot.com/2026/11/how-to-calculate-ship-date-with-daily.html). To measure how often you kept it, see [How to Calculate Your On Time Shipping Rate](https://tidytabs.blogspot.com/2026/10/how-to-calculate-your-on-time-shipping.html).

## Example

The orders below are made up: 70 orders, #3101 to #3170, dated 10/02/2026 to 11/24/2026, with 2 not shipped yet. The shop is closed on 10/12, 11/11, 11/26 and 11/27/2026. A few rows from the **Orders** tab:

| Order | Order date | Ship date | Business days |
|---|---|---|---|
| #3101 | 10/02/2026 | 10/06/2026 | 2 |
| #3102 | 10/03/2026 (Sat) | 10/06/2026 | 1 |
| #3112 | 10/10/2026 (Sat) | 10/10/2026 | 0 |
| #3114 | 10/12/2026 (closed) | 10/16/2026 | 3 |
| #3149 | 11/05/2026 | 11/17/2026 | 7 (across 11/11) |
| #3155 | 11/11/2026 (closed) | 11/16/2026 | 2 |
| #3169 | 11/23/2026 | | Not shipped yet |

The **Promise** tab for the 68 shipped orders:

| Result | Value |
|---|---|
| Average | 2.72 |
| Median | 3.0 |
| Fastest / slowest | 0 / 7 |
| 90th percentile | 4.3 |
| Processing time to promise | 5 business days |
| Orders shipped within 5 days | 65 of 68 (95.59%) |
| 95th percentile | 5.0 |

And the coverage table:

| Business days | Orders shipped within | Share |
|---|---|---|
| 0 | 4 | 5.88% |
| 1 | 16 | 23.53% |
| 2 | 32 | 47.06% |
| 3 | 46 | 67.65% |
| 4 | 61 | 89.71% |
| 5 | 65 | 95.59% |
| 6 | 67 | 98.53% |
| 7 to 10 | 68 | 100.00% |

## Troubleshooting

These cases were reproduced in LibreOffice Calc 24.2.7.2.

### Every stat is blank and CHECK is red
**B1** was typed as 90 instead of 90% or 0.9, or a closed day in **G5:G34** was typed as text. With a text closed day, every order shows "Check: Promise settings or holidays". Retype the percentile or the date and the numbers come back.

### A weekend order shipped Monday shows 0, not 2
Weekends are skipped, and the -1 counts the next business day as 1 only for a weekday order. A Saturday 11/21/2026 order shipped Monday 11/23/2026 is 0: NETWORKDAYS counts only the Monday, and the -1 takes it away.

### The average is shorter than what you should promise
In the sample the average is 2.72 business days, but 3 days covered only 67.65% of orders. The rounded-up 90th percentile, 5 days, covered 95.59%. Promise from the coverage table, not the average.

### B8:B13 say "No shipped orders"
Every ship date is blank, or every shipped order is outside the dates in **B2** and **B3**. Clear or widen the date range, or add ship dates.

### An order shows "Check: ship date before order date"
The ship date in **C** is earlier than the order date in **B**, usually a typo in the month or day. Fix the date; the row is left out of every stat until then.

## Template

[Download the .xlsx](templates/tidy-tabs-processing-time-percentile.xlsx)

**Orders** takes 500 orders (rows 5 to 504), the closed-day list in **G5:G34**, and a summary in **I4:J9** that also catches rows typed below row 504. **Promise** holds the settings in **B1:B3**, the stats in **B6:B16**, the coverage table in **D5:F16** and the checks in **B19:B23** on one page. **How to use** explains each formula. Every tab prints on US Letter, the data tabs Landscape with the header row repeated.

Yellow cells are inputs and gray cells are formulas. The promise and its coverage row are light green, "Not shipped yet" and "Outside the date range" are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
