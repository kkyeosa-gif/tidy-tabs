---
threads_url: https://www.threads.com/@tin_ylab/post/DeHLWJiFBKr
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-add-business-days-to-date-in.html
title: How to Add Business Days to a Date in Excel and Google Sheets
labels: formulas, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Add business days to an order date with WORKDAY, skip weekends and a holiday list, and check the result with NETWORKDAYS. Free .xlsx template.
image_prompts: A shipping table in a small US online shop with three taped cardboard boxes and a tape gun, a paper wall calendar hanging above with several dates circled in red marker, all calendar text blurred and unreadable, no logos, daylight from a side window
image_alt: Packed shipping boxes and a tape gun below a wall calendar with circled dates
threads: 5 business days from 09/28/2026 is 10/05/2026, not 10/03.\nWORKDAY(C2,D2,Holidays!$A$2:$A$20) skips weekends plus every date in your holiday list, so a shipping promise survives Columbus Day and Thanksgiving.\nNETWORKDAYS minus 1 counts back to check it.
---
Use `=WORKDAY(C2,D2,Holidays!$A$2:$A$20)` to add business days to an order date; it skips weekends and every date in your holiday list.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template and every result below were checked. Excel and Google Sheets steps follow the Microsoft and Google help pages for WORKDAY and NETWORKDAYS. Neither app was hands-on tested here.

![Orders table listing order dates, business days needed, and calculated ship-by dates](images/2026-09-30-team-04-business-days-ship-by-date-excel-google-sheets/business-days-ship-by-date-template.png) *LibreOffice Calc 24.2 PDF export of the Orders sheet with sample data (fake). The Customer and Calendar days columns are hidden so the numbers stay readable.*

## Steps

1. Make a sheet named **Orders** with these headers in row 1: Order #, Customer, Order date, Business days to ship, Ship by, Calendar days, Business days check.
2. Type each order date in column C and the number of business days you need in column D. Use real dates, not text.
3. Make a second sheet named **Holidays**. Put a header in A1, then one date per row starting at A2. Add a name for each day off in column B if you like.
4. In **Orders**, click E2 and type `=WORKDAY(C2,D2,Holidays!$A$2:$A$20)`.
5. Format column E as a date.
6. Copy E2 down the column. The `$` signs keep the holiday range fixed on every row.

The holiday range runs to row 20, so you can add new days off below your last date without editing the formula.

### In Excel

1. Type the formula from step 4 in E2 and press **Enter**.
2. If the result shows a number such as 46300, apply a date format to the column.
3. Microsoft explains the arguments on its [WORKDAY function page](https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33).

### In Google Sheets

1. Type the same formula in E2 and press **Enter**.
2. If it shows a number, format the column as a date. Google explains number formats on its [Format numbers in a spreadsheet page](https://support.google.com/docs/answer/56470).
3. Google covers the arguments on its [WORKDAY help page](https://support.google.com/docs/answer/3093059).

### Check the result with NETWORKDAYS

WORKDAY gives you a date. To confirm it, count the business days back with `=NETWORKDAYS(C2,E2,Holidays!$A$2:$A$20)-1` in column G. NETWORKDAYS counts both the start and end day, so subtract 1.

The answer should equal column D. If it doesn't, a date is stored as text or the holiday range is wrong.

Column F, `=E2-C2`, shows how many calendar days the order actually takes. See the [NETWORKDAYS page from Microsoft](https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7) or the [Google version](https://support.google.com/docs/answer/3092979).

## Example

Customers and orders below are fictional sample data. The Holidays sheet lists five 2026 US days off: 09/07/2026, 10/12/2026, 11/11/2026, 11/26/2026, and 12/25/2026.

| Order # | Customer | Order date | Business days | Ship by | Calendar days | Check |
|---|---|---|---|---|---|---|
| 2001 | Harbor Dental | 09/28/2026 | 5 | 10/05/2026 | 7 | 5 |
| 2002 | Lakeside Yoga | 10/06/2026 | 10 | 10/21/2026 | 15 | 10 |
| 2003 | Main St. Gift Shop | 11/09/2026 | 3 | 11/13/2026 | 4 | 3 |
| 2004 | Gateway Bakery | 11/20/2026 | 5 | 11/30/2026 | 10 | 5 |

Order 2001 has no holiday, so 5 business days is one full week plus the weekend: 7 calendar days.

Order 2002 skips 10/12/2026, so 10 business days takes 15 calendar days. Order 2003 skips 11/11/2026, and order 2004 skips Thanksgiving on 11/26/2026.

The Check column returns 5, 10, 3, and 5, matching column D on every row.

## Troubleshooting

### The ship-by date shows a number like 46300
The cell has a General or Number format. Apply a date format to column E.

### A holiday is not skipped
The date is missing from the Holidays list, sits outside `A2:A20`, or is stored as text. Type it as a real date, or extend the range in every formula if your list is longer.

### The formula returns #VALUE!
Usually the order date is text, for example a pasted value with an apostrophe in front. Retype the date, or check that a left-aligned date is not a text string.

### Your weekend is not Saturday and Sunday
WORKDAY treats Saturday and Sunday as the weekend. If your shop closes on other days, use WORKDAY.INTL instead, which lets you pick the weekend days.

### The check column is off by one
You left out the `-1`. NETWORKDAYS counts the order date itself as a business day, so the raw count is one higher than column D.

## Template

[Download the .xlsx](templates/tidy-tabs-business-days-ship-by-date.xlsx)

The file has three tabs. **Orders** holds the four sample orders with the Ship by, Calendar days, and Business days check formulas already filled in. **Holidays** lists the five sample days off in A2:A6, and the formulas read down to row 20. **How to use** has short fill-in notes.

To open it in Google Sheets, go to **File > Import > Upload** and select the file. Replace the sample holidays with your own days off, such as market days or vacation, before you rely on the dates.
