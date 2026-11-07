---
threads_url: https://www.threads.com/@tin_ylab/post/Dfcx-AumkKJ
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-return-by-date-from.html
title: How to Calculate a Return-By Date From the Delivery Date in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 26 workbooks and 13,208 comparisons with 0 mismatches. Checked: the sample, no orders, a delivery on the as-of date and one day either side of the 30 day edge, both sides of the Closing soon limit, a date with a time, a text date, a blank, 0, negative and decimal windows, per-order windows of 0, 14, 30, 365 and 400000, text overrides, year end, the 02/29/2028 leap day, the last Excel date 12/31/9999, an as-of date before every delivery, bad settings, 100 rows, and six random sets. Excel and Google Sheets were not used. Not run: TODAY(), a non-English weekday name, date serials below 61.
search_description: Add a return window to the delivery date in Excel and Google Sheets with INT, then count the days left. Free template flags Closing soon and Closed.
image_prompts: A brown paper parcel tied with string, a roll of packing tape and a blank return label on a worn wooden packing table, soft side light, realistic photo, no screens, no readable text, no logos
image_alt: Brown paper parcel tied with string beside a tape roll and an empty label
threads: Customer asks if they can still send it back?\nReturn-by date: =INT(B11)+D11. Days left: =E11-$B$3, with the as-of date typed in B3.\nA parcel delivered 10/20/2026 on a 30 day window is good until 11/19/2026. The window is your own policy.
---
Add the return window to the delivery date with `=INT(B11)+D11`, then count the days left with `=E11-$B$3`, where B3 is an as-of date you type. A parcel delivered 10/20/2026 with a 30 day window is returnable until Thursday 11/19/2026, which is 4 days left as of 11/15/2026 (your own shop policy, not a legal rule).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (26 workbooks, 13,208 comparisons). The formulas use only INT, IF, AND, ISNUMBER and TEXT, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Return-by dates and days left for twelve sample orders, with a status column](images/2026-11-07-team-02-return-by-date-from-delivery-date-days-left-excel-google-sheets/return-by-date-and-days-left-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. On **Orders**, the settings are in **B2:B4** and the orders start in row 8: order number in **A**, delivery date in **B**, an optional window for that order in **C**.

1. In **B2**, type your default return window in days, such as 30. This is your own shop policy in calendar days.
2. In **B3**, type the as-of date, such as 11/15/2026. It is typed on purpose. `TODAY()` changes every time the file opens, so a printed copy could not be matched to a date.
3. In **B4**, type how many days or fewer count as Closing soon, such as 7.
4. From row 8, log each order and its delivery date. Leave **C** empty to use the default window, or type a number to override it for that order.
5. In **D8**, type `=IF(B8="","",IF(C8="",$B$2,C8))` for the window used. The template wraps it in more checks so text or a negative number shows Check.
6. In **E8**, type `=INT(B8)+D8`. This is the return-by date. `INT` drops any time of day from the delivery date. Format the column as a date.
7. In **F8**, type `=TEXT(E8,"ddd")` to show the weekday.
8. In **G8**, type `=E8-$B$3`. Negative means the window has closed.
9. In **H8**, type `=IF(G8<0,"Closed",IF(G8=0,"Last day",IF(G8<=$B$4,"Closing soon","Open")))`.
10. Fill **D8:H8** down to row 107.

The delivery date counts as day 0. A 30 day window delivered on 10/20/2026 ends on 11/19/2026, and a window of 0 makes the delivery day the last day.

Help pages: Microsoft's [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) and [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349). They describe the functions, not our test.

## Example

Sample data is fictional: 12 rows, 11 with a delivery date. Default window 30 days, as-of date 11/15/2026, Closing soon at 7 days or fewer. Six of the rows:

| Order | Delivered | Window | Return-by | Day | Days left | Status |
|---|---|---|---|---|---|---|
| 1041 | 10/09/2026 | 30 | 11/08/2026 | Sun | -7 | Closed |
| 1044 | 10/20/2026 | 30 | 11/19/2026 | Thu | 4 | Closing soon |
| 1046 | 11/02/2026 | 14 | 11/16/2026 | Mon | 1 | Closing soon |
| 1047 | 11/04/2026 | 30 | 12/04/2026 | Fri | 19 | Open |
| 1049 | 11/13/2026 | 60 | 01/12/2027 | Tue | 58 | Open |
| 1051 | 11/15/2026 | 0 | 11/15/2026 | Sun | 0 | Last day |

Across all 11 delivered orders the summary shows 5 Open, 2 Closing soon, 1 Last day and 3 Closed, with no row showing Check. Order 1052 has no delivery date, so its row stays blank.

## Troubleshooting

### The row says Check
The delivery date is text or only a time, the window is negative, a decimal or text, or a setting in **B2:B4** is wrong. Retype the date as a plain date such as 11/03/2026 and the window as a whole number. A bad setting makes every row and **B5** show Check.

### The return-by date is a day off
The delivery date counts as day 0, so the window is added to it, not counted from the next day. If your policy counts the delivery day as day 1, type a window one day shorter.

### Days left changed after reopening the file
It should not. The as-of date in **B3** is typed. If you replaced it with `TODAY()`, the numbers move every day. We did not run `TODAY()` in the test.

### The weekday shows in another language
`TEXT(E8,"ddd")` follows the spreadsheet's language setting. We only ran it in English.

### A date with a time shows the right day
That is expected. `INT` drops the time, so 11/02/2026 3:40 PM counts as 11/02/2026.

## Template

[Download the .xlsx](templates/tidy-tabs-return-deadline.xlsx)

The file has two tabs. **Orders** has the three yellow settings in **B2:B4**, a **Settings check** in **B5**, 100 order rows from row 8 with 12 fictional sample orders, and a summary in **K7:K13**: delivered orders listed, Open, Closing soon, Last day, Closed, rows showing Check, and a **CHECK** cell. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The window is your own policy and the template gives no advice on what a buyer is entitled to. Carrier-promised delivery dates are not looked up.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
