---
threads_url: https://www.threads.com/@tin_ylab/post/Dezy-w6IIYE
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-your-busiest-hour-of-day.html
title: How to Find Your Busiest Hour of the Day From Order Timestamps in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model, including 15 edited variants for this file (boundary times such as 9:59:59 AM and 10:00:00 AM, midnight, a date with no time, a time typed as text, a duplicate hour in the list, a text amount, 299 orders). No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Count orders and add up sales by hour of the day with the HOUR function in Excel and Google Sheets. Free sales by hour template with sample orders.
image_prompts: A round analog wall clock above a small shop counter with a brass service bell and a stack of kraft paper parcels, soft afternoon light, no screens, no readable text, no logos
image_alt: Brass service bell and wrapped parcels on a shop counter under a wall clock
threads: Want your busiest hour from order timestamps?\nAdd a helper column: =IF(ISNUMBER(A2),HOUR(A2),"Check time"). Then COUNTIFS on the hour number.\nIn the sample, 10 AM wins with 9 orders and $237.00.
---
Group order times by hour with a helper column, `=IF(ISNUMBER(A2),HOUR(A2),"Check time")`, then count the orders in each hour with `=COUNTIFS(Orders!$D$2:$D$300,A2)`. HOUR returns a whole number from 0 to 23, so an order at 10:00 AM counts in hour 10.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![By hour table showing orders, sales and share of sales for each hour](images/2026-10-22-team-02-sales-by-hour-of-day-hour-countifs-sumifs-excel-google-sheets/sales-by-hour-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

You need a column of order date and time values, each in one cell, and a column of amounts.

1. On a tab named **Orders**, type or paste the order date and time in **A2:A300**, for example 10/10/2026 10:15:00 AM. Type the amounts in **B2:B300**.
2. In **D2**, type `=IF(ISNUMBER(A2),HOUR(A2),"Check time")` and fill it down to row 300. This is the helper column. It shows the hour, 0 to 23.
3. On a second tab named **By hour**, type the hours you want to see in **A2:A12** as numbers: 8 to 18.
4. In **C2**, count the orders in that hour: `=COUNTIFS(Orders!$D$2:$D$300,A2)`. Fill it down to C12.
5. In **D2**, add up the dollars in that hour: `=SUMIFS(Orders!$B$2:$B$300,Orders!$D$2:$D$300,A2)`. Fill it down to D12.
6. To name the busiest hour, use `=INDEX(B2:B12,MATCH(MAX(D2:D12),D2:D12,0))`, where B holds a label such as 10 AM.

Hour 10 means 10:00:00 AM up to 10:59:59 AM. An order at 9:59:59 AM is in hour 9.

### Why a helper column

The HOUR column is a whole number, so COUNTIFS and SUMIFS match it exactly. You do not need a "from 10:00 to 11:00" time range.

In LibreOffice Calc 24.2, a time range did drop a row. On a five row test, `=MOD(A3,1)` for 10:00:00 AM came back as 0.416666666656965, a hair under `TIME(10,0,0)`. A SUMIFS with `>=TIME(10,0,0)` on that column gave 30 where the HOUR helper gave 50. This was only run in LibreOffice. It was not run in Excel or Google Sheets, so this post does not say it happens there or that it does not.

### In Excel

The formulas work as typed. Microsoft documents the HOUR function in its Excel help. That is the help page, not our test.

### In Google Sheets

Google Sheets has HOUR, COUNTIFS and SUMIFS. See Google's Docs Editors Help for HOUR. The same formulas should work as typed, but Sheets was not opened for this post.

## Example

Sample data is fictional: 50 orders from a small shop that sells at a market and online, on 10/03/2026, 10/10/2026 and 10/11/2026.

| Hour | Orders | Sales |
|---|---|---|
| 8 AM | 3 | $47.00 |
| 9 AM | 4 | $124.50 |
| 10 AM | 9 | $237.00 |
| 11 AM | 6 | $187.00 |
| 12 PM | 5 | $138.00 |
| 1 PM | 4 | $85.00 |
| 2 PM | 5 | $137.50 |
| 3 PM | 3 | $54.00 |
| 4 PM | 3 | $76.50 |
| 5 PM | 3 | $48.00 |
| 6 PM | 2 | $51.00 |
| Outside these hours | 3 | $57.00 |
| Total | 50 | $1,242.50 |

In this fictional shop, the busiest hour by sales and by orders is 10 AM: 9 orders and $237.00, which is 19.07% of all sales. That describes these 50 sample orders only.

## Outside hours, midnight and bad times

The template adds three rows so nothing disappears:

- **Outside these hours** counts every order before the first hour or after the last one in your list. The sample has 7:45 AM, 7:10 PM and a 12:00:00 AM online order, so 3 orders and $57.00 land here.
- Midnight is hour 0, so it lands in that outside row. A date typed with no time is also hour 0.
- **Rows with a bad time** counts the rows where D says "Check time".

The cell **H7** on the By hour tab is a CHECK cell. It says OK when every order is counted once and the dollars add up to the Orders total. In the sample it says OK, with 50 orders and $1,242.50.

## Troubleshooting

### Column D says Check time
The timestamp is stored as text, for example 10/10/2026 9:59 AM typed into a cell formatted as Text. HOUR needs a real date and time. Retype it or convert it, then the hour appears.

### An order at 10:00 AM shows in the wrong hour
It should not. HOUR gives 10 for 10:00:00 AM and 9 for 9:59:59 AM. If a row looks wrong, check whether the cell holds the time you think it does, or whether it was a time range formula on a MOD column as described above.

### H7 says CHECK
One of four things: a bad time row, an amount with no time, a text amount such as "twelve", or the same hour listed twice in A2:A12. A duplicate hour means the orders for the hour after it are not counted.

### Midnight orders are missing from the table
They are in the Outside these hours row, because midnight is hour 0 and the table starts at 8. Change the hour list in A2:A12 if you want to see them.

## Template

[Download the .xlsx](templates/tidy-tabs-sales-by-hour.xlsx)

The file has three tabs. **Orders** holds the yellow inputs in **A2:C300** (order date and time, amount, item) and the hour helper in **D**. **By hour** holds the hours in **A2:A12**, the counts, sales and share of sales, the Outside and bad time rows, the summary and the CHECK cell. A white to green color scale shades the sales column. **How to use** has short notes. All tabs print portrait on US Letter.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
