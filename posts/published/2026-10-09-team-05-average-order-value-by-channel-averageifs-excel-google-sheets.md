---
threads_url: https://www.threads.com/@tin_ylab/post/DeUXHGQG0fp
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-average-order-value-by.html
title: How to Find Average Order Value by Sales Channel With AVERAGEIFS in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same AVERAGEIFS, COUNTIFS and SUMIFS functions; the help link is from Microsoft Support (not tested in either app).
search_description: Average order value per channel with =AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2), plus order counts and a #DIV/0! fix. Free .xlsx template.
image_prompts: Small online shop packing table with a roll of kraft packing paper and a shipping box, a folded market-stall canopy tag, and a small tray of handmade soap bars, soft daylight, any labels blurred and unreadable, no screens, no logos
image_alt: Packed shipping box and tray of handmade soap bars beside a folded market tent
threads: Which sales channel actually brings your biggest orders?\n=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2) gives the average per channel name.\nPut =COUNTIFS next to it so you can see how many orders each average rests on. A channel with no orders returns #DIV/0!.
---
Use `=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)` to get the average order amount for the channel named in A2, such as Etsy, Shopify or Farmers market. Add `=COUNTIFS(Orders!$B$2:$B$200,A2)` next to it to show how many orders each average is based on.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support. Neither app was opened for this post.

![Channel summary showing orders, revenue and average order value for Etsy, Shopify and Farmers market](images/2026-10-09-team-05-average-order-value-by-channel-averageifs-excel-google-sheets/average-order-value-by-channel-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Keep every order on one tab, one row per order. Put the channel name in column B and the amount in column C.

1. Name one tab **Orders**. In row 1, type these headers in A to C: Order date, Channel, Amount.
2. Type each order below the headers. Spell each channel the same way every time, for example Etsy, Shopify, Farmers market.
3. Add a second tab named **By channel**. In row 1, type these headers in A to E: Channel, Orders, Revenue, Average order value, Check (revenue / orders).
4. In **A2:A4**, type one channel name per row.
5. In **B2**, type `=COUNTIFS(Orders!$B$2:$B$200,A2)`. This counts the orders for that channel.
6. In **C2**, type `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)`. This adds up the revenue.
7. In **D2**, type `=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)`.
8. In **E2**, type `=C2/B2`. It should match D2.
9. Select **B2:E2** and drag the fill handle down to row 4.
10. In **A5**, type All channels. In **B5** and **C5**, type `=SUM(B2:B4)` and `=SUM(C2:C4)`. In **D5**, type `=AVERAGE(Orders!C2:C200)`.
11. Format C and D as currency with two decimals.

AVERAGEIFS takes the range to average first, then the range to test, then the value it must match. Here it averages the Amount column wherever the Channel column equals A2. The `$` signs lock the ranges so they stay put when you fill down.

The ranges run to row 200. If your order log is longer, change 200 to a bigger number in every formula.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [AVERAGEIFS function](https://support.microsoft.com/en-us/office/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690) page and its [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) page.

### In Google Sheets

The same formulas should work as typed, since Sheets has AVERAGEIFS, COUNTIFS and SUMIFS. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is ten orders from a small shop that sells on Etsy, on Shopify and at a farmers market.

| Channel | Orders | Revenue | Average order value |
|---|---|---|---|
| Etsy | 3 | $115.50 | $38.50 |
| Shopify | 3 | $219.50 | $73.17 |
| Farmers market | 4 | $116.00 | $29.00 |
| All channels | 10 | $451.00 | $45.10 |

Check Etsy by hand: $115.50 divided by 3 orders is $38.50. Shopify is $219.50 divided by 3, which is 73.1666..., shown as $73.17. In the recalculated file, column E matched column D on every channel row.

The all-channels figure is the average of the 10 orders, $451.00 divided by 10. It is not the average of the three channel averages, which would be $46.89. The two differ because the channels have different order counts. That $46.89 is plain arithmetic and not a cell in the template.

Shopify has the highest average at $73.17, but Farmers market has the most orders. That is why the Orders column sits next to the average.

## Troubleshooting

### The cell shows #DIV/0!
AVERAGEIFS returns #DIV/0! when no row matches the channel name, because there is nothing to average. This is documented behavior and was not reproduced in a test here. If you list channels you have not sold through yet, wrap the formula as `=IFERROR(AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2),"")`.

### A channel shows fewer orders than expected
The spelling probably differs. "Etsy" and "Etsy " with a trailing space are different values to a matching formula, and so are "Farmers market" and "Farmers Market" in some apps. Retype the name or fix the Orders column. Case and space handling was not tested across apps.

### Column E does not match column D
Divide revenue by orders by hand for that row. If the two disagree, an amount in that channel's rows is probably stored as text or left blank. Retype it as a plain number such as 52.50. The template's check column exists to catch this.

### The all-channels average looks wrong
Make sure D5 reads `=AVERAGE(Orders!C2:C200)`. Averaging D2:D4 gives the average of the channel averages, which weights a channel with three orders the same as one with thirty.

### Orders after row 200 are ignored
Every formula stops at row 200. Raise that number in all of them if your log grows past it.

## Template

[Download the .xlsx](templates/tidy-tabs-average-order-value-by-channel.xlsx)

The file has three tabs. **Orders** holds 10 sample orders with a date, channel and amount. **By channel** has the order count, revenue, average order value and check column for three channels, plus an all-channels row. **How to use** has short fill-in notes, including how to add date conditions to average a single month.

The channels and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
