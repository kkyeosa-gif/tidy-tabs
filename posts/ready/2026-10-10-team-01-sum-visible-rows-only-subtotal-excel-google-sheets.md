---
title: How to Sum Only the Visible Rows After Filtering in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. Filter criteria and hidden rows were written into a file and recalculated; Excel and Google Sheets were not opened.
search_description: Use =SUBTOTAL(109,C2:C16) to total only the rows a filter leaves visible. SUM ignores filters. Counts, averages and a free .xlsx template included.
image_prompts: Small online shop packing table with three Etsy-style parcels in kraft mailers beside a stack of mixed shipping envelopes and a roll of tape, soft window light, no screens, no readable text, no logos
image_alt: Kraft mailers sorted apart from a pile of mixed envelopes next to tape
threads: Filtered your sales list and SUM still shows the same total?\nSUM adds hidden rows too. =SUBTOTAL(109,C2:C16) only adds what the filter leaves visible.\nIn the sample, Etsy shows $75.50 while SUM stays at $200.50.
---
Use `=SUBTOTAL(109,C2:C16)` instead of `SUM` to total only the rows a filter leaves visible. `SUM` keeps adding filtered-out rows, so its total does not change when you filter.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and the results below were recalculated and checked against hand math. The Excel and Google Sheets steps use the same SUBTOTAL function and link to Microsoft and Google help pages. Neither Excel nor Google Sheets was opened for this post.

![Sales list filtered to Etsy with visible total, count and average beside the unchanged SUM total](images/2026-10-10-team-01-sum-visible-rows-only-subtotal-excel-google-sheets/subtotal-filtered-sales-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Keep your list in one block with a header row, one sale per row. Here the dates are in column A, the channel in B and the amount in C.

1. Type headers in row 1: Date, Channel, Amount. Type the sales in rows 2 to 16.
2. Select any cell in the list, then turn on the filter. In Excel, use **Data > Filter**. In Google Sheets, use **Data > Create a filter**.
3. Put your totals somewhere a filter cannot hide them: a separate tab, or a cell above the header row. A total typed under the last row can disappear when a filter hides that row.
4. For the visible total, type `=SUBTOTAL(109,C2:C16)`.
5. For the visible count, type `=SUBTOTAL(103,C2:C16)`.
6. For the visible average, type `=SUBTOTAL(101,C2:C16)`.
7. Click the filter arrow on the Channel column and pick one channel. The three results change to match the rows you can see.

The first argument tells SUBTOTAL what to do. 109 means sum, 103 means count non-empty cells, and 101 means average. The second argument is the range to work on.

If your totals live on another tab, add the sheet name to the range, for example `=SUBTOTAL(109,Sales!C2:C16)`. The Totals tab in the template does exactly that.

If you add rows below row 16, extend the range in each formula.

### In Excel

Microsoft documents the function numbers on its [SUBTOTAL function](https://support.microsoft.com/en-us/office/subtotal-function-7b027003-f060-4ade-9040-e478765b9939) page. That page also explains how the 1 to 11 codes and the 101 to 111 codes treat rows you hide by hand. The two sets differ there, so read that page before relying on hidden rows. This post only checked rows hidden by a filter.

### In Google Sheets

Google describes SUBTOTAL on its [SUBTOTAL help page](https://support.google.com/docs/answer/3093649). The same formulas should work as typed. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is 15 sales from a small shop, with Sales in A1:C16 and the results on a Totals tab.

| Cell on Totals | Formula | No filter | Channel = Etsy | Channel = Farmers market |
|---|---|---|---|---|
| C2 | `SUBTOTAL(109)` visible total | $200.50 | $75.50 | $68.50 |
| C3 | `SUBTOTAL(103)` visible count | 15 | 3 | 6 |
| C4 | `SUBTOTAL(101)` visible average | $13.37 | $25.17 | $11.42 |
| C5 | `SUM` of the same range | $200.50 | $200.50 | $200.50 |

Check Etsy by hand: $34.00 + $18.50 + $23.00 is $75.50, and $75.50 divided by 3 is 25.1666..., shown as $25.17. Check Farmers market: $22.00 + $8.50 + $6.00 + $14.00 + $7.50 + $10.50 is $68.50.

The last row is the point of the post. `SUM` reads $200.50 in all three columns because it adds every row, hidden or not.

How the filters were checked: the filter criteria and the hidden rows were written into a copy of the file, then LibreOffice Calc 24.2.7 recalculated it. Nobody clicked the filter arrow in the LibreOffice interface for this post, and nothing was run in Excel or Google Sheets.

## Troubleshooting

### The total does not change when you filter
You are probably using `SUM`. Replace it with `=SUBTOTAL(109,C2:C16)`. `SUM` adds filtered-out rows too.

### The total sits under the data and vanishes
A filter hides whole rows. If the total is in row 17 and your filter range reaches it, a filter can hide that row. Move the totals to another tab or above the header, as the template does.

### The count shows every row
Check the first argument. 103 counts visible non-empty cells. A plain `COUNTA` or `COUNT` counts hidden rows as well.

### Rows you hid by hand behave differently
The 1 to 11 codes and the 101 to 111 codes differ on rows hidden by hand. Microsoft's help page describes the difference. Do not assume Excel and Google Sheets treat these the same way without checking the help pages above. The template check here covered filter-hidden rows only.

### The result is $0.00 or a number looks too small
An amount may be stored as text, for example a number typed with a leading apostrophe. SUBTOTAL skips text. Retype it as a plain number such as 22.00.

## Template

[Download the .xlsx](templates/tidy-tabs-subtotal-filtered-sales.xlsx)

The file has three tabs. **Sales** holds 15 sample sales in A1:C16 with an AutoFilter already on. **Totals** shows each formula as text in column B and the live formulas in C2:C6: visible total, visible count, visible average, a plain `SUM` for comparison, and a `COUNT`. **How to use** has short fill-in notes.

The channels and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
