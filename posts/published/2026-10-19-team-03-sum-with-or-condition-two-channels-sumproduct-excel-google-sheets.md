---
threads_url: https://www.threads.com/@tin_ylab/post/DesgD-vmkST
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-sum-sales-from-either-of-two.html
title: How to Sum Sales From Either of Two Channels (OR Logic) in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the Etsy plus Shopify totals were checked against a Python Decimal model for every pair of the 4 channels, plus the same channel twice, a channel with no rows, a blank channel cell, a trailing space, lower case text, an amount typed as text, a $64.00 threshold and an empty Orders tab. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Sum sales from either of two channels with SUMIFS plus SUMIFS or SUMPRODUCT in Excel and Google Sheets. Free OR template.
image_prompts: Two kraft shipping boxes, one with a small handmade product tucked inside, and a roll of packing tape beside a plain market cash tin on a craft table, soft natural window light, no screens, no readable text, no logos
image_alt: Packed shipping boxes and a tape roll beside a metal cash tin
threads: Etsy plus Shopify, one total.\nSUMIFS ignores OR on its own, so add two of them: =SUMIFS(...,B1)+SUMIFS(...,B2). If both cells hold the same channel, a SUMPRODUCT version avoids counting twice.
---
Add two SUMIFS together, one for each channel: `=SUMIFS(Orders!D:D,Orders!C:C,B1)+SUMIFS(Orders!D:D,Orders!C:C,B2)`, with the channel names typed in Totals B1 and B2. If both cells could ever hold the same channel, use `=SUMPRODUCT(((Orders!C2:C25=B1)+(Orders!C2:C25=B2)>0)*Orders!D2:D25)` so nothing is counted twice.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The formulas use functions that Excel and Google Sheets also have, but neither app was opened for this post.

![Orders table and totals block summing Etsy and Shopify sales for one small shop](images/2026-10-19-team-03-sum-with-or-condition-two-channels-sumproduct-excel-google-sheets/sum-with-or-two-channels-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

SUMIFS applies every condition at once, so it cannot take "Etsy or Shopify" for one column. Each SUMIFS finds one channel, and you add the results. This works because a row can only match one of them.

The orders sit on a tab named **Orders**: Date in **A**, Order ID in **B**, Channel in **C** and Amount in **D**, one order per row from row 2. The channels sit on a tab named **Totals**: channel 1 in **B1**, channel 2 in **B2** and a large-order threshold in **B3**.

1. On **Totals**, type the first channel (for example Etsy) in **B1** and the second (Shopify) in **B2**.
2. In **B5**, type `=SUMIFS(Orders!$D$2:$D$200,Orders!$C$2:$C$200,B1)+SUMIFS(Orders!$D$2:$D$200,Orders!$C$2:$C$200,B2)`. The template reads rows 2 to 200 so new orders are included.
3. For a version that cannot double count, type this in **B7**: `=SUMPRODUCT(--((Orders!$C$2:$C$200=$B$1)*($B$1<>"")+(Orders!$C$2:$C$200=$B$2)*($B$2<>"")>0),Orders!$D$2:$D$200)`. Each row adds 1 if it matches either channel, and the `>0` turns 2 back into 1, so a row counts once.
4. To total only the large orders from either channel, type this in **B12**: `=SUMPRODUCT(--((Orders!$C$2:$C$200=$B$1)*($B$1<>"")+(Orders!$C$2:$C$200=$B$2)*($B$2<>"")>0),--(Orders!$D$2:$D$200>B3),Orders!$D$2:$D$200)`. It adds an AND to the OR: either channel, and an amount over **B3**.

The template's SUMPRODUCT uses a comma between the test and the amounts, not `*`. When an amount is typed as text, the multiply form returns #VALUE!, while the comma form skips that row. The `($B$1<>"")` parts stop a blank channel cell from matching the empty rows below your data.

The template also has **B6**, `=SUM(SUMIFS(Orders!$D$2:$D$200,Orders!$C$2:$C$200,{"Etsy","Shopify"}))`, which puts both names in one array constant. The names are typed inside the formula, so edit it if you change **B1** or **B2**.

### In Excel

The formulas work as typed. Microsoft describes them on its [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) pages.

### In Google Sheets

The same formulas should work as typed. Google describes [SUMIFS](https://support.google.com/docs/answer/3238496) in its help. The array constant in **B6** and the comma form of SUMPRODUCT ran in LibreOffice Calc only, so check them against **B5** the first time you open the file in Sheets.

## Example

Sample data below is fictional. It is a small shop with 24 orders from 10/01/2026 to 10/15/2026 across four channels, with Etsy in **B1**, Shopify in **B2** and a $50.00 threshold in **B3**.

| Cell | What it shows | Result |
|---|---|---|
| B5 | SUMIFS for Etsy + SUMIFS for Shopify | $605.50 |
| B6 | One SUMIFS with both names in an array constant | $605.50 |
| B7 | SUMPRODUCT, each row counted once | $605.50 |
| B8 | COUNTIF + COUNTIF | 14 orders |
| B10 | All other channels | $870.00 |
| B11 | Everything minus the other channels | $605.50 |
| B12 | Either channel, amount over $50.00 | $339.00 |
| B13 | Check | OK |

Etsy brings in $257.00 and Shopify $348.50. **B13** says OK when **B5**, **B6**, **B7** and **B11** agree and the two order counts match. An order of exactly $50.00 would not count in **B12**, because the test is "over".

## Troubleshooting

### The total is exactly double what it should be

Both channel cells hold the same name, so the two SUMIFS count every matching row twice. With Etsy in both, **B5** shows $514.00, while **B7** shows the correct $257.00. The template's **B13** says Same channel twice. Pick two different channels, or use the SUMPRODUCT version.

### The total ignores some orders from one channel

Look for a trailing space, such as "Etsy " in the Channel column. Matching ignores upper and lower case but not spaces, so that row matches neither channel and falls into the other channels total. In the test, one such row dropped Shopify to $563.00 for the pair, and **B13** still said OK, so the check cannot catch it. Retype the name or clear the space.

### Amounts typed as text are missing

Every method here skips an amount stored as text, such as "42.50" left-aligned in the cell. The totals stay lower with no error, and **B13** still says OK. Convert those cells to numbers.

### B13 says CHECK after I changed a channel

**B6** has Etsy and Shopify typed inside the formula, so it keeps adding those two. Edit the names inside **B6** to match **B1** and **B2**. A channel with no rows gives the same result, because **B6** still adds the other name.

## Template

The download has three tabs, all portrait for US Letter: **Orders** (24 fictional orders, with a dropdown in the Channel column), **Totals** (the formulas above) and **How to use**. The input cells are yellow, and printing covers the first 40 rows of Orders.

[Download the .xlsx](templates/tidy-tabs-sum-with-or.xlsx)

To use it in Google Sheets, open a blank sheet and choose **File > Import > Upload**, then pick the .xlsx file. The file was recalculated in LibreOffice Calc only and was not opened in Excel or Google Sheets.
