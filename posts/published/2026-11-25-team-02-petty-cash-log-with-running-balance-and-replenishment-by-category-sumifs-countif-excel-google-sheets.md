---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-keep-petty-cash-log-and.html
title: How to Keep a Petty Cash Log and Calculate the Replenishment in Excel and Google Sheets
labels: excel-formulas, small-business, small-business-bookkeeping, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 217 workbooks and 338,737 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Petty cash log with a running balance, refill cycles and category totals: refill by fund minus counted cash, not receipts, and record over/short.
image_prompts: A small gray metal petty cash box with its lid open showing a few bills and coins, a short stack of crumpled store receipts and a strip of postage stamps beside it on a wooden small office desk, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Open petty cash box with a few bills, crumpled receipts and a strip of stamps
threads: Refilled petty cash by the receipts total, and now every count starts $0.26 short.\nRefill by fund minus counted cash instead, and log the gap once as over/short.\n$200.00 fund, $41.37 counted: the refill is $158.63, not $158.37, and the box is back at exactly $200.00.
---
Refill petty cash by `=fund-counted`, not by the receipts total, and record the gap with `=ROUND(counted-(fund-spent),2)` as over or short. With a made-up $200.00 fund, $158.37 in receipts this cycle and $41.37 counted in the box, the cash should be $41.63, so you are $0.26 short and the refill is $158.63, which puts the box back at exactly $200.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (217 workbooks, 338,737 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Petty cash categories with cycle totals beside a $158.63 refill and $0.26 short summary](images/2026-11-25-team-02-petty-cash-log-with-running-balance-and-replenishment-by-category-sumifs-countif-excel-google-sheets/petty-cash-replenishment-template.png) *Replenish tab of a fictional $200.00 petty cash fund: $158.37 spent this cycle across Postage, Supplies and Travel, $41.37 counted against $41.63 expected, $0.26 short, so the refill is $158.63, not the receipts total. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

This is a bookkeeping log, not tax advice. Categories are your own labels, not tax categories, and nothing here says what is deductible. The $200.00 fund is a sample, not a suggested size.

The steps are the same in Excel and Google Sheets.

1. On the **Log** tab, type your fund in **B1** ($200.00).
2. On row 5, type the day you filled the box in **A**, "Start fund" in **B**, the cash in **E** (Added) and pick **Start** in **F** (Added as).
3. For each payout, add a row: Date, who you paid, a Category from the dropdown, the amount in **D** (Paid out) and Y or N in **G** (Receipt?). Never fill Paid out and Added on one row.
4. **H** keeps the running balance: `=IF(AND(D5="",E5=""),"",ROUND(SUM($E$5:E5)-SUM($D$5:D5),2))`. More in [How to Add a Running Balance Column](https://tidytabs.blogspot.com/2026/10/how-to-add-running-balance-column-in.html).
5. **I** numbers each refill cycle by counting Replenish rows so far: `=IF(AND(D5="",E5=""),"",COUNTIF($F$5:F5,"Replenish"))`. Rows before the first refill are cycle 0. See the [COUNTIF help from Microsoft](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34).
6. On the **Replenish** tab, **H5** counts every Replenish row to get the current cycle, and **B5** totals one category for that cycle only: `=IF(TRIM(A5)="","",SUMIFS(Log!$D$5:$D$504,Log!$C$5:$C$504,A5,Log!$I$5:$I$504,$H$5))`. See the [SUMIFS help from Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b).
7. When the box runs low, count the cash and type it in **Log B2**. The Replenish tab then shows Cash expected in the box `=ROUND(H4-H7,2)` (fund minus spent this cycle), Over (+) or short (-) `=IF(H10="","",ROUND(H10-H9,2))` and the Replenish amount `=IF(H10="","",ROUND(H4-H10,2))` in green.
8. Log the gap first as a payout in the **Over/short** category ($0.26 here; a negative amount for cash over). This keeps the log balance equal to the box. For counting cash, see [How to Count a Cash Drawer and Find Over or Short](https://tidytabs.blogspot.com/2026/10/how-to-count-cash-drawer-and-find-over.html).
9. Add the refill row: the Replenish amount in **E** and **Replenish** in **F**. That row starts the next cycle. Count again and type the new count in **B2**.
10. Check that **CHECK** (Replenish **H20**, also **Log M7**) says OK.

## Example

The sample payees and amounts are made up. Fund $200.00, cash counted $41.37 on 11/24/2026.

| Date | Paid to / note | Category | Paid out | Added | Added as | Balance | Cycle |
|---|---|---|---|---|---|---|---|
| 11/02/2026 | Start fund | | | $200.00 | Start | $200.00 | 0 |
| 11/03/2026 | Post office | Postage | $18.40 | | | $181.60 | 0 |
| 11/05/2026 | Hardware store | Supplies | $23.17 | | | $158.43 | 0 |
| 11/09/2026 | Coffee for market crew | Meals | $14.60 | | | $143.83 | 0 |
| 11/12/2026 | Post office | Postage | $9.85 | | | $133.98 | 0 |
| 11/14/2026 | Parking | Travel | $12.00 | | | $121.98 | 0 |
| 11/16/2026 | Refill from bank | | | $78.02 | Replenish | $200.00 | 1 |
| 11/18/2026 | Craft store | Supplies | $31.49 | | | $168.51 | 1 |
| 11/19/2026 | Post office | Postage | $27.30 | | | $141.21 | 1 |
| 11/20/2026 | Ice for cooler | Supplies | $6.99 | | | $134.22 | 1 |
| 11/21/2026 | Parking | Travel | $15.00 | | | $119.22 | 1 |
| 11/23/2026 | Tape and labels | Supplies | $22.84 | | | $96.38 | 1 |
| 11/24/2026 | Post office | Postage | $54.75 | | | $41.63 | 1 |

The Replenish tab for cycle 1, which started 11/16/2026:

| Line | Value |
|---|---|
| Postage (2 payouts) | $82.05 |
| Supplies (3 payouts) | $61.32 |
| Travel (1 payout) | $15.00 |
| Spent this cycle | $158.37 |
| Cash expected in the box | $41.63 |
| Cash counted | $41.37 |
| Over (+) or short (-) | -$0.26 |
| Replenish amount | $158.63 |
| Payouts without a receipt this cycle | 1 |
| CHECK | OK |

The one payout without a receipt this cycle is the ice on 11/20/2026; the coffee on 11/09/2026 belongs to cycle 0. Meals shows $0.00 this cycle and $14.60 across all cycles.

## Troubleshooting

### The box never gets back to the full fund
You refilled by the receipts total: $158.37 plus $41.37 counted is $199.74, and every later count starts $0.26 short. Refill by fund minus counted ($158.63) and log the $0.26 once as Over/short.

### A payout is missing from the category totals
The row has no category or one not on the Replenish list (such as "Snacks"). The row shows Check and Spent outside the listed categories shows the amount. Pick a listed category or add yours to **Replenish A5:A14**.

### The cycle did not restart after a refill
Added as says something else, such as "Refill". The row shows Check and does not count as a refill. Pick **Replenish** from the dropdown.

### CHECK says Check after you raised the fund
You logged a **Fund change** row but left **Log B1** at the old amount, so Log balance minus cash expected shows $50.00 for a $50.00 raise. Update B1.

### The Replenish amount is blank or negative
**Log B2** is blank or text, so the count, over/short and refill cells stay blank and CHECK says Check. A count above the fund gives a negative refill and "Check: fund or count (Log B1:B2)".

## Template

[Download the .xlsx](templates/tidy-tabs-petty-cash-log.xlsx)

The file has three tabs. **Log** holds the fund and count in **B1:B2** and rows 5 to 504: you type **A:G**, and **H:J** hold the balance, cycle and a Status that names each problem (no amount, both columns filled, more than 2 decimals, no category, balance below $0.00 and more). **L4:M7** repeats cash expected, over or short, the replenish amount and CHECK.

**Replenish** lists the categories in **A5:A14** (Postage, Supplies, Meals, Travel, Other and Over/short, all editable) with this cycle's spending and payouts, plus the summary in **G4:H20**. **How to use** walks through the formulas and the refill steps.

Yellow cells are inputs and gray cells are formulas. Replenish rows and the replenish amount are light green, and every Check turns red. Pages print on US Letter, fit one page wide, with the header row repeated; the data tabs print Landscape.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
