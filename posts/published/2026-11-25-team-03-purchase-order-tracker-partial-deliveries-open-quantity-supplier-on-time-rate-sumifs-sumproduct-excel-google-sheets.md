---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-track-purchase-orders-with.html
title: How to Track Purchase Orders With Partial Deliveries in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 215 workbooks and 871,180 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The COUNTIFS(range,"<>") counting problem was reproduced in LibreOffice Calc only.
search_description: Log each delivery as a row, total it per PO line with SUMIFS, get open quantity with MAX(0,ordered-received), and score suppliers on time and fill rate.
image_prompts: A half-unpacked cardboard shipping carton on a candle workshop floor, showing a single layer of black metal candle tin lids in cardboard dividers with the lower layer visibly empty, next to a sealed bag of soy wax flakes and a bundle of cotton wicks on the workbench, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Half-empty carton of black tin lids beside soy wax flakes and cotton wicks
threads: Log each delivery as its own row and let SUMIFS total it per PO line.\n=MAX(0,ordered-received) gives what is still coming, so an over-delivery never goes negative.\nIn the sample, PO-1002 shows 120 of 240 tin lids received, Late 15 days, $48.00 still open.
---
Log each delivery as its own row, total it into the PO line with `=SUMIFS(qty_received,receipt_po,po,receipt_item,item)`, get what is still coming with `=MAX(0,ordered-received)`, and flag a line late when open is above 0 and your typed as-of date is past the expected date. In the made-up candle supply orders as of 11/20/2026, PO-1002 shows 120 of 240 black tin lids received, Partial, Late 15 days, with $48.00 still open.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was compared there with an independent Python model (215 workbooks, 871,180 cell comparisons, 0 mismatches). The COUNTIFS "<>" counting problem in Troubleshooting was reproduced in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Seven purchase order lines showing received, open quantity, partial, late days and open value](images/2026-11-25-team-03-purchase-order-tracker-partial-deliveries-open-quantity-supplier-on-time-rate-sumifs-sumproduct-excel-google-sheets/purchase-order-tracker-template.png) *PO lines tab of fictional candle supply orders as of 11/20/2026: PO-1002 has 120 of 240 black tin lids received, Partial, Late 15 days with $48.00 still open, the soy wax line is Over by 1, and three completed lines are On time. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The sheet tracks quantities, dates and open value, not payment terms, invoices or three-way matching. For due dates, see [How to Calculate an Invoice Due Date From Payment Terms](https://tidytabs.blogspot.com/2026/10/how-to-calculate-invoice-due-date-from.html).

The steps are the same in Excel and Google Sheets.

1. On the **PO lines** tab, type your as-of date in **B1**, such as 11/20/2026. The file uses no TODAY, so it gives the same answer whenever you open it.
2. From row 5, type one row per item on each PO: PO #, Supplier, Item, Qty ordered, Unit cost, Order date and Expected date in **A:G**.
3. On the **Receipts** tab, type one row per delivery: Date received, PO #, Item and Qty received in **A:D**. Spell the PO # and item exactly as on the PO line.
4. Back on **PO lines**, total the receipts in **H5** with `=SUMIFS(Receipts!$D$5:$D$1004,Receipts!$B$5:$B$1004,A5,Receipts!$C$5:$C$1004,C5)`. The template wraps it in a check that stays blank until the line is filled in. See the [SUMIFS help from Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b).
5. In **I5**, get the open quantity with `=IF(H5="","",MAX(0,D5-H5))`. MAX keeps an over-delivery from going negative. **J5** labels the line Not received, Partial, Received or Over by n.
6. In **K5**, flag lateness with `=IF(OR(I5="",NOT(ISNUMBER($B$1)),NOT(ISNUMBER(G5))),"",IF(I5=0,"",IF($B$1>G5,"Late "&($B$1-G5)&IF($B$1-G5=1," day"," days"),"")))`. Only open lines can be late.
7. In **L5**, find the completion date: `=IF(I5="","",IF(I5>0,"",SUMPRODUCT(MAX((Receipts!$B$5:$B$1004=A5)*(Receipts!$C$5:$C$1004=C5)*Receipts!$F$5:$F$1004))))`. Receipts column **F** holds the date as a number, or 0 when it is blank or text. See the [SUMPRODUCT help from Microsoft](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e).
8. **M5** compares it with the expected date: `=IF(OR(L5="",NOT(ISNUMBER(G5))),"",IF(L5<=G5,"On time","Late"))`.
9. **N5** caps the received quantity with `=IF(H5="","",MIN(H5,D5))`, and **O5** gets the open value with `=IF(I5="","",ROUND(I5*E5,2))`.
10. On the **Suppliers** tab, type each supplier once in **A**. **D5** counts on-time lines with `=IF(TRIM(A5)="","",COUNTIFS('PO lines'!$B$5:$B$204,A5,'PO lines'!$M$5:$M$204,"On time"))`, **E5** counts "Late", **C5** adds them, and **F5** divides **D5** by **C5**. See the [COUNTIFS help from Microsoft](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842).
11. The fill rate in **I5** is the SUMIFS of Received up to ordered (**H5**) divided by the SUMIFS of Qty ordered (**G5**).
12. Fix any row showing Check until **S11** on PO lines and **P7** on Suppliers say OK.

Receipts dated after the as-of date still count as received; the as-of date only decides what is late.

## Example

The sample suppliers, items, prices and dates are made up. As of 11/20/2026:

| PO # | Supplier | Item | Ordered | Expected | Received | Open | Receipt status | Late as of B1 | On time? | Open value |
|---|---|---|---|---|---|---|---|---|---|---|
| PO-1001 | Wick Co | Cotton wick 6in | 500 | 10/30/2026 | 500 | 0 | Received | | On time | $0.00 |
| PO-1001 | Wick Co | Wood wick 5in | 200 | 10/30/2026 | 200 | 0 | Received | | Late | $0.00 |
| PO-1002 | Jar Supply | 8 oz amber jar | 240 | 11/05/2026 | 240 | 0 | Received | | On time | $0.00 |
| PO-1002 | Jar Supply | Black tin lid | 240 | 11/05/2026 | 120 | 120 | Partial | Late 15 days | | $48.00 |
| PO-1003 | Soy Wax Direct | Soy wax 45 lb | 4 | 11/12/2026 | 5 | 0 | Over by 1 | | On time | $0.00 |
| PO-1004 | Jar Supply | 4 oz amber jar | 120 | 11/24/2026 | 0 | 120 | Not received | | | $102.00 |
| PO-1005 | Wick Co | Cotton wick 6in | 300 | 11/19/2026 | 200 | 100 | Partial | Late 1 day | | $8.00 |

The wood wicks came in two deliveries, 150 on 10/29/2026 and 50 on 11/06/2026, so the line finished after its expected date and counts as Late. PO-1004 is open but not due until 11/24/2026. The summary in **R4:S11** shows 7 lines, 3 open, 2 late, $158.00 open value and CHECK OK.

| Supplier | Completed | On time | On-time rate | Ordered | Received up to ordered | Fill rate | Fill rate if not capped | Open value |
|---|---|---|---|---|---|---|---|---|
| Wick Co | 2 | 1 | 50% | 1,000 | 900 | 90% | 90% | $8.00 |
| Jar Supply | 1 | 1 | 100% | 600 | 360 | 60% | 60% | $150.00 |
| Soy Wax Direct | 1 | 1 | 100% | 4 | 4 | 100% | 125% | $0.00 |

## Troubleshooting

### The on-time rate is too low
In LibreOffice Calc, `COUNTIFS(supplier_range,A5,on_time_range,"<>")` counted the open lines' formula blanks ("") as completed: 3 lines for Wick Co instead of 2, so 33% instead of 50%. Count "On time" plus "Late" instead, as the template does.

### A delivery does not show on its line
The PO # or item does not match exactly. A receipt for "Black tin lid " with a trailing space showed Check: no matching PO line and was not added. Upper and lower case did not matter in LibreOffice Calc.

### The fill rate is above 100%
An over-delivery is in the total: uncapped, Soy Wax Direct shows 125% for 5 received against 4 ordered. `MIN(received,ordered)` per line brings it to 100%, and the line still says Over by 1.

### Open quantity is negative
The open formula is missing `MAX(0,...)`. In a template copy without it, the wax line showed -1 and the open value total dropped from $158.00 to $96.00.

## Template

[Download the .xlsx](templates/tidy-tabs-purchase-order-tracker.xlsx)

The file has four tabs. **PO lines** holds the as-of date in **B1** and up to 200 lines: you type **A:G**, and **H:P** hold the formulas, ending in a Status column that flags blanks, a qty that is not a whole number above 0, the same item twice on one PO, or a receipt dated before the order date. **Receipts** takes up to 1,000 deliveries. **Suppliers** shows on-time rate, fill rate (capped and uncapped), open lines and open value for up to 20 suppliers. **How to use** explains each formula.

Yellow cells are inputs, gray cells are formulas, Late turns light orange and every Check turns red. The data tabs print Landscape on US Letter with the header row repeated.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
