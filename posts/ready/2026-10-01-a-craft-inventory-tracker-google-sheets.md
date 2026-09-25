---
title: Google Sheets Inventory Tracker for a Small Craft Business
labels: google-sheets-templates, inventory
tested_in: LibreOffice Calc 24.2.7 (Linux), 09/25/2026 (template formulas recalculated with the sample data). Google Sheets steps from Google Docs Editors Help pages.
image_prompts: A tidy craft workbench with stacked soy candles, wrapped soap bars, and small kraft boxes beside a laptop showing an abstract blank grid, flat illustration | Shelves of handmade goods with a few empty spots and a clipboard with a blank checklist, warm flat illustration
image_alt: Illustration of a craft workbench with candles, soap, and shipping boxes next to a laptop | Illustration of shelves of handmade goods with a blank clipboard
search_description: Track craft inventory in Google Sheets with two tabs: Items and Stock log. On hand = Starting qty + Qty in - Qty out, with a REORDER flag.
---
Track craft inventory with two tabs: **Items** (one row per product or supply) and **Stock log** (one row per stock movement). On hand is `=F2+G2-H2` (Starting qty + Qty in - Qty out), and Status flips to REORDER once On hand drops to the Reorder at number.

> Works in: Google Sheets (steps from Google Docs Editors Help). Tested in: LibreOffice Calc 24.2.7 only, where the template's formulas recalculated to the results shown below.

## Steps
1. Download the template from the Template section below.
2. In Google Sheets, click **File > Import**, upload the .xlsx, and pick an import option such as **Create new spreadsheet**. Google lists every option on its [import data sets and spreadsheets](https://support.google.com/docs/answer/40608) page.
3. On the **Items** tab, type over the sample rows with your own products and supplies: SKU, Item, Category, Unit cost, Sale price, Starting qty, and Reorder at.
4. Keep SKUs short, unique, and free of spaces, like `CND-LAV8` or `SOP-OAT`. The Stock log matches on the SKU text exactly.
5. On the **Stock log** tab, add one row for every batch you make, every order you ship, and every market day. Pick the SKU from the dropdown in column B.
6. Put incoming stock in **Qty in** and outgoing stock in **Qty out**. Leave the other one at 0 or blank.
7. Back on **Items**, read **Status**. REORDER means On hand is at or below your Reorder at number.
8. Read **Stock value (cost)** for each row. The **Total stock value** for everything is in cell O1, to the right of the table.
9. To add a product, copy the last product row and paste it into the next empty row, then type over the white cells. The gray formula cells come along with it.

The total, the Stock log dropdown, and the red REORDER highlight all cover rows 2 through 500 of Items, so a new row is picked up without editing any formula.

### Build it yourself
If you'd rather start from a blank sheet, use these column letters on Items: A SKU, B Item, C Category, D Unit cost, E Sale price, F Starting qty, G Qty in, H Qty out, I On hand, J Reorder at, K Status, L Stock value (cost). Stock log uses A Date, B SKU, C Qty in, D Qty out, E Note.

Row 2 formulas on Items, copied down:

- Qty in: `=SUMIFS('Stock log'!$C:$C,'Stock log'!$B:$B,A2)`
- Qty out: `=SUMIFS('Stock log'!$D:$D,'Stock log'!$B:$B,A2)`
- On hand: `=F2+G2-H2`
- Status: `=IF(A2="","",IF(I2<=J2,"REORDER","OK"))`
- Stock value: `=I2*D2`

For the SKU dropdown, select Stock log column B from row 2 down, click **Data > Data validation > Add rule**, choose **Dropdown from a range**, and point it at your SKU cells on Items. Google's [dropdown help page](https://support.google.com/docs/answer/186103) walks through the same dialog.

To make REORDER stand out, select the Status cells, click **Format > Conditional formatting**, set **Format cells if** to **Custom formula is**, type `=$K2="REORDER"`, pick a red fill, and click **Done**. See Google's [conditional formatting guide](https://support.google.com/docs/answer/78413).

## Example
The sample shop sells soy candles, goat milk soap, brass earrings, and letterpress cards, and tracks shipping boxes as a supply. Here is how four rows work out after the September log entries:

| Item | Starting qty | Qty in | Qty out | On hand | Reorder at | Status |
| --- | --- | --- | --- | --- | --- | --- |
| Lavender soy candle, 8 oz | 40 | 24 | 12 | 52 | 10 | OK |
| Charcoal soap bar | 30 | 0 | 14 | 16 | 20 | REORDER |
| Brass hoop earrings | 15 | 0 | 11 | 4 | 5 | REORDER |
| Kraft mailer box 6x6x4 | 100 | 0 | 64 | 36 | 40 | REORDER |

The lavender candle got a poured batch of 24 on 09/01/2026 and lost 12 to Etsy orders, so it sits at 52. The charcoal soap went out in one wholesale order of 14 and dropped below its reorder point of 20. Total stock value at cost for all seven items comes to $475.12. We checked these results in LibreOffice Calc 24.2.

![Items tab with seven sample products, three marked REORDER](images/2026-10-01-a-craft-inventory-tracker-google-sheets/items-tab-libreoffice.png) *Items tab with the sample data. PDF export from LibreOffice Calc 24.2.*

![Stock log tab with dated stock-in and stock-out rows](images/2026-10-01-a-craft-inventory-tracker-google-sheets/stock-log-tab-libreoffice.png) *Stock log tab. Every row is one batch, sale, or market day. PDF export from LibreOffice Calc 24.2.*

Stock value is simply unit cost times On hand. Use it as a quick check on how much money is sitting on your shelves, not as a figure for your taxes.

## Troubleshooting
### On hand doesn't change after I log a sale
The SKU in the Stock log row doesn't match the Items SKU, often because of a typo or an extra space. Pick the SKU from the dropdown instead of typing it.

### On hand is negative
You logged more going out than you ever recorded coming in. Add the missing Qty in row for that batch, or fix the Starting qty on Items.

### A new product row shows blanks
The formulas weren't copied into the new row. Copy the gray cells (Qty in through Stock value) from the row above. Past row 500, also widen the ranges in the O1 total, the dropdown rule, and the REORDER highlight.

### Supplies show REORDER even though I don't sell them
That's working as intended. Boxes, wicks, and ear wires run out too, and the flag tells you to order more before a busy weekend.

### The sheet gets slow with thousands of log rows
The SUMIFS formulas read whole columns (`$C:$C`). Switch to bounded ranges such as `'Stock log'!$C$2:$C$5000` and `'Stock log'!$B$2:$B$5000`, and copy the new formula down.

## Template
[Download the craft inventory tracker (.xlsx)](templates/tidy-tabs-craft-inventory-tracker.xlsx)

The file has three tabs: **Items**, **Stock log**, and **How to use**. Gray columns are formulas, so type only in the white ones. The SKU dropdown on Stock log covers rows 2 through 1000. The shop, products, and log entries are fictional sample data; delete the Stock log rows and type over the Items rows to start fresh.

To open it in Google Sheets, click **File > Import** and upload the .xlsx.
