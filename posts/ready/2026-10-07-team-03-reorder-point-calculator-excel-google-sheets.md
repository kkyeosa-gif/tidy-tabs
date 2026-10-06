---
title: How to Calculate a Reorder Point in Excel and Google Sheets
labels: excel-formulas, inventory, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same functions; the ROUNDUP links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Calculate a reorder point with =ROUNDUP(C2*D2+E2,0) and flag low stock with =IF(G2<=F2,"REORDER","OK"). Free .xlsx template with six sample supplies.
image_prompts: Shelves in a small candle workshop holding a nearly empty box of glass jars, a half-full bag of soy wax, and a spool of ribbon, with a clipboard hanging on the side, soft light, all paper text blurred and unreadable, no screens, no logos
image_alt: Workshop shelf with a nearly empty box of glass jars, wax bag, and ribbon spool
threads: Reorder point = daily sales x lead time + safety stock.\n=ROUNDUP(C2*D2+E2,0) turns 90 units a month, 7 lead days and 6 spare into 27.\n=IF(G2<=F2,"REORDER","OK") flags the rows that need an order now.
---
Use `=ROUNDUP(C2*D2+E2,0)`: daily sales times lead time in days, plus safety stock. 90 units sold in 30 days is 3 a day, so 7 lead days and 6 safety stock gives a reorder point of 27.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help for ROUNDUP. Neither app was opened for this post.

![Reorder sheet for six supplies with reorder points and REORDER flags on the low-stock rows](images/2026-10-07-team-03-reorder-point-calculator-excel-google-sheets/reorder-point-calculator-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Set up one row per supply. The columns below match the template.

1. In row 1, type these headers in A to H: Supply, Units sold (last 30 days), Daily sales, Lead time (days), Safety stock, Reorder point, On hand, Flag.
2. In **K1**, type 30. This is the number of days your units-sold count covers.
3. In row 2, type the supply name in **A2**, units sold in the last 30 days in **B2**, days from placing an order to getting it in **D2**, your safety stock in **E2**, and the count on your shelf in **G2**.
4. In **C2**, type `=B2/$K$1` for daily sales. The dollar signs keep the 30 fixed when you copy the formula down.
5. In **F2**, type `=ROUNDUP(C2*D2+E2,0)`. This is what you expect to sell while the order is on its way, plus your cushion.
6. In **H2**, type `=IF(G2<=F2,"REORDER","OK")`. It says REORDER when on hand is at or below the reorder point.
7. Format **C2** with two decimals so 1.1 units a day shows as 1.10.
8. Select **C2**, **F2** and **H2** and fill them down for each supply.

ROUNDUP is deliberate. A reorder point of 19.4 units means you need 20, because you cannot sell a fraction of a jar.

Safety stock is your own call. It is the spare units you want on hand in case a delivery runs late or a week sells faster than usual. The sheet only does the arithmetic.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) page.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [ROUNDUP function](https://support.google.com/docs/answer/3093444) page.

## Example

Sample data below is fictional. It is six supplies for a small candle shop.

| Supply | Units sold (30 days) | Daily sales | Lead time (days) | Safety stock | Reorder point | On hand | Flag |
|---|---|---|---|---|---|---|---|
| Glass jars, 8 oz | 90 | 3.00 | 7 | 6 | 27 | 40 | OK |
| Soy wax, 1 lb bags | 60 | 2.00 | 10 | 8 | 28 | 28 | REORDER |
| Candle wicks, pack of 50 | 33 | 1.10 | 14 | 4 | 20 | 21 | OK |
| Cotton ribbon spools | 75 | 2.50 | 5 | 10 | 23 | 18 | REORDER |
| Kraft mailer boxes | 120 | 4.00 | 12 | 15 | 63 | 80 | OK |
| Shipping label rolls | 45 | 1.50 | 21 | 5 | 37 | 30 | REORDER |

Check the first row by hand: 90 / 30 = 3 a day, 3 x 7 = 21, and 21 + 6 = 27. With 40 jars on hand, the flag says OK.

The soy wax row sits exactly on the line. On hand is 28 and the reorder point is 28, and the flag says REORDER. That is what `<=` does, and it is on purpose: at the reorder point you order now, not one sale later.

Three rows round up. Wicks are 1.10 x 14 + 4 = 19.4, which becomes 20. Ribbon is 2.50 x 5 + 10 = 22.5, which becomes 23. Labels are 1.50 x 21 + 5 = 36.5, which becomes 37.

## Troubleshooting

### The flag says REORDER when on hand equals the reorder point
This is how the formula is written. `<=` includes the tie. Change it to `<` in **H2** if you would rather wait until stock drops below the point.

### Daily sales shows 3 instead of 3.00, or 1.1 instead of 1.10
The cell has General formatting, which drops trailing zeros. Format column C with two decimals. The value underneath is the same.

### The reorder point is one unit higher than my hand math
ROUNDUP always goes up when there is any fraction. 19.4 becomes 20, not 19. Use `ROUND` instead if you want normal rounding, but the sheet then plans for slightly less than you expect to sell.

### The numbers stop updating when I change the 30
Daily sales divides by **K1**. If you count sales over 60 days, type 60 in **K1** and use units sold over 60 days in column B. This was not reproduced here, so check one row by hand after changing it.

### The flag shows an error
Check that columns B, D, E and G hold numbers and not text, and that **K1** is not empty or 0. Text-formatted numbers were not tested in this template.

## Template

[Download the .xlsx](templates/tidy-tabs-reorder-point-calculator.xlsx)

The file has two tabs. **Reorder** holds the six sample supplies with the daily sales, reorder point and flag formulas in columns C, F and H, the 30-day setting in **K1**, and conditional formatting that turns a REORDER row red in the flag column. **How to use** has short fill-in notes and the 90-unit example.

If you also track stock movements, the [craft inventory tracker template](templates/tidy-tabs-craft-inventory-tracker.xlsx) has its own Reorder at column and status flag. This file is a separate, formula-only sheet for working out the number itself.

The supplies and numbers are fictional. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
