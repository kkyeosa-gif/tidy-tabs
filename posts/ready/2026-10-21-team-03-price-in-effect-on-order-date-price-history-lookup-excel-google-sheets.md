---
title: How to Look Up the Price That Was in Effect on a Certain Date in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model for 30 sample orders and many edited variants (order on the change day, the day before, before the first price, unknown product, trailing space, same-day price rows, unsorted dates, text dates, blank prices). Excel and Google Sheets were not opened; the LOOKUP form was not run in Google Sheets. Help links are from Microsoft Support and Google Docs Editors Help.
search_description: Look up the price in effect on a date with LOOKUP or INDEX and MATCH in Excel and Google Sheets. Free price history template with order date checks.
image_prompts: A wholesale coffee bag with a folded kraft paper price list beside a brass scoop and a roll of twine on a roasting shop counter, soft window light, no screens, no readable text, no logos
image_alt: Burlap coffee bag, folded price sheet and brass scoop on a roaster's counter
threads: 28 of 30 orders got a price. 2 say "No price yet".\nThat is on purpose: they come before the product's first effective date. =LOOKUP(2,1/((Prices!$A$2:$A$40=B2)*(Prices!$B$2:$B$40<=A2)),Prices!$C$2:$C$40) does the rest.
---
Look up the price in effect on an order date with `=LOOKUP(2,1/((Prices!$A$2:$A$40=B2)*(Prices!$B$2:$B$40<=A2)),Prices!$C$2:$C$40)`, where A2 is the order date and B2 the product. It returns the price from the latest effective date on or before the order date.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post, and the LOOKUP form was not run in Google Sheets.

![Orders sheet with prices in effect beside a product price history list](images/2026-10-21-team-03-price-in-effect-on-order-date-price-history-lookup-excel-google-sheets/price-in-effect-on-order-date-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

Keep one row per price change on a **Prices** tab: Product in column A, Effective date in column B, New price in column C. Put the order list on an **Orders** tab: Order date in A, Product in B, Quantity in C.

1. On **Prices**, type each price change in rows 2 to 40. Type dates as MM/DD/YYYY, for example 03/01/2026, so they are real dates and not text.
2. On **Orders**, type the order date in **A2** and the product in **B2**.
3. In **D2**, type `=LOOKUP(2,1/((Prices!$A$2:$A$40=B2)*(Prices!$B$2:$B$40<=A2)),Prices!$C$2:$C$40)`.
4. Fill the formula down. In **E2**, type `=C2*D2` for the line total.

How it works: the two tests multiply to 1 where the product matches and the effective date is on or before the order date, and to 0 everywhere else. Dividing 2 by that gives a #DIV/0! error on every row that does not qualify. LOOKUP skips errors and takes the last row that qualifies, which is the latest price in effect. An order dated on the change day gets the new price.

The template's version is a little longer:

`=IFERROR(LOOKUP(2,1/((Prices!$A$2:$A$40=B2)*(Prices!$B$2:$B$40<=A2)*ISNUMBER(Prices!$B$2:$B$40)*ISNUMBER(Prices!$C$2:$C$40)),Prices!$C$2:$C$40),"No price yet")`

The IFERROR shows "No price yet" instead of #N/A. The two ISNUMBER tests skip any price row with a blank or text date or price. Without them, LibreOffice returned an empty-looking 0 for a blank price.

### In Excel

Type the formulas as shown. LOOKUP is documented on Microsoft's [LOOKUP function page](https://support.microsoft.com/en-us/office/lookup-function-446d94af-663b-451d-8251-369d5e3864cb). Excel was not opened for this post.

### In Google Sheets

The same formula should work as typed, and LOOKUP appears on Google's [function list](https://support.google.com/docs/table/25273). The LOOKUP form was not run in Sheets for this post, so check a few results against your price list.

### Simpler version for one product

If the list holds one product with dates sorted oldest to newest, use INDEX and MATCH. With the Cold Brew Pack dates in Prices B2:B4 and prices in C2:C4, and an order date in **I10**:

`=INDEX(Prices!$C$2:$C$4,MATCH(I10,Prices!$B$2:$B$4,1))`

The 1 means approximate match: MATCH finds the last date that is not after I10. This form only works when the dates are sorted oldest first. INDEX is explained on Microsoft's [INDEX function page](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd).

## Example

Sample data is fictional: a coffee supplier's wholesale prices per 5 lb bag, 16 price rows and 30 orders.

| Product | Order date | Price in effect |
|---|---|---|
| Dark Roast | 02/28/2026 | $43.00 |
| Dark Roast | 03/01/2026 | $45.50 |
| House Blend | 06/30/2025 | $42.00 |
| House Blend | 07/01/2025 | $44.50 |
| Decaf Blend | 08/31/2025 | $48.00 |
| Decaf Blend | 09/01/2025 | $49.50 |
| Decaf Blend | 10/12/2026 | $52.50 |

Dark Roast changed from $43.00 to $45.50 on 03/01/2026. An order on that day gets $45.50, and one the day before gets $43.00. The 10/12/2026 Decaf Blend order is later than every row, so it gets the latest price.

Of the 30 orders, 28 are priced and 2 show "No price yet": Single Origin Ethiopia on 03/15/2025 comes before its first price on 06/01/2025, and Cold Brew Pack on 11/20/2025 comes before 01/01/2026. The line totals add up to $7,087.50. The INDEX and MATCH block gives $38.00 for a Cold Brew Pack order on 08/15/2026.

The summary cell **I7** shows OK when every order row is either priced, "No price yet" or flagged, and no price row needs a check.

## Troubleshooting

### The cell says "No price yet"

Either the order date is before the product's first effective date, or the product name does not match. "Dark Roast " with a trailing space counts as a different product than "Dark Roast". Capitals do not matter: "dark roast" is priced.

### Which price does an order on the change day get?

The new one. The formula tests the effective date with `<=`, so a price that starts on 03/01/2026 applies to an order dated 03/01/2026.

### INDEX and MATCH returns the wrong price

The dates are not sorted oldest first. Approximate MATCH assumes sorted data and has no fixed answer otherwise. With the first product's dates out of order, the block showed "No price yet" for 08/15/2026 when the right price was $38.00. The cell **I12** shows "Not sorted" when this happens. The LOOKUP form still gave the right price in the same test.

### Two prices share the same day

If two rows for one product have the same effective date, the lower row wins. The **Prices** tab shows "Same day" on that lower row so you can delete the one you do not want.

### The price shows 0 or a price row is ignored

A price or effective date in **Prices** is blank or stored as text, for example a date typed as text. The row shows "Check" in column D and the lookup skips it. Retype the value as a real date or number.

## Template

The download has three tabs: **Prices** (16 sample price changes), **Orders** (30 sample orders, the lookup, line totals, a summary and the INDEX and MATCH block) and **How to use**. Yellow cells are inputs. Prices takes up to 39 rows and Orders up to 199.

[Download the .xlsx](templates/tidy-tabs-price-by-effective-date.xlsx)

To open it in Google Sheets, choose **File > Import > Upload**, select the file and import it as a new spreadsheet.
