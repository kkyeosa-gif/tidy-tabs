---
title: How to Look Up a Price by Row and Column With INDEX and MATCH in Excel and Google Sheets
labels: excel-formulas, pricing, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. The template was recalculated there and every result below was checked against the rate card. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Look up a price where a service row meets a speed column with =INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0)). Typo fix and free .xlsx template.
image_prompts: Print shop counter with three paper stock sample fans in different weights, a rubber stamp marked with a plain rush tag and a ruler, soft window light, no screens, no readable text, no logos
image_alt: Paper sample fans, a rush tag stamp and a ruler on a print counter
threads: VLOOKUP only reads one direction, and a rate card has two.\nUse =INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0)): one MATCH finds the row, the other the column.\nBrochure + Rush returns $500.00. Wrap it in IFERROR for typos.
---
Use `=INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0))` to return the price where a service row meets a speed column. The first `MATCH` finds the row for the service in F1 and the second finds the column for the turnaround in G1.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked. The Excel and Google Sheets steps use the same functions and link to Microsoft and Google help pages. Neither Excel nor Google Sheets was opened for this post.

![Rate card with three services and three turnaround speeds beside a quote cell showing the Brochure Rush price](images/2026-10-10-team-04-two-way-rate-card-lookup-index-match-excel-google-sheets/two-way-rate-card-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

A two-way rate card has services down the left and speeds across the top. A one-column lookup cannot read both directions, so you ask `MATCH` for each position and let `INDEX` pick the cell.

1. Type **Service** in **A1**, then **Standard**, **Rush** and **Same day** in **B1:D1**.
2. Type one service per row in **A2:A4**: Logo, Brochure, Website.
3. Type the prices in **B2:D4**. Format them as currency.
4. In **F1**, type the service you want, for example Brochure.
5. In **G1**, type the turnaround, for example Rush.
6. In **H1**, type `=INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0))`.

Here is how the formula reads. `MATCH(F1,A2:A4,0)` returns the position of the service in the list, so Brochure is 2. `MATCH(G1,B1:D1,0)` returns the position of the speed, so Rush is 2. `INDEX(B2:D4,2,2)` returns the cell in row 2, column 2 of the price block.

The `0` at the end of each `MATCH` means exact match. Leave it out and `MATCH` may return a nearby label instead of an error.

### Show a message instead of #N/A

If a label is misspelled, `MATCH` has nothing to find and the formula shows `#N/A`. Wrap it so the cell says what to fix:

`=IFERROR(INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0)),"Check spelling")`

### In Excel

The formula works as typed. Microsoft documents the arguments on its [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) page.

### In Google Sheets

The same formula should work as typed, since Sheets has INDEX, MATCH and IFERROR. Google describes the lookup half on its [MATCH help page](https://support.google.com/docs/answer/3093378). This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is the rate card of a small design freelancer who sells logos, brochures and websites at three speeds.

| Service | Standard | Rush | Same day |
|---|---|---|---|
| Logo | $300.00 | $420.00 | $600.00 |
| Brochure | $350.00 | $500.00 | $700.00 |
| Website | $1,200.00 | $1,600.00 | $2,100.00 |

Pick Brochure and Rush and the formula returns $500.00. That is the cell in the Brochure row and the Rush column.

In the template, the dropdown cells are on a separate Quote tab, in **Quote!B2** (service) and **Quote!B3** (turnaround). The rate card sits on its own tab, so the template formulas point at `'Rate card'!` ranges instead of the plain `A2:A4` used above. The logic is the same.

The check was run for all 9 service and speed pairs. Each was typed into Quote!B2 and Quote!B3, and the INDEX and MATCH cell, the IFERROR cell and a third SUMPRODUCT cell all returned the rate card price every time. A grid on the Quote tab repeats the formula for all 9 cells, and a count cell in **Quote!B14** shows 9 matches.

You can also compare against a different formula: `=SUMPRODUCT((A2:A4=F1)*(B1:D1=G1)*B2:D4)` returns the same number. It is useful as a cross-check, but see the first problem below.

## Troubleshooting

### The cell shows #N/A
One of the two labels is not in the table. "Spa" with Rush gives `#N/A`, and so does Brochure with "Rus". Fix the spelling, or use the IFERROR version above to get "Check spelling" instead.

### SUMPRODUCT shows $0.00 for a typo
The SUMPRODUCT version has no error to show. In the same "Spa" and Rush test it returned $0, which looks like a real price. Use it only to double-check `INDEX` and `MATCH`, not to replace them.

### Lowercase text still works
`MATCH` ignores letter case, so "brochure" and "rush" still return $500.00. This is useful for typing, but do not count on it to catch capitalization mistakes. Extra spaces are a different matter: "Rush " with a trailing space is not "Rush" to the formula.

### The price is in the wrong cell
`INDEX` counts from the first cell of the range you give it. If the range starts at A1 instead of B2, row 2 and column 2 point to the wrong cell. Make the price range exactly the numbers (B2:D4) and keep the label ranges to the same size: three rows for the service list and three columns for the speed list.

### The dropdown lists are empty
The template has dropdowns on Quote!B2 and Quote!B3 that read the labels from the Rate card tab. They were created in the file, but clicking them was not tested; values were typed into the cells for every check above. If a list looks empty in your app, type the label directly.

## Template

[Download the .xlsx](templates/tidy-tabs-two-way-rate-card.xlsx)

The file has three tabs. **Rate card** holds the services, speeds and prices in A1:D4. **Quote** has the two selection cells, the INDEX and MATCH price, the IFERROR price, the SUMPRODUCT cross-check, and the 9-cell grid with its match count. **How to use** has short fill-in notes, including how to add a service or a speed by inserting a row or column inside the table and extending the ranges.

The services and prices are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.

If your lookup only needs one dimension, see the earlier posts on the vlookup-price-list and quantity-price-tiers templates.
