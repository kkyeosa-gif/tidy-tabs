---
threads_url: https://www.threads.com/@tin_ylab/post/DeSADFlmNvw
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-look-up-wholesale-price-tiers-by.html
title: How to Look Up Wholesale Price Tiers by Quantity in Excel and Google Sheets
labels: excel-formulas, vlookup, wholesale-pricing
tested_in: LibreOffice Calc 24.2.7 (Linux) only. The Excel and Google Sheets steps use the same functions; the VLOOKUP and IFERROR help links are from Microsoft Support and Google Docs Editors Help (not tested in those apps).
search_description: Look up a wholesale unit price by quantity with =VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE). Sorted tier table, IFERROR for low quantities, free .xlsx template.
image_prompts: A wholesale price list built from physical goods on a workshop shelf: four stacks of small kraft gift boxes in rising quantities, each tied with a different colored twine, and a clipboard with a blank ruled order sheet beside a brass pricing tag, soft window light, all printed text blurred and unreadable, no screens, no logos
image_alt: Four stacks of kraft gift boxes tied with colored twine beside a clipboard and brass tag
threads: Wholesale price breaks without a nest of IFs.\nSort your tiers by minimum quantity, then =VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE) returns $10.50 for 35 units and $9.25 for 36.\nThe TRUE at the end is the whole trick. Leave the tiers unsorted and you get a wrong price or #N/A.
---
Use an approximate-match lookup on a sorted tier table: `=VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE)` returns the unit price for the last tier whose minimum quantity is less than or equal to the order quantity. `LOOKUP` and `INDEX` with `MATCH(...,1)` give the same price.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and the results below were recalculated. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Wholesale price tier sheet with eight order quantities, unit prices and order totals](images/2026-10-09-team-01-quantity-price-tiers-lookup-excel-google-sheets/quantity-price-tiers-lookup-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

Wholesale price lists usually work in breaks: the more you buy, the lower the unit price. A lookup against a small tier table replaces a long nested IF.

1. Make a sheet named **Tiers**. In **A1:B1**, type Minimum quantity and Unit price.
2. In **A2:B5**, enter the tiers from smallest to largest: 1 / $12.00, 12 / $10.50, 36 / $9.25, 72 / $8.00. Each number in column A is the first quantity that earns that price.
3. Make a second sheet named **Orders**. Put the customer or item in column A, and the order quantity in **C2**.
4. In **D2**, type `=VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE)`. The 2 returns the price column, and `TRUE` asks for the closest minimum that does not go over.
5. In **E2**, type `=C2*D2` for the order total.
6. Format **D2:E2** as currency with two decimals, select them, and fill down.

The dollar signs lock the tier table so it does not slide when you copy the formula down. The tier minimums must be sorted from smallest to largest. Approximate match depends on that order.

Two other formulas return the same price. `=LOOKUP(C2,Tiers!$A$2:$A$5,Tiers!$B$2:$B$5)` takes the minimums and the prices as separate ranges. `=INDEX(Tiers!$B$2:$B$5,MATCH(C2,Tiers!$A$2:$A$5,1))` finds the tier position, then reads the price. The `1` in `MATCH` is the approximate-match setting. Pick whichever you find easiest to read.

If your price list is by exact product code instead of by quantity, use the exact-match version in [the VLOOKUP price list post](posts/published/2026-09-30-team-05-vlookup-price-list-excel-google-sheets.md).

### Handle quantities below the first tier

If someone enters 0, VLOOKUP has no tier to return and shows #N/A. Wrap it so the cell says what happened:

`=IFERROR(VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE),"Below first tier")`

This catches every error, not only the low quantity, so a typo in the formula would show the same message. Check the result when you first set it up.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [VLOOKUP function](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1) and [IFERROR function](https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611) pages.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [VLOOKUP help page](https://support.google.com/docs/answer/3093318?hl=en).

## Example

Sample data below is fictional. The tiers are 1+ at $12.00, 12+ at $10.50, 36+ at $9.25 and 72+ at $8.00.

| Quantity | Unit price | Order total |
|---|---|---|
| 1 | $12.00 | $12.00 |
| 11 | $12.00 | $132.00 |
| 12 | $10.50 | $126.00 |
| 35 | $10.50 | $367.50 |
| 36 | $9.25 | $333.00 |
| 71 | $9.25 | $656.75 |
| 72 | $8.00 | $576.00 |
| 500 | $8.00 | $4,000.00 |

The rows next to each break are the ones to check. At 11 units you pay $12.00, and at 12 you pay $10.50. At 35 you pay $10.50, and one more unit drops the price to $9.25. A 72-unit order costs $576.00, which is less than the $656.75 for 71 units.

In the template, the VLOOKUP, LOOKUP and INDEX/MATCH columns agreed on all eight rows in LibreOffice Calc. Quantity 0 gave #N/A with plain VLOOKUP and "Below first tier" with the IFERROR version.

## Troubleshooting

### The formula returns a wrong price or #N/A
The tier table is probably not sorted. The **Unsorted demo** tab holds the same tiers in the order 36, 1, 72, 12. In LibreOffice Calc it returned #N/A for quantity 5 and 20, $12.00 for 40 (the correct price is $9.25) and $10.50 for 80 (the correct price is $8.00). Sort column A smallest to largest. Excel and Sheets were not tested on this, so treat any wrong price or #N/A the same way.

### Every price is the first tier, or the last tier
The last argument may be `FALSE`, or the quantities may be text. `FALSE` asks for an exact match, which only fits a quantity that appears in the tier table. Set it to `TRUE`. Also check that the quantities are numbers and not text such as "'12", which usually sit left-aligned in the cell.

### A quantity of 0 shows #N/A
The smallest minimum is 1, so 0 is below every tier. Use the IFERROR version above, or add a tier that starts at 0 if you want a price there.

### The price is right but the tier table broke when I copied the formula
The table reference lost its dollar signs and moved down a row, so the lookup range shrinks. Use `Tiers!$A$2:$B$5` with the dollar signs.

### I added a fifth tier and it is ignored
The formula still points at rows 2 to 5. Extend the range to `Tiers!$A$2:$B$6`, or insert the new row inside the table so the range grows with it. Inserting inside the range was not tested here.

## Template

[Download the .xlsx](templates/tidy-tabs-quantity-price-tiers.xlsx)

The file has four tabs. **Tiers** holds the four price breaks and a small test cell: type a quantity in **F2** and **F3** shows the price, or "Below first tier". **Orders** holds eight sample quantities with the VLOOKUP, LOOKUP and INDEX/MATCH prices side by side and an order total. **Unsorted demo** shows what an unsorted table does. **How to use** has short fill-in notes.

The quantities and prices are fictional. The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
