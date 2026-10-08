---
threads_url: https://www.threads.com/@tin_ylab/post/DePbTTPiQ3n
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-sales-tax-in-excel-and.html
title: How to Calculate Sales Tax in Excel and Google Sheets
labels: excel-formulas, sales-tax, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps use the same functions; the ROUND links are from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Calculate sales tax with =ROUND(B2*$F$1,2) and pull tax out of a total with =ROUND(B2-B2/(1+$F$1),2). Free .xlsx template with sample Etsy-style rows.
image_prompts: A wooden craft-fair table with a small cash box holding a few bills, a stack of kraft paper bags tied with twine, and a pad of blank receipt slips beside a pencil, soft daylight, all printed text blurred and unreadable, no screens, no logos
image_alt: Craft-fair table with a cash box, kraft paper bags, and a receipt pad
threads: Sales tax in a sheet is two formulas.\nType your rate once in F1, then =ROUND(B2*$F$1,2) gives the tax on $18.00 at 6.25% as $1.13.\nAlready have a total with tax in it? =ROUND(B2-B2/(1+$F$1),2) pulls the tax back out.
---
Type your tax rate in one cell, such as **F1**, then use `=ROUND(B2*$F$1,2)` to get the tax on the price in **B2**. At a made-up 6.25%, an $18.00 mug has $1.13 of tax, and `=B2+C2` gives a total of $19.13.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help for ROUND. Neither app was opened for this post.

![Sales tax sheet for six items with a 6.25% rate cell, tax amounts and totals](images/2026-10-08-team-01-sales-tax-calculator-excel-google-sheets/sales-tax-calculator-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

The rate is the one number you will change, so it gets its own cell. Every formula points at it.

1. In row 1, type these headers in A to D: Item, Price before tax, Sales tax, Total.
2. In **E1**, type Sales tax rate. In **F1**, type 6.25%. This rate is made up for the example.
3. In **A2**, type the item name. In **B2**, type the price before tax, such as 18.
4. In **C2**, type `=ROUND(B2*$F$1,2)`. The dollar signs lock the rate cell, so it stays put when you copy the formula down.
5. In **D2**, type `=B2+C2` for the total.
6. Format **B2:D2** as currency with two decimals.
7. Select **C2:D2** and fill down for each item.
8. Under the last item, type `=SUM(B2:B7)` and copy it across to column D to total each column.

ROUND matters. At 6.25%, $18.00 is exactly $1.125 of tax, and you cannot collect half a cent. `ROUND(...,2)` rounds it to $1.13. Without it, the cell may display $1.13 while holding 1.125 underneath, and your totals can drift by a cent.

Tax rates vary by state, and often by county and city. Look up yours on your state's own revenue department site, listed on the [state tax agencies page](https://taxadmin.org/state-tax-agencies/). This sheet only does arithmetic and is not tax advice. Ask your state or an accountant what you owe and when.

### Pull the tax out of a total

Sometimes the number you have already includes tax, such as a booth price of $50.00 paid in one amount. To split it, divide by 1 plus the rate. With the total in **B2** and the rate in **F1**:

1. In **C2**, type `=ROUND(B2-B2/(1+$F$1),2)`. This is the tax inside the total.
2. In **D2**, type `=B2-C2`. This is the price before tax.

Do not multiply the total by the rate. That takes 6.25% of a number that already contains the tax, and it comes out too high.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) page.

### In Google Sheets

The same formulas work as typed. Google documents the arguments on its [ROUND function](https://support.google.com/docs/answer/3093440?hl=en) page.

## Example

Sample data below is fictional. It is six Etsy-style items at a 6.25% rate.

| Item | Price before tax | Sales tax | Total |
|---|---|---|---|
| Ceramic mug | $18.00 | $1.13 | $19.13 |
| Sticker pack | $4.50 | $0.28 | $4.78 |
| Canvas tote bag | $24.00 | $1.50 | $25.50 |
| Soy candle, 8 oz | $16.00 | $1.00 | $17.00 |
| Art print, 8 x 10 | $12.99 | $0.81 | $13.80 |
| Beaded earrings | $32.50 | $2.03 | $34.53 |
| Total | $107.99 | $6.75 | $114.74 |

Check two rows by hand. The sticker pack is 4.50 x 0.0625 = 0.28125, which rounds to $0.28. The art print is 12.99 x 0.0625 = 0.811875, which rounds to $0.81.

The totals row adds the rounded amounts: $1.13 + $0.28 + $1.50 + $1.00 + $0.81 + $2.03 = $6.75. That matches what you would collect item by item.

Now the reverse. Feed those same totals into the second tab and the formula returns the original prices: $19.13 gives $1.13 of tax and $18.00 before tax, and $34.53 gives $2.03 and $32.50. A $50.00 tax-inclusive total gives $2.94 of tax and a $47.06 price before tax.

## Troubleshooting

### The tax shows $1.13 but the total is off by a cent
The tax cell is probably not rounded, so the real value is 1.125 and only the display is rounded. Wrap the formula in `ROUND(...,2)` as in step 4.

### The tax is huge, like $112.50 on an $18.00 item
The rate cell holds 6.25 instead of 6.25%. Type the percent sign, or format **F1** as a percentage. Check that **F1** reads 0.0625 underneath.

### The tax changes to zero when I copy the formula down
The rate reference lost its dollar signs and slid down to **F2**, which is empty. Use `$F$1` so the reference is locked.

### Pulling tax out gives a number that is a few cents high
You likely multiplied the total by the rate. Use `=ROUND(B2-B2/(1+$F$1),2)` instead, which divides by 1 plus the rate.

### My state taxes some items and not others
The template applies one rate to every row. If part of your sales is exempt, add a column with 0 or 1 and multiply the tax by it. This was not built into the template or tested here.

## Template

[Download the .xlsx](templates/tidy-tabs-sales-tax-calculator.xlsx)

The file has three tabs. **Sales tax** holds the six sample items with the tax and total formulas, a totals row, and the rate in **F1**. **Tax-inclusive totals** holds seven totals with the pull-the-tax-out formulas, and its rate follows **F1** on the first tab. **How to use** has short fill-in notes.

The items, prices and rate are fictional. The file was recalculated only in LibreOffice Calc 24.2.7 (not in Excel or Google Sheets). To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608?hl=en).
