---
threads_url: https://www.threads.com/@tin_ylab/post/DeM2fh5lJIG
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-sale-price-and.html
title: How to Calculate a Sale Price and Discount in Excel and Google Sheets
labels: excel-formulas, pricing, small-business-finance
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets were not opened for this post.
search_description: Calculate a sale price with =B2*(1-C2), find dollars saved, work back to the original price, and see why 20% then 10% off is not 30%. Free .xlsx.
image_prompts: A rack of price tags and a small handwritten sale sign next to two stacked candle jars and a bar of wrapped soap on a craft-fair table, soft daylight, all tag and sign text blurred and unreadable, no screens, no logos
image_alt: Craft-fair table with candle jars, wrapped soap, and a rack of price tags
threads: Is 20% off then 10% off really 30% off?\nNope. =B8*(1-B9)*(1-B10) turns $100.00 into $72.00.\nAdding the percents gives $70.00, so the real discount is 28%.
---
Multiply the original price by one minus the discount: `=B2*(1-C2)`. A $48.00 item at 25% off comes to $36.00.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated. Excel and Google Sheets use the same formulas, but neither app was opened for this post, so the menu steps for them are plain descriptions, not tested ones.

![Sale prices sheet with eight sample items, percent off, sale price, and dollars saved](images/2026-10-07-team-01-sale-price-discount-calculator-excel-google-sheets/sale-price-discount-calculator-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

## Steps

The template keeps one item per row: the original price in column B and the percent off in column C. Build the same layout like this.

1. In row 1, type these headers: Item, Original price, Percent off, Sale price, You save.
2. In **A2**, type an item name. In **B2**, type the original price, such as 48.
3. In **C2**, type the discount as a percent, such as 25%.
4. In **D2**, type `=B2*(1-C2)`. This keeps the part of the price you still charge: 100% minus 25% is 75%.
5. In **E2**, type `=B2-D2` to show the dollars saved.
6. Select **D2:E2** and drag the fill handle down to copy the formulas to the other rows.

### In Excel

Select the cells in column B, D and E, then pick the Currency or Accounting style from the Home tab. Select column C and click the Percent Style button in the Number group. Excel then reads 25 typed into that column as 25%.

### In Google Sheets

Select column C and choose **Format > Number > Percent**. For the dollar columns, choose **Format > Number > Currency**. The formulas work as typed.

## Work back to the original price

If you only know the sale price and the percent off, divide instead of multiply: `=D2/(1-C2)`.

A $36.00 tag with 25% off means the original price was $48.00. The sale price is 75% of the original, so dividing by 0.75 undoes it. Do not add 25% to $36.00. That gives $45.00, which is wrong.

## Two discounts in a row

20% off, then a 10% coupon, is not 30% off. The second discount applies to the already lower price. Multiply the two keep-fractions: `=B8*(1-B9)*(1-B10)`.

Adding the percents first, `=B8*(1-(B9+B10))`, gives the wrong answer. The template shows both side by side.

## Example

The sample items below are fictional Etsy shop products. Prices and discounts are made up.

| Item | Original price | Percent off | Sale price | You save |
|---|---|---|---|---|
| Lavender soy candle, 8 oz | $18.00 | 10% | $16.20 | $1.80 |
| Cedar soy candle, 8 oz | $18.00 | 15% | $15.30 | $2.70 |
| Oatmeal goat milk soap bar | $8.00 | 25% | $6.00 | $2.00 |
| Charcoal soap bar | $8.50 | 10% | $7.65 | $0.85 |
| Brass hoop earrings | $24.00 | 15% | $20.40 | $3.60 |
| Letterpress birthday card | $6.00 | 25% | $4.50 | $1.50 |
| Beaded necklace | $48.00 | 25% | $36.00 | $12.00 |
| Candle and soap gift set | $32.00 | 10% | $28.80 | $3.20 |

The totals block in **H1:H3** shows $162.50 at original prices, $134.85 at sale prices and $27.65 saved.

The stacked case, with a $100.00 original price:

| Method | Formula | Price |
|---|---|---|
| 20% off, then 10% off | `=B8*(1-B9)*(1-B10)` | $72.00 |
| Add the percents (30% off) | `=B8*(1-(B9+B10))` | $70.00 |

The two answers differ by $2.00. The real total discount is 28%, not 30%. The hand calculation matches the recalculated LibreOffice file.

## Troubleshooting

### The sale price is a huge negative number

The percent column holds 25 instead of 25%. In a General cell, 25 means 25, so the formula becomes `=48*(1-25)` and returns -$1,152.00. Format column C as a percent, then retype the value as 25%. Typing 0.25 also works.

### The sale price shows 0.75 or 36 with no dollar sign

The formula is fine. The cell is formatted as General. Apply the Currency format to columns B, D and E so values show as $36.00.

### The price has more than two decimals

A $19.99 item at 15% off is $16.9915 before display rounding. The cell may show $16.99 while still holding the long number. Wrap the formula to store cents: `=ROUND(B2*(1-C2),2)`.

### The reverse formula returns the wrong original price

Check that C holds the percent you saw on the tag, not the percent of the original you pay. For a sale price after 25% off, C must be 25%, and the formula divides by 0.75.

### The coupon total looks too low

You added the percents. Use one `(1-discount)` factor for each discount, multiplied together, as in the stacked example above.

## Template

[Download the .xlsx](templates/tidy-tabs-sale-price-discount-calculator.xlsx)

The file has three tabs. **Sale prices** holds the eight sample items with the sale price and savings formulas in columns D and E, and the totals in **H1:H3**. **Reverse and stacked** holds three work-back-to-original rows and the 20% then 10% example. **How to use** has short fill-in notes.

The math is arithmetic only. It does not include sales tax or shipping. The items and prices are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
