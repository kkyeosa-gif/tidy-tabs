---
threads_url: https://www.threads.com/@tin_ylab/post/DeF4IjLGswO
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-markup-and-profit.html
title: How to Calculate Markup and Profit Margin in Excel and Google Sheets
labels: formulas, pricing, etsy-sellers
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps follow Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Price from markup with =B2*(1+C2), find margin with =(D2-B2)/D2, and hit a target margin with =B2/(1-G2). Free .xlsx calculator with sample prices.
image_prompts: A craft fair folding table with three soy candles in glass jars, each with a blank kraft paper price tag tied on with twine, next to a small digital kitchen scale, canvas tent shade in the background, all tag text blurred and unreadable, no logos
image_alt: Soy candles with blank kraft price tags beside a small kitchen scale
threads: A 100% markup is only a 50% margin, and that gap is how prices end up too low.\nFor a target margin use =B2/(1-G2). Using =B2*(1+G2) with a 70% target prices a $3.40 item at $5.78, about a 41% margin.\nThe template has both directions with sample craft prices.
---
Price from markup with `=B2*(1+C2)` and margin with `=(D2-B2)/D2`; to hit a target margin use `=B2/(1-G2)`, not `=B2*(1+G2)`.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template's formulas were recalculated and checked against hand math. The Excel and Google Sheets steps use the same formulas and link to Microsoft Support and Google Docs Editors Help. Neither app was hands-on tested here.

![Pricing sheet listing four handmade products with cost, markup percent, price, and margin columns](images/2026-09-30-team-03-markup-vs-margin-calculator-excel-google-sheets/markup-vs-margin-calculator-template.png) *LibreOffice Calc 24.2 PDF export of the Pricing sheet with sample data (fake). The Profit per unit and Margin check columns are hidden so the numbers stay readable.*

Markup and margin are two ways to describe the same profit. Markup is profit divided by cost. Margin is profit divided by price. A 100% markup only gives a 50% margin, and that gap is where most pricing mistakes start.

## Steps

Set up these columns on one sheet, with your first item in row 2: **A** Item, **B** Unit cost, **C** Markup %, **D** Price from markup, **E** Profit per unit, **F** Actual margin %, **G** Target margin %, **H** Price for target margin, **I** Margin check.

1. Type each item's name in column A and its cost per unit in column B.
2. Type your markup in column C as a percentage. 100% means the price is double the cost.
3. In D2, type `=ROUND(B2*(1+C2),2)` to get the price. ROUND keeps it at whole cents.
4. In E2, type `=D2-B2` to get the profit per unit.
5. In F2, type `=E2/D2`. This is the same as `=(D2-B2)/D2`. It shows the margin your markup really produces.
6. To price from a margin instead, type the margin you want in G2, for example 70%.
7. In H2, type `=ROUND(B2/(1-G2),2)`. Do not use `=B2*(1+G2)`. That treats a margin like a markup and prices you too low.
8. In I2, type `=(H2-B2)/H2` to check the margin you actually get at the rounded price.
9. Select D2:I2 and drag the fill handle down to copy the formulas to the other rows.

### In Excel

1. Format the cells in columns C, F, G, and I as Percentage with one decimal place.
2. Format the cells in columns B, D, E, and H as Currency so prices show as $4.10.

Microsoft documents the rounding function on its [ROUND function page](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c).

### In Google Sheets

1. Format the cells in columns C, F, G, and I as Percent with one decimal place.
2. Format columns B, D, E, and H as Currency.

Google explains the number formats on its [Format numbers in a spreadsheet page](https://support.google.com/docs/answer/56470).

## Example

Sample data below is fictional: made-up products and costs for a small craft business.

| Item | Unit cost | Markup % | Price from markup | Profit per unit | Actual margin % |
|---|---|---|---|---|---|
| Lavender soy candle, 8 oz | $4.10 | 100.0% | $8.20 | $4.10 | 50.0% |
| Oatmeal goat milk soap bar | $1.80 | 150.0% | $4.50 | $2.70 | 60.0% |
| Brass hoop earrings | $3.40 | 200.0% | $10.20 | $6.80 | 66.7% |
| Letterpress birthday card | $0.95 | 300.0% | $3.80 | $2.85 | 75.0% |

Read the candle row: a 100% markup doubles $4.10 to $8.20, but the margin is 50.0%, not 100%.

Now the other direction. Say you want a 70% margin on the earrings.

| Item | Unit cost | Target margin % | Price for target margin | Margin check |
|---|---|---|---|---|
| Lavender soy candle, 8 oz | $4.10 | 50.0% | $8.20 | 50.0% |
| Oatmeal goat milk soap bar | $1.80 | 60.0% | $4.50 | 60.0% |
| Brass hoop earrings | $3.40 | 70.0% | $11.33 | 70.0% |
| Letterpress birthday card | $0.95 | 75.0% | $3.80 | 75.0% |

The earrings price is `=ROUND(3.40/(1-0.70),2)`, which is $11.33. The exact result is $11.3333..., so rounding to cents lands the margin at 69.99%. The sheet shows 70.0% because the cell displays one decimal place. To see the real value, add decimal places to the Margin check cell.

Compare that to `=B2*(1+G2)` with 70%: $3.40 x 1.70 = $5.78. That is a 41.2% margin, far below the 70% you wanted.

## Troubleshooting

### The margin comes out lower than the markup
That is expected. Margin divides profit by the price, and the price is always larger than the cost. A markup of 100% is a margin of 50%. To convert, use margin = markup / (1 + markup).

### The price for a target margin is too low
You probably used `=B2*(1+G2)`. Switch to `=B2/(1-G2)`. Check by typing the price into the margin formula and confirming it returns your target.

### Margin check shows 70.0% but the exact value is 69.99%
The price is rounded to whole cents, so the margin can miss the target by a fraction of a percent. Round up to the next cent if you need to meet the target exactly.

### The formula returns a huge or negative price
A target margin of 100% divides by zero, and above 100% the result is negative. Margin has to be under 100%. Also confirm you typed 70%, not 70, in the cell. A plain 70 means 7,000%.

### Percent cells show 1 instead of 100%
The cell is formatted as a number. Apply a percent format, and type the percent sign when you enter a value, such as 100%.

## Template

[Download the .xlsx](templates/tidy-tabs-markup-vs-margin-calculator.xlsx)

The file has two tabs: **Pricing**, with the four sample items and the formulas above already filled in, and **How to use**, with short fill-in instructions. Prices are arithmetic only. They do not include shipping, fees, or taxes.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
