---
title: How to Raise Prices by a Percent and See the New Margin in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and its new prices, margins, price changes, statuses and summary were compared with a Python model. Checked: the Nearest cent and End in .99 rounding choices, a 0% increase, a 10% price cut, a cost above the price, an old price of $0.00, text in a cost cell and a 170% increase. The rounding dropdown menu was not clicked, only its values were set. Excel and Google Sheets were not used.
search_description: Raise prices by a percent and see the new profit margin in Excel and Google Sheets. Free template with ROUND, CEILING, a .99 price option and a below-cost flag.
image_prompts: A craft table with a hand-lettered price tag on a lavender candle, a folded kraft paper bag and a stack of cardboard shipping boxes beside it, soft window light, no screens, no readable text, no logos
image_alt: Lavender candle with a blank price tag beside a kraft bag and shipping boxes
threads: Raised your prices and not sure what the margin did?\nNew price: =ROUND(C5*(1+$B$1),2). New margin: =(D5-B5)/D5.\nA $20.00 cost on an $18.00 item shows -11.1%, and that is correct.
---
Type the increase as a percent in B1, then calculate each new price with `=ROUND(C5*(1+$B$1),2)`, where C is the old price. See the new margin next to it with `=(D5-B5)/D5`, where B is the cost and D is the new price.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. ROUND and CEILING exist in Excel and Google Sheets, but neither app was opened for this post.

![Price list with costs, old and new prices, and margins rising after an 8% increase](images/2026-11-01-team-02-price-increase-planner-new-price-margin-excel-google-sheets/price-increase-new-margin-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The layout matches the template. Row 5 is the first item. **A** is the item name, **B** is the cost, **C** is the old price, **D** is the new price, **E** is the old margin, **F** is the new margin, **G** is the price change and **H** is a status. The increase goes in **B1** and the rounding choice in **B2**.

The steps are the same in Excel and Google Sheets.

1. In **B1**, type the increase as a percent, such as 8%. A negative percent is a price cut.
2. In **A5:C5**, enter an item, its cost and its old price. Enter cost and price as numbers.
3. In **D5**, type the new price formula for rounding to the nearest cent: `=ROUND(C5*(1+$B$1),2)`
4. In **E5**, type the old margin: `=(C5-B5)/C5`
5. In **F5**, type the new margin: `=(D5-B5)/D5`
6. In **G5**, type the price change: `=D5-C5`
7. Format B, C, D and G as currency, and E and F as percent with one decimal.
8. Select **D5:G5** and drag the fill handle down to your last item.

The `$B$1` is locked so every row reads the same increase. Margin here is profit as a share of the price. It is not a markup on cost, so a 50% margin means the cost is half the price.

### Make prices end in .99

To round up to the next whole dollar and take a cent off, replace the formula in **D5** with:

`=ROUND(CEILING(ROUND(C5*(1+$B$1),2),1)-0.01,2)`

The inner ROUND removes tiny fraction errors. CEILING then rounds up to a whole dollar, and the last part subtracts $0.01. In the template, a dropdown in **B2** switches between "Nearest cent" and "End in .99" and the **D** formula reads it.

One quirk was checked: a result that is already a whole dollar still drops a cent. A $20.00 price with a 5% increase is $21.00, and this formula returns $20.99. At 0% an $18.00 price becomes $17.99.

Help pages: Microsoft's [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and [CEILING function](https://support.microsoft.com/en-us/office/ceiling-function-0a5cd7c8-0720-4f0a-bd2c-c943e510899f), and Google's [CEILING page](https://support.google.com/docs/answer/3093471). Those pages describe the functions, not our test.

## Example

Sample data is fictional: 12 made up craft products, an 8% increase and the End in .99 choice. Three of the 12 rows are shown.

| Item | Cost | Old price | New price | Old margin | New margin | Price change |
|---|---|---|---|---|---|---|
| Lavender soy candle, 8 oz | $4.10 | $18.00 | $19.99 | 77.2% | 79.5% | $1.99 |
| Enamel mug | $5.20 | $16.00 | $17.99 | 67.5% | 71.1% | $1.99 |
| Gift set box | $6.75 | $6.50 | $7.99 | -3.8% | 15.5% | $1.49 |

The gift set box cost more than its old price, so its old margin is negative. The increase fixes that, but check whether customers will pay $7.99 before you rely on it.

The other new prices are $19.99 for the cedar candle, $8.99 for the oatmeal soap, $9.99 for the charcoal soap, $25.99 for the hoop earrings, $6.99 for the birthday card, $23.99 for the tote, $16.99 for the art print, $9.99 for the notebook and $4.99 for the sticker pack.

The summary shows 12 items, 0 below cost after the change, an average old margin of 72.6% and an average new margin of 76.2%. Those averages count each item equally, not weighted by sales.

Switch **B2** to Nearest cent and the 8% increase gives prices like $19.44 for an $18.00 item (18.00 times 1.08). Change **B1** to -10% with End in .99 and both candles become $16.99 and the oatmeal soap $7.99. This is arithmetic on typed costs. It does not tell you what price your customers will accept.

## Troubleshooting

### The new price column is blank and B3 shows a message
In the template, **B3** checks the settings. A percent above 100% or below -50% shows "Percent: between -50% and 100%" and every row shows Check settings. In the test, 170% did this. Type a percent in range.

### New price shows $0.00 and margins are blank
The old price is $0.00. The template returns the old price as is, leaves both margins blank and sets the status to No price. Enter the real old price.

### Status says Check row
The cost or the old price is not a number. In the test, text in a cost cell blanked the margins and set Check row. Retype the value as a number without letters.

### Status says Below cost
The new price is under the cost. In the test, a -10% cut on the gift set box gave $5.99 against a $6.75 cost. Raise the price or lower the cut.

### A cost is higher than the price and the margin is negative
That is correct. The formula `=(C5-B5)/C5` returns a negative number when cost is above price. A $20.00 cost on an $18.00 item shows -11.1%.

## Template

[Download the .xlsx](templates/tidy-tabs-price-increase.xlsx)

The file has two tabs. **Prices** has the yellow increase and rounding cells in **B1** and **B2**, a settings check in **B3**, 12 fictional products in **A5:C16** (40 slots, rows 5 to 44), the new price, margin, change and status columns, and a summary in **J** and **K**. **How to use** has short notes. Cost is only what you type, so add shipping, fees and labor yourself if they belong.

Limits: the file was recalculated only in LibreOffice Calc 24.2.7.2. More than 40 items were not tested. The sheet prints landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
