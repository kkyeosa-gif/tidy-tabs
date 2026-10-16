---
threads_url: https://www.threads.com/@tin_ylab/post/DeiM4XRoE-a
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-sell-through-rate-in.html
title: How to Calculate Sell Through Rate in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python, including a received count of 0, more sold than received, a sold count typed as text, a blank sold cell and an empty sheet. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Sell through rate: =C2/B2 formatted as a percent. 41 of 48 lavender candles sold is 85.4%. Free .xlsx inventory template with Strong and Slow flags.
image_prompts: Three small handmade candles in glass jars next to a brown paper shipping box and a bundle of kraft price tags on a wooden shelf, soft daylight, no screens, no readable text, no logos
image_alt: Candles in glass jars beside a shipping box and kraft tags on a wooden shelf
threads: Divide units sold by units received and you have a sell through rate.\n=C2/B2 formatted as a percent turns 41 of 48 lavender candles into 85.4%.\nIn the sample sheet, the two items under 40% still hold 101 units.\nText typed as a number gets skipped, so the post shows how to spot it.
---
Divide units sold by units received with `=C2/B2` and format the cell as a percent. Selling 41 of 48 lavender candles is an 85.4% sell through rate.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Product list with units received, sold and left, plus sell through rates and a summary](images/2026-10-15-team-03-sell-through-rate-inventory-excel-google-sheets/sell-through-rate-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

1. On one tab, type these headers in row 1: Product, Units received, Units sold, Units left, Sell through, Status.
2. Type each product in column **A**, the units received in column **B** and the units sold in column **C**. Use plain numbers, not text.
3. In **D2**, type `=IF(OR(B2="",C2=""),"",B2-C2)`. This is units left.
4. In **E2**, type `=IF(OR(NOT(ISNUMBER(B2)),NOT(ISNUMBER(C2))),"",IF(B2<=0,"",C2/B2))`. This is the rate. The short version is `=C2/B2`; this one stays blank when a count is missing.
5. Select column **E** and format it as a percent with one decimal place.
6. In **I2**, type 70% as the Strong cutoff. In **I3**, type 40% as the Slow cutoff. These are examples, not industry rules. Use your own.
7. In **F2**, type `=IF(E2="","",IF(E2>1,"Check counts",IF(E2>=$I$2,"Strong",IF(E2<$I$3,"Slow","OK"))))`.
8. Fill **D2:F2** down to the last row you need.
9. For the totals, type `=SUM(B2:B30)` in **I5** and `=SUM(C2:C30)` in **I6**. Then type `=IF(I5=0,"",I6/I5)` in **I7** for the overall rate.
10. To count slow items, type `=COUNTIF(F2:F30,"Slow")` in **I8**. To see how many units they still hold, type `=SUMIF(F2:F30,"Slow",D2:D30)` in **I9**.

Do not average the percentages in column E. A rate on 25 units counts the same as a rate on 120 units that way. Divide total sold by total received instead, as in step 9.

### In Excel

The formulas work as typed. Microsoft describes the two summary functions on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) pages. Format column **E** from **Home > Number Format > Percentage**, then use **Home > Increase Decimal** once.

### In Google Sheets

The same formulas should work as typed. Format column **E** with **Format > Number > Percent**, then use **Increase decimal places** in the toolbar once. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small craft seller's stock for one season.

| Product | Received | Sold | Left | Sell through | Status |
|---|---|---|---|---|---|
| Lavender candle | 48 | 41 | 7 | 85.4% | Strong |
| Cedar candle | 48 | 30 | 18 | 62.5% | OK |
| Oatmeal soap | 120 | 96 | 24 | 80.0% | Strong |
| Charcoal soap | 60 | 21 | 39 | 35.0% | Slow |
| Brass hoop earrings | 30 | 12 | 18 | 40.0% | OK |
| Birthday card | 100 | 78 | 22 | 78.0% | Strong |
| Linen tote | 25 | 25 | 0 | 100.0% | Strong |
| Gift tag set | 80 | 18 | 62 | 22.5% | Slow |

| Result | Cell | Value |
|---|---|---|
| Units received | I5 | 511 |
| Units sold | I6 | 321 |
| Overall sell through | I7 | 62.8% |
| Slow items | I8 | 2 |
| Units left on slow items | I9 | 101 |

The brass hoop earrings sit at exactly 40.0%. The test is "below 40%", so they show OK, not Slow. If you want them flagged, raise the Slow cutoff in **I3**.

## Troubleshooting

### A product shows no rate and a negative number in Units left
The received count is 0 or blank-looking, so the rate formula returns nothing. Units left still subtracts, so it goes negative. In the recalculated copy, setting Gift tag set received to 0 left E and F blank and D at -18. There is no warning flag for this, so scan column D for negatives. The overall rate also changed, from 62.8% to 321 of 431, or 74.5%.

### A product is missing from the totals
A number typed as text is skipped by SUM. In the recalculated copy, "96" typed as text in the Oatmeal soap sold cell left E and F blank, and D still showed 24. But I6 dropped to 225 and the overall rate to 44.0%. The row was quietly left out. Text usually lines up on the left of the cell. Retype it as a plain number.

### The rate is above 100%
You sold more than you received, which usually means a typo or a missing restock. In the recalculated copy, 60 sold of 48 received gave 125.0%, the status showed Check counts and Units left showed -12. Fix the count, or add the restock to Units received.

## Limits

The sheet counts units, not dollars, so a $4.00 card and a $38.00 candle weigh the same. Use the same time period for every product, such as one season, or the rates are not comparable. The 70% and 40% cutoffs are your own judgment. A received count of 0 and a text sold count show no warning flag, only blanks.

## Template

[Download the .xlsx](templates/tidy-tabs-sell-through-rate.xlsx)

The file has two tabs. **Products** holds 8 sample products with yellow input cells for rows 2 to 30, the Strong and Slow cutoffs, and the summary cells. **How to use** has short fill-in notes. Formulas read rows 2 to 30.

The product names and counts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
