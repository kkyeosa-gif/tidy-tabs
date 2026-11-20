---
threads_url: https://www.threads.com/@tin_ylab/post/Df9uym6IKho
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-price-custom-size-between-two.html
title: How to Price a Custom Size Between Two Sizes on Your Price List in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 25 workbooks and 7,000 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Price a custom size between two listed sizes with linear interpolation: MATCH finds the row below, INDEX draws the line. 30 in is $92.50.
image_prompts: Unpainted wooden sign blanks of three different widths leaned in a row against a shop wall from narrowest to widest, a yellow retractable tape measure stretched across the middle one, and a carpenter pencil resting on the floor in front, warm workshop light, realistic photo, no screens, no readable text, no logos
image_alt: Three wooden sign blanks of increasing width with a tape measure across one
threads: Customer wants a 30 in sign and your list jumps from 24 in to 36 in?\nDraw a straight line: $75.00 + 6 x ($110.00 - $75.00) / 12 = $92.50.\nMATCH finds the row below, INDEX does the rest. Wrap MATCH in MIN so the largest size still works.
---
Find the list row just below the custom size with `=MIN(MATCH(x,sizes,1),COUNT(sizes)-1)` as m, then draw a straight line to the next size with `=INDEX(prices,m)+(x-INDEX(sizes,m))*(INDEX(prices,m+1)-INDEX(prices,m))/(INDEX(sizes,m+1)-INDEX(sizes,m))`, with the sizes sorted smallest first. In the made-up sign shop list, 24 in at $75.00 and 36 in at $110.00 give a 30 in sign $92.50, the largest size of 48 in still returns $150.00 because of the MIN cap, and the download shows Outside price list for 10 in or 60 in instead of an error.

This is linear interpolation: the price on the straight line between the two listed sizes around the custom size. It is a starting number, not a rule for what to charge.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (25 workbooks, 7,000 cell comparisons, 0 mismatches), plus the edge cases below. MATCH and INDEX also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six sign quotes with width wanted, list sizes below and above, and interpolated price](images/2026-11-16-team-03-price-custom-size-between-price-list-sizes-linear-interpolation-index-match-excel-google-sheets/custom-size-price-quote-template.png) *Six fictional sign jobs priced from a five-size list: a 30 in sign falls between 24 in and 36 in and comes out at $92.50, while 60 in shows Outside price list (the helper column is hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets.

1. On the **Price list** tab, type your widths in inches in **A5** down, smallest first, one size per row with no blank row between. Type the price for each size in **B**.
2. On the **Quote** tab, type the customer or job in **A5** and the width wanted in **B5**.
3. In **C5**, type `=MIN(MATCH(B5,'Price list'!$A$5:$A$24,1),COUNT('Price list'!$A$5:$A$24)-1)`. MATCH with 1 returns the position of the largest size less than or equal to the width, and Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) page says the list must be in ascending order. For 30 in it returns 3, the 24 in row.
4. Keep the MIN part. For 48 in, the largest size, MATCH returns 5, and the next step would look for a row 6 that is not on the list. MIN caps it at 4, so 48 in uses the 36 in to 48 in line and lands on $150.00.
5. In **F5**, type `=INDEX('Price list'!$B$5:$B$24,C5)+(B5-INDEX('Price list'!$A$5:$A$24,C5))*(INDEX('Price list'!$B$5:$B$24,C5+1)-INDEX('Price list'!$B$5:$B$24,C5))/(INDEX('Price list'!$A$5:$A$24,C5+1)-INDEX('Price list'!$A$5:$A$24,C5))`. [INDEX](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) pulls the size and price in row C5 and the next row. The formula adds the price change per inch, times the inches past the lower size, to the lower price.
6. Fill **C5:F5** down to the last quote.

For 30 in: $75.00 + (30 - 24) x ($110.00 - $75.00) / (36 - 24) = $92.50.

The template adds checks. C5 stays blank unless the width is a number and the Price list CHECK in **F7** shows OK. F5 shows Outside price list for a width below the smallest size or above the largest, and rounds the price to cents. **D5:E5** show the sizes below and above, and **G5** the Status: OK, Outside list or Check.

To land on your own price points, use the [rounding formulas for $0.05 or .99 endings](https://tidytabs.blogspot.com/2026/10/how-to-round-prices-to-nearest-005-or.html). This method uses width only; for width and height together, see [pricing a custom job by square foot](https://tidytabs.blogspot.com/2026/10/how-to-price-custom-job-by-square-foot.html).

## Example

The sign shop, its prices and the jobs are fictional. The price list is 12 in $40.00, 18 in $55.00, 24 in $75.00, 36 in $110.00 and 48 in $150.00.

| Customer or job | Width wanted (in) | Row below | Size below | Size above | Price | Status |
|---|---|---|---|---|---|---|
| Bakery window sign | 30 | 3 | 24 | 36 | $92.50 | OK |
| Cafe menu board | 20 | 2 | 18 | 24 | $61.67 | OK |
| Farm stand banner | 40 | 4 | 36 | 48 | $123.33 | OK |
| Wedding welcome sign | 24 | 3 | 24 | 36 | $75.00 | OK |
| Porch sign | 48 | 4 | 36 | 48 | $150.00 | OK |
| Storefront banner | 60 | 4 | | | Outside price list | Outside list |

A listed size such as 24 in gets its own price. The 20 in menu board works out to $61.666..., shown as $61.67, and a decimal width of 30.5 in comes out $93.96.

The summary in **J4:J9** shows the Price list CHECK OK, 6 quotes listed, 5 priced, 1 outside, 0 rows showing Check and CHECK OK.

With only 12 in at $40.00 and 48 in at $150.00 on the list, 30 in comes out $95.00.

## Troubleshooting

### #REF! for the largest size on the list
You left out the MIN cap. In LibreOffice, plain `MATCH(48,...,1)` returned 5 on the five-size list, and `INDEX(prices,5+1)` on that 5-row range gave #REF!. Add the MIN cap from step 3.

### The price is wrong but there is no error
The sizes are not sorted smallest first. In LibreOffice, `MATCH(30,{12,24,18,36,48},1)` returned 3, the 18 in row, with no error. Sort the Price list by width. The template shows Check: order on the row that breaks the order, and every quote shows Check price list until you fix it.

### Check: order on a size that looks fine
The same width is listed twice, which would make the formula divide by zero, so the template flags the second row. Delete one.

### Outside price list for 10 in or 60 in
The width is below your smallest size or above your largest, so there is no pair of sizes to draw a line between. Price those jobs by hand.

### Check on the Quote or the Price list
On the Quote tab, the width is blank or text, such as "30" or "18 in". Retype it as a number. On the Price list tab, Check: number means a width or price is text (for example "$55.00" typed as text), and Check: row above means a blank row sits inside the list.

## Template

[Download the .xlsx](templates/tidy-tabs-custom-size-price-interpolation.xlsx)

The file has three tabs. **Price list** holds up to 20 sizes in rows 5 to 24: width and price in **A:B** (inputs), a Status in **C**, and **F4:F7** for sizes listed, rows showing Check, anything typed below row 24 (not read) and the CHECK, which needs at least two sizes. **Quote** holds 50 jobs in rows 5 to 54: Customer or job and Width wanted in **A:B** (inputs), then Row below (helper), Size below, Size above, Price and Status in **C:G**, with the summary in **J4:J9**. **How to use** explains each step.

Yellow cells are inputs, gray cells are formulas, and Quote Status turns red for Check and orange for Outside list. Both data tabs print Landscape on US Letter with the header row repeated. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
