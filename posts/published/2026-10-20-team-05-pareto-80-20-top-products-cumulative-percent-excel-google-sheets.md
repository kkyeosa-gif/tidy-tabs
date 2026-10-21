---
threads_url: https://www.threads.com/@tin_ylab/post/Dewr22Im8AH
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-products-that-make-up-80.html
title: How to Find the Products That Make Up 80 Percent of Your Sales in Excel and Google Sheets
labels: excel-formulas, small-business, sales-reports
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 15 sample products were checked against Python Decimal cumulative sums, plus edited copies (a product landing on exactly 80.0%, a product crossing 80%, ties, a refund row, text revenue, unsorted data, blank rows, one product, a text cut line, an empty table). The Top row color was seen only in the LibreOffice PDF export. Sorting with Data > Sort and the conditional formatting were not exercised in any spreadsheet app. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Find the products that make up 80 percent of sales with a cumulative percent formula in Excel and Google Sheets. Free Pareto 80/20 template, no chart.
image_prompts: Product stock on a storeroom shelf arranged by quantity: a tall stack of boxed candles at the left, a shorter stack of soap bars, and one small box of matches at the right, soft natural window light, no screens, no readable text, no logos
image_alt: Stockroom shelf with tall candle boxes and shorter soap and match stacks
threads: Sort by revenue, then =SUM($B$2:B2)/SUM($B$2:$B$40) gives the running share of sales.\nIn the sample, 6 of 15 products make 85.3% of revenue, so it is closer to 40/85 than 80/20.\nThe cut line sits in its own cell, so you can change 80%.
---
Sort products by revenue, largest first, then use `=SUM($B$2:B2)/SUM($B$2:$B$40)` for the cumulative share of sales. The rows up to the one that first reaches 80% are your top group.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and the results below were recalculated and checked against Python. The formulas use only SUM, IF, COUNTIF and SUMIF, so the Excel and Google Sheets steps use the same text as typed. Neither Excel nor Google Sheets was opened for this post, and sorting was not run in any app.

![Sorted product table with share, cumulative percent and Top or Rest groups, top rows shaded green](images/2026-10-20-team-05-pareto-80-20-top-products-cumulative-percent-excel-google-sheets/pareto-top-products-template.png) *PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

These steps use a **Products** tab with Product in column **A** and Revenue in column **B**.

1. Type Product in **A1**, Revenue in **B1**, Share in **C1**, Cumulative in **D1** and Group in **E1**.
2. Paste your products in **A2** and revenue in **B2**, one row per product. The template holds up to 39 products (rows 2 to 40).
3. Select **A1:B40** and sort Revenue from largest to smallest.
4. In **C2**, type `=B2/SUM($B$2:$B$40)`. This is each product's share of all sales.
5. In **D2**, type `=SUM($B$2:B2)/SUM($B$2:$B$40)`. The first `$B$2` stays fixed while `B2` grows as you fill down, so each row adds up everything from the top.
6. Type your cut line, 80%, in a cell of its own. The template uses **Summary!B1**.
7. In **E2**, type `=IF(SUM($B$1:B1)/SUM($B$2:$B$40)<Summary!$B$1,"Top","Rest")`. The sum stops one row above, so a product is Top while the revenue before it is still under 80%. On row 2, `B1` is the header and counts as 0.
8. Select **C2:E2** and drag the fill handle down to row 40.
9. On **Summary**, count the top group with `=COUNTIF(Products!$E$2:$E$40,"Top")`. Get its revenue share with `=SUMIF(Products!$E$2:$E$40,"Top",Products!$B$2:$B$40)/SUM(Products!$B$2:$B$40)`.

The group test looks one row up because the product that crosses 80% is still needed to reach it, so it stays Top.

The template's formulas wrap these in an `IF` so blank rows stay blank and text revenue shows Check.

### In Excel

Use **Data > Sort** for step 3. Microsoft explains it on the [Sort data in a range or table](https://support.microsoft.com/en-us/office/sort-data-in-a-range-or-table-62d0b95d-2a90-4610-a6ae-2e545c4a4654) page. Format C and D as Percentage with **Home > Number**.

### In Google Sheets

Select the range, then use **Data > Sort range > Advanced range sorting options** and sort column B Z to A. Google describes this on its [Sort & filter your data](https://support.google.com/docs/answer/3540681?hl=en) page. Format C and D with **Format > Number > Percent**. This menu path comes from the help page, not from a run in Google Sheets.

## Example

The sample data is fictional: 15 products from an imaginary candle and soap shop, with last year's revenue.

| Product | Revenue | Share | Cumulative | Group |
|---|---|---|---|---|
| Soy candle, 8 oz | $12,480.00 | 25.6% | 25.6% | Top |
| Lavender soap bar | $9,860.50 | 20.2% | 45.8% | Top |
| Holiday gift set | $7,420.00 | 15.2% | 61.0% | Top |
| Beeswax food wrap | $5,310.25 | 10.9% | 71.9% | Top |
| Linen tote bag | $3,890.00 | 8.0% | 79.8% | Top |
| Ceramic mug | $2,675.50 | 5.5% | 85.3% | Top |
| Wool dryer balls | $1,980.00 | 4.1% | 89.4% | Rest |

The other eight products are Rest, ending at 100.0% with Soap dish.

Total revenue is $48,810.25. Six of 15 products are in Top, which is 40.0% of the products. Their revenue is $41,636.25, or 85.3% of the total.

So this shop is closer to 40/85 than to 80/20. The 80/20 split is a rule of thumb and real data rarely lands on it exactly. The groups show where your revenue sits. They do not say why a product sells or what to restock.

Ceramic mug is Top even though its cumulative share, 85.3%, is past the cut line. The revenue above it, 79.8%, is still under 80%.

## Troubleshooting

### A product past 80% is marked Top
That is by design. The group test looks at the revenue above the row, so the product that crosses 80% stays in Top. In a copy with revenue of 500, 350, 100, 30 and 20, the cumulative shares are 50%, 85% and so on. The second product is the last Top and the third is Rest.

### Groups look wrong and Summary B6 says Sort largest first
The table is not sorted. Groups follow row order, not size, so they are wrong until you sort. In a copy with revenue of 50, 900, 100, 500, 300 and 20, B6 showed Sort largest first and B9 showed CHECK. Sort Revenue largest to smallest and both clear. A blank row above a number also trips B6, so delete blank rows between products.

### Group shows Check
A revenue cell holds text such as "n/a" or "300" typed as text, or a negative number such as a refund. Text is left out of the total. A negative number is in the total, so fix it before you read the groups. In a copy with a -$50.00 refund, that row showed Check and B9 showed CHECK. Also, if the cut line in Summary B1 is text, every group shows Check and B9 shows Check settings.

### The top group has every product, or none
Check the cut line. At 100% every product is Top and at 0 every product is Rest. Type it as a percent such as 80%, not as 80.

### Two products with the same revenue land in different groups
Groups follow row order, so one of two equal products can be Top and the other Rest. Put the one you care about first.

## Template

[Download the .xlsx](templates/tidy-tabs-pareto-top-products.xlsx)

The file has three tabs. **Products** holds the 15 sample rows in yellow input cells **A2:B40**, with Share, Cumulative and Group in **C:E** filled to row 40. Top rows turn green and Check cells turn light red. **Summary** has the cut line in **B1**, the Top counts and shares in **B2:B5**, the Sorted check in **B6** and the overall check in **B9**. **How to use** has short notes. All tabs print portrait on US Letter, and **Products** prints rows 1 to 41.

There is no chart. For in-cell bars, see the REPT post on this blog.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets, and the green Top rows were seen only in the LibreOffice PDF export. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
