---
title: How to Set Clearance Prices With a Minimum Margin in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 33 workbooks and 4,686 comparisons with 0 mismatches. Checked: the sample, an empty list, days in stock of 0, 1, 59, 60, 61, 119, 120, 121, 179, 180, 181 and 1000, negative, text and blank days, text and blank cost or price, a floor above the regular price, margin settings of 0, 50, 99, 100, negative, text and blank, half-cent markdowns, odd-cent floors, steps that do not rise, custom steps, and random lists of 3 to 100 items. The model also asserted that every Marked down or At floor price keeps at least the margin setting. Excel and Google Sheets were not used. Not run: TODAY(), more than 100 items, a non-US language setting.
search_description: Set clearance prices that never drop below a minimum margin in Excel and Google Sheets with MIN, MAX and ROUNDUP. Free template with markdown steps.
image_prompts: A shop shelf with a stack of folded linen napkins and two canvas tote bags, blank kraft tags hanging from them, soft shop lighting, realistic photo, no screens, no readable text, no logos
image_alt: Folded linen napkins and canvas totes on a shelf with blank hanging tags
threads: $11.00 would earn only a 7.3% margin on a $10.20 item.\nSo the price stops at $12.75 with =MIN(C6,MAX(F6,G6)).\nThe floor is =ROUNDUP(B6/(1-$B$2),2), and the free template applies it to every item.
---
Take the marked-down price in F6 but never go below the floor with `=MIN(C6,MAX(F6,G6))`, where the floor is `=ROUNDUP(B6/(1-$B$2),2)`. A $10.20 item marked down 50% from $22.00 stays at $12.75 with a 20.0% minimum margin.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. MIN, MAX, ROUNDUP, INDEX and MATCH exist in current Excel and Google Sheets, but neither app was opened for this post.

![Clearance table where one linen item stops at its margin floor while others drop further](images/2026-11-05-team-04-clearance-price-minimum-margin-floor-max-roundup-excel-google-sheets/clearance-price-margin-floor-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*


## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Clearance** has the lowest margin you accept in **B2** and one item per row from row 6: item in **A**, unit cost in **B**, regular price in **C**, days in stock in **D**. **Steps** lists when each markdown starts.

1. On **Steps**, type the first day of each step in **A5:A8**: 0, 60, 120 and 180. Type the markdown for each in **B5:B8**: 0%, 15%, 30% and 50%.
2. On **Clearance**, type the lowest margin you accept in **B2**, such as 20%.
3. In **E6**, find the markdown for the item's age: `=INDEX(Steps!$B$5:$B$8,MATCH(D6,Steps!$A$5:$A$8,1))`. MATCH with 1 finds the last step start that is not larger than the days in stock. Day 60 starts the second step. Day 59 stays in the first.
4. In **F6**, type `=ROUND(C6*(1-E6),2)`. This is the price after markdown.
5. In **G6**, type `=ROUNDUP(B6/(1-$B$2),2)`. This is the lowest price that still earns your margin. ROUNDUP goes up to the next cent, so the margin is never a hair short.
6. In **H6**, type `=MIN(C6,MAX(F6,G6))`. MAX picks the higher of the markdown and the floor. MIN stops the result from going above the regular price.
7. In **I6**, type `=(H6-B6)/H6` and format it as a percent. Fill rows 6 to 105 down.

The margin here is margin on price, not markup on cost. A 20% margin on a $12.75 price means $2.55 is profit, not cost plus 20%.

The template wraps these formulas in checks for blank rows and text, and rounds the floor to six decimals before ROUNDUP so a floating-point crumb cannot push it up a cent. The core of each formula is as above.

Help pages: Microsoft's [MIN function](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152), [MAX function](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098) and [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test, and the pages were not read for this post.

## Example

Sample data is fictional: a gift shop with 12 items and a 20.0% minimum margin. Four rows from the template:

| Item | Cost | Regular | Days | Markdown | Floor | Clearance | Status |
|---|---|---|---|---|---|---|---|
| Cedar soy candle, 8 oz | $4.25 | $18.00 | 74 | 15% | $5.32 | $15.30 | Marked down |
| Canvas tote bag | $4.80 | $14.00 | 232 | 50% | $6.00 | $7.00 | Marked down |
| Linen napkin set | $10.20 | $22.00 | 296 | 50% | $12.75 | $12.75 | At floor |
| Oatmeal soap bar | $1.85 | $8.00 | 8 | 0% | $2.32 | $8.00 | Full price |

The linen napkin set is row 15. Its markdown in **F15** is $11.00, which would earn only 7.3%. The floor in **G15** is $12.75, so **H15** returns $12.75 and the status reads At floor. The canvas tote bag has a floor of $6.00, so its $7.00 markdown stands.

Across the 12 items, 8 are marked down, including 1 held at the floor. Regular prices total $159.50 and clearance prices total $124.88. If one of each sells, the margin is 66.0%. **CHECK** reads OK.

This is arithmetic only. It ignores marketplace fees, shipping and sales tax, so a 20% floor here is not a 20% take-home margin.

## Troubleshooting

### The clearance price is higher than the markdown
The floor is higher than the marked-down price, so MAX chose the floor. That is the point of the formula. Lower the margin in **B2** if you accept less profit on old stock.

### The status says Below min margin
The floor is above the regular price, so even full price does not earn your margin. The price stays at the regular price. Raise the price or accept a lower margin.

### The margin looks too high or too low
The margin is profit divided by price, not by cost. A 50% markup on cost is a 33.3% margin. Type the margin you want in **B2**, not a markup.

### A row shows Check
The days, cost or price is blank or typed as text. In the test, text days and a blank price each showed Check on that row, and the **CHECK** cell said Check.

### Every row is blank
The setting in **B2** must be a percent from 0% up to 99%. In the test, 100%, a negative number, text and a blank cell all stopped the prices, and **B3** said what to fix.

### A step starts at the wrong day
The days in **Steps A5:A8** must start at 0 and rise. In the test, steps that did not rise made **CHECK** say Check.

## Template

[Download the .xlsx](templates/tidy-tabs-clearance-price-floor.xlsx)

The file has three tabs. **Clearance** has the margin setting in **B2** and 100 rows from row 6 (yellow inputs in **A:D**, formulas in **E:J**), with a summary in **L5:M11**. **Steps** has the four yellow day starts and markdown percents. **How to use** has notes on each formula.

**CHECK** in **M11** says OK when the margin setting is valid, no row shows Check or Below min margin, the step days start at 0 and rise, and every markdown is between 0% and 100%.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
