---
title: How to Calculate Consignment Payouts for Each Shop in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model (cent rounding with Decimal) in 26 workbooks (17,524 cell comparisons, 0 mismatches). Checked: rates of 0% and 100%, a half-cent sale, capital letters and a trailing space in a shop name, a shop with no sales, a shop paid more than it owes, a blank quantity, a price typed as text, an unlisted shop, bad rates, a duplicate shop, 200 filled rows and a row past the table. The PDF export pages are US Letter. Excel and Google Sheets were not used.
search_description: Calculate consignment payouts by shop with ROUND and SUMIFS in Excel and Google Sheets. Free template with a rate per shop and what each owes you.
image_prompts: A small stack of handwritten consignment receipt slips under a ceramic mug and a bundled six-card set of blank greeting cards on a wooden craft-fair table, soft daylight, no screens, no readable text, no logos
image_alt: Receipt slips beside a ceramic mug and blank greeting cards on a craft table
threads: Selling through three shops at three different commission rates?\nRound each sale's cut to cents, then one SUMIFS per shop adds it up.\nThe sheet subtracts what each shop already paid you, so "still owed" is one number per shop.
---
Multiply each sale by the shop's commission rate and round it to cents with `=ROUND(F5*rate,2)`, then add the results per shop with `=SUMIFS(Sales!$H$5:$H$204,Sales!$B$5:$B$204,A5)`. What a shop owes you is its sales total, minus its cut, minus anything it already paid you.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The functions are ROUND, COUNTIFS, INDEX, MATCH and SUMIFS, which exist in Excel and Google Sheets, but neither app was opened for this post.

![Three consignment shops with commission rate, gross sales, shop's cut, your share and still owed](images/2026-10-31-team-02-consignment-payout-by-shop-sumifs-excel-google-sheets/consignment-payouts-by-shop-template.png) *caption: the Shops tab of the template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

You need two lists. **Shops** has one row per shop: name in column A, the commission rate the shop keeps in B, and what it has paid you so far in C, starting in row 5. **Sales** has one row per sale: date in A, shop in B, item in C, price each in D and quantity in E, starting in row 5.

The steps are the same in Excel and Google Sheets.

1. On **Shops**, type each shop name once in **A5:A14**. Type its rate in B as a percent, such as 30%, and what it has paid you in C.
2. On **Sales**, type or paste each sale in columns A to E. Spell the shop exactly as it appears on **Shops**.
3. On **Sales**, in **F5**, type the sale total `=IF(COUNTA(A5:E5)=0,"",IF(AND(ISNUMBER(D5),ISNUMBER(E5)),ROUND(D5*E5,2),"Check amounts"))`.
4. In **G5**, type the shop's cut formula below.
5. In **H5**, type your share `=IF(ISNUMBER(G5),ROUND(F5-G5,2),"")`.
6. Select **F5:H5** and drag the fill handle down to your last sale row.
7. On **Shops**, in **F5**, type `=SUMIFS(Sales!$H$5:$H$204,Sales!$B$5:$B$204,A5)` for your share. In **D5** and **E5**, use the same formula with Sales column F (gross) and column G (cut) as the first range.
8. In **G5** on Shops, type the amount still owed `=ROUND(F5-N(C5),2)`. Drag all the Shops formulas down.

The shop's cut formula in **G5** on Sales:

`=IF(NOT(ISNUMBER(F5)),"",IF(COUNTIFS(Shops!$A$5:$A$14,B5,Shops!$B$5:$B$14,">=0",Shops!$B$5:$B$14,"<=1")=0,"Check shop",ROUND(F5*INDEX(Shops!$B$5:$B$14,MATCH(B5,Shops!$A$5:$A$14,0)),2)))`

MATCH finds the shop's row and INDEX returns its rate. The COUNTIFS part shows "Check shop" when the name is not on the list or has no rate between 0% and 100%, so a typo never gets a made-up rate.

The cut is rounded on each sale on purpose. That way the cut plus your share equals the sale total to the cent, and the shop totals add up to the whole Sales list.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b), [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [function list](https://support.google.com/docs/table/25273), which covers all of them. Those pages describe the functions, not our test. The Google page could not be checked from our build machine.

## Example

Sample data is fictional: 3 made up shops and 30 sales between 09/01/2026 and 10/28/2026.

| Shop | Rate | Gross sales | Shop's cut | Your share | Paid so far | Still owed |
|---|---|---|---|---|---|---|
| Maple Street Gifts | 30% | $380.00 | $114.00 | $266.00 | $150.00 | $116.00 |
| Harbor Art Co-op | 40% | $382.00 | $152.80 | $229.20 | $0.00 | $229.20 |
| Cedar & Stone Boutique | 35% | $332.00 | $116.22 | $215.78 | $85.00 | $130.78 |
| Total | | $1,094.00 | $383.02 | $710.98 | $235.00 | $475.98 |

The shops sold 58 items and the check cell says OK.

One sale shows the rounding: 3 greeting card 6-packs at $12.50 is $37.50. At 35% the cut is $13.125, which the template rounds to $13.13, leaving $24.37 for you. Set a shop's rate to 0% and its cut is $0.00, so your share equals its gross. At 100% your share is $0.00.

This is arithmetic on the numbers you type. It says nothing about your agreement with a shop, unsold stock, or sales tax.

## Troubleshooting

### A sale says Check shop
The shop name in Sales does not match the Shops list. In the tests, a trailing space ("Maple Street Gifts ") found no rate, so the sale was left out of the totals and the check cell showed Check. Capital letters do not matter: "MAPLE STREET GIFTS" matched. Retype the name.

### A sale says Check amounts
The price or quantity is blank or typed as text, such as "12.00" with quotes or a leading apostrophe. The sale is left out of the gross, cut and share, and the check cell shows Check. Retype it as a plain number.

### Still owed is negative
The shop has paid you more than your share so far. The Flag column says "Check: paid more than owed". The check cell stays OK here, because it can be a real advance.

### A shop shows Check rate
The rate is blank, typed as text like "30%" with quotes, or outside 0% to 100%. Type it as a number such as 30% or 0.3. Sales for that shop show Check shop until you do.

### The check cell says Check but every row looks fine
It also shows Check when a shop is listed twice, when a paid amount is text, or when something is typed below row 204 on Sales. A duplicate shop would count its sales twice, so remove one row.

## Template

[Download the .xlsx](templates/tidy-tabs-consignment-payouts.xlsx)

The file has three tabs. **Shops** lists 3 fictional shops in **A5:C14** (10 slots) with gross, cut, your share, still owed, items sold and a flag in **D** to **I**, plus a totals row and a summary with the check cell in **K** and **L**. **Sales** holds 30 fictional sales, with yellow inputs in **A5:E204**. **How to use** has short notes.

Limits: it holds 10 shops and 200 sales. More was not tested, and neither were shop names with * or ?. Shops prints landscape and Sales prints portrait, both on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
