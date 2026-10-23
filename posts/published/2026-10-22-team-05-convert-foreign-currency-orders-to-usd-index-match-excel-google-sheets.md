---
threads_url: https://www.threads.com/@tin_ylab/post/De11cssEZG3
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-convert-orders-in-canadian.html
title: How to Convert Orders in Canadian Dollars, Pounds or Euros to US Dollars in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 19 edited variants (a code not in the table, lower case codes, a trailing space, blank, zero, text and negative rates, text amounts, refunds, half cent amounts, a duplicate code in the rate table, 100 random orders). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Convert Canadian dollar, pound and euro orders to US dollars with a rate table and INDEX and MATCH in Excel and Google Sheets. Free template.
image_prompts: A glass jar of mixed foreign coins, a small brown paper parcel tied with twine and a pair of scissors on a craft packing table, soft window light, no screens, no readable text, no logos
image_alt: Jar of mixed coins, twine-tied parcel and scissors on a craft packing table
threads: Orders in CAD, GBP and EUR, and you need dollars?\nKeep a small rate table, then INDEX and MATCH the code: 48.50 CAD at 0.73 is $35.41.\nType your own rates. The sample ones are made up, and the file does not fetch any.
---
Convert an order to US dollars with `=IFERROR(ROUND(D2*INDEX(Rates!$B$2:$B$10,MATCH(C2,Rates!$A$2:$A$10,0)),2),"Check currency")`, where D2 is the amount and C2 is the currency code such as CAD or GBP. The rates are ones you type in from your own payment statement, and the file does not fetch them.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Orders sheet with Canadian dollar, pound and euro amounts converted to US dollars](images/2026-10-22-team-05-convert-foreign-currency-orders-to-usd-index-match-excel-google-sheets/currency-to-usd-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: keep one small table of currency codes and their rates, then let each order look up its own rate.

1. Open a tab named **Rates**. In **A2:A10**, type the currency codes: USD, CAD, GBP, EUR, AUD. Use the three letter code exactly as it appears on your orders.
2. In **B2:B10**, type how many US dollars one unit of that currency is worth. USD is 1. For example, 1 CAD = 0.73 US dollars in the sample.
3. On the **Orders** tab, put the currency code in column **C** and the amount in the buyer's currency in column **D**.
4. In **E2**, type the formula from the top of this post and fill it down.
5. Format column E as currency with two decimals.

Here is how the formula reads from the inside out:

- `MATCH(C2,Rates!$A$2:$A$10,0)` finds the row of the code in the rate table. The 0 means an exact match.
- `INDEX(Rates!$B$2:$B$10, ...)` returns the rate from that same row.
- `D2*` multiplies the amount by the rate.
- `ROUND(...,2)` rounds each order to cents.
- `IFERROR(...,"Check currency")` shows words instead of #N/A when the code is not in the table.

The dollar signs lock the table, so it does not slide down as you fill the formula.

### What the template adds

The template's E2 is longer than the bare formula. It first looks for the code with MATCH, so an unknown code shows "Check currency". Then it shows "Check amount" when the amount is blank or text, and "Check rate" when the rate cell is blank, text, 0 or negative. The rounded result is the same ROUND, INDEX and MATCH core.

Use the most recent rate from your own payment or bank statement, and type the date in the "Rate as of" column. The template's sample rates are examples, not real quotes.

### In Excel

The formulas work as typed. INDEX and MATCH have been in Excel for many years, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed. Google Sheets also has a GOOGLEFINANCE function that can pull rates, described on Google's [GOOGLEFINANCE help page](https://support.google.com/docs/answer/3093281). The template does not use it, and it was not tested for this post.

## Example

Sample data is fictional: Etsy style orders in other currencies. The rates are made up for the example (1 CAD = 0.73, 1 GBP = 1.27, 1 EUR = 1.08, 1 AUD = 0.65 US dollars), so type your own.

| Row | Currency | Amount in buyer's currency | Rate | USD amount |
|---|---|---|---|---|
| 2 | CAD | 48.50 | 0.7300 | $35.41 |
| 3 | GBP | 36.00 | 1.2700 | $45.72 |
| 4 | USD | 24.00 | 1.0000 | $24.00 |
| 5 | EUR | 52.00 | 1.0800 | $56.16 |
| 6 | CAD | 64.00 | 0.7300 | $46.72 |
| 7 | AUD | 40.00 | 0.6500 | $26.00 |

Row 2 is 48.50 x 0.73 = 35.405, which rounds up to $35.41.

The template converts 30 sample orders. The total in US dollars is $1,046.40 (**H1** on Orders), with 30 orders converted and none flagged. **H5** is a check cell that shows OK when no row is flagged and the totals agree.

The Rates tab also shows each currency's typed amounts, the same amounts converted at once, and the sum of the rounded rows. Rounding each GBP order gives $306.72, while converting the $241.50 GBP total once gives $306.71. That one cent is rounding, not a mistake. The other currencies matched to the cent in the sample.

## Troubleshooting

### The cell says Check currency
The code in column C is not in the rate table. Common causes are a code that was never added, a blank code on a row that has an amount, and a trailing space such as "GBP ". The match is exact apart from capital letters, so "gbp" and "Cad" still work but "GBP " does not. Retype the code or add it to the table.

### The cell says Check rate
The code was found, but its rate is blank, text, 0 or negative. A rate typed as text such as "1.27" in a text formatted cell is the usual cause. Retype it as a number. Every order in that currency is flagged until you do.

### The cell says Check amount
The amount is blank or text, for example "$48.50" typed as text or "n/a". Type just the number.

### My total is a cent off from my payment statement
Each order is rounded to cents before it is added. Your payment service may round the total once, or use its own rate. Compare the two columns on the Rates tab to see how large the rounding gap is. Rate differences are a separate issue and will not show up in the template.

### The check cell says CHECK
At least one row is flagged, or a currency code is listed twice in the rate table. With a duplicate code, the first row is used for the orders, but the by currency totals count the dollars twice. Remove the duplicate row.

## Template

[Download the .xlsx](templates/tidy-tabs-currency-to-usd.xlsx)

The file has three tabs. **Orders** holds the yellow inputs in **A2:D101** (date, order, currency, amount, up to 100 orders) and the converted amounts in **E2:E101**, with the total and check in **G1:H5**. **Rates** holds the yellow rate table in **A2:D10** (code, US dollars per 1 unit, name, rate as of date) and a one cent rounding comparison in **E:H**. **How to use** has short notes. The Orders and How to use tabs print portrait on US Letter, and Rates prints landscape on US Letter.

The 30 orders and the rates are fictional and the rates are not real quotes. The file was recalculated only in LibreOffice Calc 24.2.7.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
