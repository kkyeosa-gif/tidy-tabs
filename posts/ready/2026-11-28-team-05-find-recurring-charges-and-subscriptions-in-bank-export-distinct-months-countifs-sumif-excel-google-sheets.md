---
title: How to Find Recurring Charges and Subscriptions in a Bank Export in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 224 workbooks and 433,598 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The wildcard miscount was reproduced in LibreOffice Calc only.
search_description: Find subscriptions in a bank or card export: count the distinct months each merchant charged with COUNTIFS and SUMIF, and flag price changes and stops.
image_prompts: A home office table with twelve printed monthly card statements fanned out in a row, several lines on each marked with a yellow highlighter, the uncapped highlighter lying across them and a plain business credit card face down beside the stack, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Twelve fanned paper card statements with highlighted lines, a highlighter and a face-down card
threads: Count months, not charges, to find the subscriptions in a card export.\n=IF(COUNTIFS($B$5:B5,D5,$E$5:E5,E5)=1,1,0) flags each merchant's first charge in a month, and 3 or more months means recurring.\nIn the sample, 4 of 8 merchants were subscriptions, and one went from $24.00 to $27.00.
---
Flag the first charge of each merchant in each month with `=IF(COUNTIFS($B$5:B5,D5,$E$5:E5,E5)=1,1,0)`, then add those flags per merchant with SUMIF: a merchant that charged in 3 or more different months is recurring. In the made-up card export, 4 of 8 merchants come out as subscriptions, including one whose price went from $24.00 to $27.00 and one that stopped after 08/08/2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (224 workbooks, 433,598 cell comparisons, 0 mismatches), plus the edge cases below. The wildcard miscount in Troubleshooting was reproduced in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Eight merchants sorted into subscriptions, repeat vendors and one-time charges, with a price change flagged](images/2026-11-28-team-05-find-recurring-charges-and-subscriptions-in-bank-export-distinct-months-countifs-sumif-excel-google-sheets/recurring-charges-finder-template.png) *Recurring tab of a made-up card export from 12/2025 to 11/2026: eight merchants with charge counts, distinct months, totals and first and last amounts, four marked Subscription in green, PIXELPRO SUITE flagged Changed ($24.00 to $27.00) and MEALKIT CO Not seen in 45+ days in orange, and $629.76 a year at last prices for active subscriptions (settings rows and the Charges column hidden). PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

Count months, not charges. The steps are the same in Excel and Google Sheets.

1. Paste your bank or card export on the **Charges** tab from row 5: Date in **A**, Merchant in **B**, Amount in **C**, with charges as positive amounts. Any order works.
2. Check column **H** (Status). It shows OK, a Check message (date blank or text, no merchant, extra spaces in the merchant, amount blank or text), or "Refund (not counted)" for a negative amount.
3. Column **D** builds a key for each OK row with `=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B5,"~","~~"),"*","~*"),"?","~?")`. COUNTIF, COUNTIFS and SUMIF read `*`, `?` and `~` as wildcards (see Microsoft's [wildcard characters](https://support.microsoft.com/en-us/office/using-wildcard-characters-in-searches-ef94362e-9999-4350-ad74-4d2371110adb) page), so the key escapes them and "SQ *BAKERY" becomes "SQ ~*BAKERY".
4. Column **E** turns each date into its month with `=DATE(YEAR(A5),MONTH(A5),1)`.
5. Column **F** marks the first charge of each merchant in each month with the COUNTIFS above: 1 for the first, 0 for any repeat in the same month. See the COUNTIFS help from [Microsoft](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) or [Google](https://support.google.com/docs/answer/3256550).
6. Column **G** numbers each new merchant, and the **Recurring** tab lists them in rows 7 to 106 with `INDEX` and `MATCH`.
7. On **Recurring**, type the minimum number of months in **C1** (3 in the sample) and the days without a charge that count as stopped in **C3** (45). **C2** reads the latest date in the export.
8. For each merchant, **D** counts the OK charges, **E** adds the month flags with `=SUMIF(Charges!$B$5:$B$2004,C7,Charges!$F$5:$F$2004)`, and **F** totals the amounts with SUMIFS. **G** and **H** find the first and last charge dates, and **I** and **J** the amounts on those dates.
9. Read the **Type** in column **K**. With at least C1 months, a merchant is a Subscription when its last amount appears in 2 or more charges, otherwise a Repeat vendor. Below that it is Occasional, or One-time for one charge.
10. Column **L** says Changed when a subscription's last amount differs from its first, **M** says "Not seen in 45+ days" when its last charge is at least C3 days before C2, and **N** shows the last price x 12 for subscriptions still running.

The summary in **P6:Q18** counts each type and totals the yearly cost and subscription spending. **CHECK** in **Q18** is OK only when the settings are valid, every merchant fits in the list, and the Charges CHECK is OK.

"Recurring" here means charged in at least C1 different months. A plan billed once a year shows as One-time, so look over those lines too. The x 12 figure assumes a monthly plan.

Capital letters do not matter to COUNTIF, so "Cloudbox Storage" and "CLOUDBOX STORAGE" count as one merchant. Spaces and store numbers do matter. If your bank adds a city or store number to each name, clean the names first with [How to Auto-Categorize Expenses From Bank Descriptions](https://tidytabs.blogspot.com/2026/10/how-to-auto-categorize-expenses-from.html). PayPal and app-store charges may hide the real merchant.

The sheet lists what repeats; what to keep, cancel or deduct is your call. To track the ones you keep, see [How to Track Subscription Renewals](https://tidytabs.blogspot.com/2026/10/how-to-track-subscription-renewals-in.html).

## Example

The card export below is made up: 58 charges from 12/03/2025 to 11/21/2026 across 8 merchants, with Cloudbox Storage typed in two cases. The **Recurring** tab after recalculation:

| Merchant | Charges | Months | Total | First amt | Last amt | Type | Flag | Yearly |
|---|---|---|---|---|---|---|---|---|
| Cloudbox Storage | 12 | 12 | $119.88 | $9.99 | $9.99 | Subscription | | $119.88 |
| MEALKIT CO | 9 | 9 | $711.00 | $79.00 | $79.00 | Subscription | Not seen in 45+ days | |
| PIXELPRO SUITE | 12 | 12 | $306.00 | $24.00 | $27.00 | Subscription | Changed | $324.00 |
| STREAMLY | 12 | 12 | $185.88 | $15.49 | $15.49 | Subscription | | $185.88 |
| OFFICE MART | 5 | 4 | $229.30 | $63.18 | $23.90 | Repeat vendor | | |
| SQ *BAKERY | 3 | 2 | $57.25 | $18.50 | $16.00 | Occasional | | |
| SITEHOST PRO | 1 | 1 | $120.00 | $120.00 | $120.00 | One-time | | |
| SQ *MAIN ST BAKERY | 4 | 4 | $123.55 | $31.20 | $29.85 | Repeat vendor | | |

MEALKIT CO last charged on 08/08/2026, more than 45 days before the export's end on 11/21/2026, so it is left out of the yearly total. The summary shows 4 subscriptions, 2 repeat vendors, 1 occasional and 1 one-time merchant, $629.76 a year at last prices, and $1,322.76 spent on subscriptions over the export. The Charges tab shows 58 OK rows totaling $1,852.86.

## Troubleshooting

### A merchant with an asterisk shows too many charges
The COUNTIF criterion was not escaped, so `*` matched any text. In a LibreOffice copy of the sample, `=COUNTIF(Charges!$B$5:$B$2004,"SQ *BAKERY")` returned 7 and the matching SUMIF on the month flags returned 6, because it also counted SQ *MAIN ST BAKERY. With the escaped key "SQ ~*BAKERY" the results were 3 charges in 2 months.

### A merchant appears twice in the list or shows Check: extra spaces
The name has a leading, trailing or double space, such as "STREAMLY " or "SQ  *BAKERY". Those rows are not counted and CHECK says Check. Retype the name or clean the column with TRIM before pasting.

### A monthly subscription shows Repeat vendor
Two charges on the last date were added together, or the last amount differs from every earlier one. In the tested case, two STREAMLY charges on 11/21/2026 gave a last amount of $30.98, so the type turned Repeat vendor.

### Merchants not listed is above 0 and CHECK is red
The export has more than 100 merchants, and the Recurring tab has room for 100. With 130 merchants, **Q16** showed 30. Split the export by date range.

## Template

[Download the .xlsx](templates/tidy-tabs-recurring-charges-finder.xlsx)

**Charges** takes 2,000 rows (rows 5 to 2004) with a summary in **K4:L10** that also catches rows typed below row 2004. **Recurring** lists up to 100 merchants plus the summary in **P6:Q18**. **How to use** explains each formula. The data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Subscriptions and the yearly total are light green, Changed and Not seen are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
