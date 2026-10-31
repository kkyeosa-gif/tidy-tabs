---
title: How to See Which Promo Codes Made Money in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model that uses Decimal. Checked: a code typed in lowercase, a trailing space in a code, an unknown code, a code listed twice, a subtotal typed as text, no blank codes, and a listed code with no orders. Excel and Google Sheets were not used.
search_description: See which promo codes made money with COUNTIF and SUMIF in Excel and Google Sheets. Free template with orders, discounts given and net per order.
image_prompts: A stack of small paper discount cards with a hole punched in one corner next to a kraft shipping box and a roll of twine on a wooden packing table, soft daylight, no screens, no readable text, no logos
image_alt: Paper discount cards, a kraft shipping box and twine on a packing table
threads: Which promo code actually made money?\nCOUNTIF gives orders per code, SUMIF totals the sales, and a second SUMIF on the discount column comes off.\nWhat is left is the net per code, not just the sales.
---
Count the orders for each code with `=COUNTIF(Orders!$C$2:$C$301,A5)`, total its sales with `=SUMIF(Orders!$C$2:$C$301,A5,Orders!$D$2:$D$301)`, and subtract the discounts given with a second SUMIF on column E. What is left is what each code actually brought in.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIF, SUMIF and SUMPRODUCT also exist in Excel and Google Sheets, but neither app was opened for this post.

![Promo code table with orders, discounts given and net sales for each code](images/2026-11-01-team-03-promo-code-results-sumifs-excel-google-sheets/promo-code-results-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

You need one list of orders and one short list of codes. **Orders** has one row per order, starting in row 2: order number in A, date in B, code in C, subtotal in D, discount in E and shipping in F. Leave C blank when the customer used no code. The steps are the same in Excel and Google Sheets.

1. On a new sheet named **Summary**, type your codes once each in **A5:A9**. Type them exactly as customers enter them.
2. In **B5**, count the orders: `=IF(A5="","",COUNTIF(Orders!$C$2:$C$301,A5))`.
3. In **C5**, total the subtotals: `=IF(A5="","",SUMIF(Orders!$C$2:$C$301,A5,Orders!$D$2:$D$301))`.
4. In **D5**, total the discounts given: `=IF(A5="","",SUMIF(Orders!$C$2:$C$301,A5,Orders!$E$2:$E$301))`.
5. In **E5**, subtract: `=IF(A5="","",C5-D5)`. This is the net after discount.
6. In **F5**, find the average subtotal: `=IF(A5="","",IF(B5=0,0,ROUND(C5/B5,2)))`.
7. In **G5**, find the discount rate: `=IF(A5="","",IF(C5=0,0,D5/C5))`. Format it as a percent.
8. In **H5**, find the net per order: `=IF(A5="","",IF(B5=0,0,ROUND(E5/B5,2)))`.
9. Select **B5:H5** and drag the fill handle down to row 9.

The IF wrappers keep an empty code row blank. The `$` signs keep the ranges fixed when you drag.

### Orders with no code

COUNTIF cannot count blanks the way you want here, so row 10 uses SUMPRODUCT. Type **(no code)** in A10, then:

- **B10**: `=SUMPRODUCT((Orders!$A$2:$A$301<>"")*(Orders!$C$2:$C$301=""))`
- **C10**: `=SUMPRODUCT((Orders!$A$2:$A$301<>"")*(Orders!$C$2:$C$301=""),Orders!$D$2:$D$301)`
- **D10**: the same as C10, with column E as the last range

The first part counts only rows that have an order number, so the empty rows below your data are ignored. In row 11, total the whole Orders sheet directly: `=COUNTA(Orders!$A$2:$A$301)` in B11, `=SUM(Orders!$D$2:$D$301)` in C11 and `=SUM(Orders!$E$2:$E$301)` in D11. Copy E to H down to rows 10 and 11.

### Catch typos with a check cell

In **B13**, count orders whose code is not on your list: `=B11-SUM(B5:B9)-B10`. It should be 0. In the template, **B14** turns this into a CHECK cell that shows OK or CHECK. It also compares the dollar totals and flags a duplicated code or a subtotal typed as text.

Help pages: Microsoft's [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34), and Google's [SUMIF page](https://support.google.com/docs/answer/3093583). Those pages describe the functions, not our test.

## Example

Sample data is fictional: 60 made up orders, one per day starting 09/01/2026, and five made up codes. WELCOME15, SPRING10 and VIP20 take 15%, 10% and 20% off the subtotal. SAVE5 takes a flat $5.00 off. FREESHIP takes $0.00 off the subtotal and waives shipping.

| Code | Orders | Subtotal | Discount given | Net after discount | Discount % | Net per order |
|---|---|---|---|---|---|---|
| WELCOME15 | 8 | $575.68 | $86.36 | $489.32 | 15.0% | $61.17 |
| SPRING10 | 17 | $778.65 | $77.86 | $700.79 | 10.0% | $41.22 |
| VIP20 | 7 | $440.25 | $88.05 | $352.20 | 20.0% | $50.31 |
| SAVE5 | 10 | $521.24 | $50.00 | $471.24 | 9.6% | $47.12 |
| FREESHIP | 5 | $256.98 | $0.00 | $256.98 | 0.0% | $51.40 |
| (no code) | 13 | $737.77 | $0.00 | $737.77 | 0.0% | $56.75 |
| All orders | 60 | $3,310.57 | $302.27 | $3,008.30 | 9.1% | $50.14 |

The check cell says OK. The sample reads like this: only WELCOME15 ($61.17) beats the no-code orders ($56.75) on net per order. SPRING10 ($41.22), VIP20 ($50.31), SAVE5 ($47.12) and FREESHIP ($51.40) are all below it.

SPRING10 brought in the most net dollars ($700.79) because it had the most orders, yet each order is worth the least. Both facts are true, so look at both columns.

This is revenue only. Net per order does not subtract what the items cost you, shipping or fees, so it is not profit. Shipping sits on the Orders sheet but the summary does not use it.

## Troubleshooting

### A code's count is lower than it should be
Check for a trailing space. In the tests, one order typed as "SPRING10 " (with a space) was not counted under SPRING10. That code dropped from 17 orders to 16 and from $778.65 to $763.76. B13 showed 1 and the check cell showed CHECK. Retype the code without the space.

### Lowercase codes
Capitals do not matter. In LibreOffice, "spring10" was counted with SPRING10, and the count stayed at 17 with the check cell on OK. We did not test this in Excel or Google Sheets.

### B13 is more than 0
Some orders use a code that is not in A5:A9. In the tests, one order with an unlisted code, NEWCODE, took WELCOME15 from 8 orders to 7 and showed 1 in B13. Add the code to your list or fix the typo on the Orders sheet.

### B13 is negative
You listed the same code twice. When VIP20 was listed twice, B13 showed -2 and the check cell showed CHECK, because the orders were counted in two rows. Remove one of them.

### Totals are lower than the order list
A subtotal was typed as text. SUM skips text, so in the tests one text subtotal dropped the grand total from $3,310.57 to $3,265.97 and the check cell showed CHECK. Retype the value as a plain number.

## Template

[Download the .xlsx](templates/tidy-tabs-promo-code-results.xlsx)

The file has three tabs. **Summary** has the five codes in **A5:A9** (yellow), the (no code) and All orders rows, the unlisted-code count in **B13** and the check cell in **B14**. **Orders** holds 60 fictional orders in **A2:F61**, with yellow input cells down to row 301. **How to use** has short notes.

Limits: the formulas read Orders rows 2 to 301, so an order below row 301 is not counted. Summary prints landscape and everything prints on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
