---
title: How to Make a Price Quote With a Valid-Until Date in Excel and Google Sheets
labels: excel-formulas, small-business, quotes
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (Decimal, round half up) for the line amounts, subtotal, discount, tax, total, valid-until date, status and check cell. Edge cases run on modified copies: quantity 0, a cleared line, a 100% discount, an as-of date on the valid-until day and one day later, cent rounding, a described line with no quantity, a negative quantity and a 150% discount. TODAY() in the as-of cell was not run. Excel and Google Sheets were not used.
search_description: Make a price quote with a valid-until date in Excel and Google Sheets. Free template with days valid, an Open or Expired status, discount, tax and a check cell.
image_prompts: A photographer's desk with a stack of unmarked 8 x 10 prints, a closed canvas-wrapped frame leaning on the wall and a clipboard holding a blank sheet of paper, soft window light, no screens, no readable text, no logos
image_alt: Photographer's desk with a stack of prints, a canvas frame and a clipboard
threads: Is that quote still open, or did it quietly expire?\nOne cell for the quote date, one for days valid, then =B2+B3 gives the valid-until date.\n=IF(B6>E2,"Expired","Open") beside it flags a lapsed quote on its own.
---
Enter the quote date in one cell and the number of days it stays good in the next, then add them for the valid-until date. In the template that is `=B2+B3` in **E2**. Put `=IF(B6>E2,"Expired","Open")` beside it, where B6 is the as-of date, so a lapsed quote flags itself.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. IF, ROUND, SUM and TODAY exist in Excel and Google Sheets, but neither app was opened for this post.

![Photographer price quote listing sessions, prints and canvas, with a valid-until date and Open status](images/2026-11-01-team-01-quote-template-valid-until-date-excel-google-sheets/price-quote-valid-until-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template: settings in **B2:B6**, line items in **A8:D17**, results in column E.

1. In **B2**, type the quote date, such as 10/31/2026.
2. In **B3**, type the number of days the quote is good for, such as 30.
3. In **E2**, type `=B2+B3`. Format the cell as a date (MM/DD/YYYY). If you see a number like 46356 instead of a date, the cell has a number format, not a date format.
4. In **B6**, type the as-of date. This is the day you compare against.
5. In **E3**, type `=IF(B6>E2,"Expired","Open")`.
6. In **E8**, type `=ROUND(C8*D8,2)` and fill it down to **E17**. Quantity goes in column C and unit price in D.
7. In **E19**, type `=SUM(E8:E17)` for the subtotal.
8. In **E20**, type `=ROUND(E19*B4,2)` for the discount, with the discount percent in **B4**.
9. In **E21**, type `=ROUND((E19-E20)*B5,2)` for tax, with the tax rate in **B5**.
10. In **E22**, type `=E19-E20+E21` for the total.

The status uses `>`, not `>=`. A quote is still Open on its valid-until date and turns Expired the next day. In the tests, an as-of date of 11/30/2026 gave Open and 12/01/2026 gave Expired.

Type the as-of date on purpose, or type `=TODAY()` in **B6** to make the status live. The `=TODAY()` version was not run in the tests, so check it in your own app.

The template's formulas add guards for blank or text cells, so they are longer than the ones above. For example, its **E2** is `=IF(AND(ISNUMBER(B2),ISNUMBER(B3)),B2+B3,"")`.

Help pages: Microsoft's [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) and [TODAY function](https://support.microsoft.com/en-us/office/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9), and Google's [TODAY page](https://support.google.com/docs/answer/3092984). Those pages describe the functions, not our test.

## Example

Sample data is fictional: a made-up photographer quoting six lines. Quote date 10/31/2026, 30 days valid, 10% discount, 6.25% tax, as-of date 11/15/2026.

| Item | Qty | Unit price | Amount |
|---|---|---|---|
| Headshot session, 1 hour | 1 | $150.00 | $150.00 |
| Edited digital photos | 12 | $18.00 | $216.00 |
| Prints, 8 x 10 in | 6 | $12.50 | $75.00 |
| Canvas wrap, 16 x 20 in | 1 | $89.00 | $89.00 |
| Travel to Boston, MA 02108 | 1 | $35.00 | $35.00 |
| Rush delivery, 48 hours | 1 | $25.00 | $25.00 |

The subtotal is $590.00. The 10% discount is $59.00, which leaves $531.00 to tax. Tax at 6.25% is $33.19 (33.1875 rounded), so the total is $564.19.

Valid until shows 11/30/2026 and the status shows Open, because the as-of date 11/15/2026 is before it. The check cell says OK.

Edge cases from the tests: setting the photos to quantity 0 gives a $374.00 subtotal and a $357.64 total. Clearing the rush line gives $565.00 and $540.28. A 100% discount gives a $0.00 total with $0.00 tax. This sheet does arithmetic only. It does not say what tax you owe or whether a quote is binding.

## Troubleshooting

### Valid until shows a number like 46356
The cell is formatted as a number. Select **E2** and apply a date format, MM/DD/YYYY. The formula is fine.

### Status stays blank
The status is blank when the valid-until date is blank. That happens when the quote date or days valid is text or empty. Retype the date as a real date and the days as a number.

### The status says Open on the last day
That is intended. The quote is good through the valid-until date, so it flips to Expired the next day. To expire on the date itself, change `>` to `>=`.

### The total is a cent off from a hand calculation
Each line, the discount and the tax are rounded to cents with ROUND. In one test, three units at $33.33 with a 7.5% discount gave a $539.99 subtotal, a $40.50 discount (40.49925 rounded) and a $530.71 total. Rounding each step is why totals can differ by a cent from one multiplication.

### The check cell says CHECK
It shows CHECK when a line has a description but no quantity or price, a quantity is negative, the discount is outside 0% to 100%, or a setting is not a number. In the tests, a described line with no quantity, a quantity of -1 and a 150% discount each gave CHECK.

## Template

[Download the .xlsx](templates/tidy-tabs-price-quote.xlsx)

The file has two tabs. **Quote** has yellow input cells for the quote date, days valid, discount, tax rate and as-of date, six fictional lines in **A8:D17** (10 slots), the valid-until and status cells in **E2:E3**, the totals and a check cell in **E23**. **How to use** has short notes.

Limits: the file was recalculated only in LibreOffice Calc 24.2.7.2, and the live `=TODAY()` version was not run. It prints portrait on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
