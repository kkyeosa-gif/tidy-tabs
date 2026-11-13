---
title: How to Calculate the Interest Rate From a Monthly Payment in Excel and Google Sheets
labels: excel-formulas, small-business, loans
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 10 workbooks and 3,050 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Find the monthly rate a payment plan implies with =RATE(F5,-E5,B5-C5), then the yearly figure, for several offers side by side. Free template.
image_prompts: A new glass display case still on its shipping pallet in a small gift shop, bubble wrap pulled back from one corner, a clipboard with a blank payment plan form and a pen resting on top, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: New glass display case on a shipping pallet with a clipboard and pen on top
threads: Find the rate a payment plan implies with =RATE(F5,-E5,B5-C5): 24 payments of $109.00 on a $2,400.00 kiln is 0.7012% a month, 8.41% as rate x 12.\nPay the first one at signing and it goes to 9.17%.\nLeave the minus off the payment and RATE gives #N/A.
---
Find the monthly rate a payment plan implies with `=RATE(F5,-E5,B5-C5)`, where F5 is the number of payments, E5 the monthly payment, B5 the cash price and C5 the down payment, then multiply by 12 for a yearly figure. In the made-up sample, a $2,400.00 kiln paid as 24 payments of $109.00 works out to 0.7012% a month, 8.41% as rate x 12 and 8.75% as an effective yearly rate with `=(1+J5)^12-1`.

That number is only what the payments imply. It is not the APR the lender has to disclose, and this post is not advice on whether to finance anything. It makes no tax or legal call either.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (10 workbooks, 3,050 cell comparisons, 0 mismatches), plus the edge cases below. RATE also exists in Excel and Google Sheets, but neither app was opened for this post; how RATE converges there and how it handles the type argument are the open questions.

![Eight financing offers with cash price, monthly payment, rate times 12 and effective yearly rate](images/2026-11-14-team-05-find-interest-rate-of-payment-plan-offer-rate-function-excel-google-sheets/payment-plan-interest-rate-template.png) *Eight payment plan offers side by side (down payment and fee columns are hidden here); two 0% plans show No interest and the vinyl cutter costs less than cash. PDF export from LibreOffice Calc 24.2.*

## Steps

The formulas are the same in Excel and Google Sheets. Use one row per offer, starting in row 5.

1. Type the offer name in **A5**, the cash price in **B5**, the down payment in **C5** and any upfront fee in **D5**. Leave C5 and D5 blank when there are none.
2. Type the monthly payment in **E5** and the number of payments in **F5**.
3. In **G5**, type Yes if the first payment is due at signing, or No if it is due a month later.
4. In **H5**, add up everything you pay: `=C5+D5+E5*F5`. **I5** is `=H5-B5`, the extra over the cash price.
5. In **J5**, find the monthly rate. The short version is `=RATE(F5,-E5,B5-C5)`. The payment is negative because it is money going out.
6. To include the fee and the payment timing, use `=RATE(F5,-E5,B5-C5-D5,0,IF(G5="Yes",1,0))`. The fee comes off the amount financed, and the last argument (type) is 1 when the first payment is due at signing.
7. In **K5**, type `=J5*12` for the simple yearly figure. In **L5**, type `=(1+J5)^12-1` for the effective yearly rate, which counts interest on interest.
8. Format J5 as a percentage with 4 decimals (0.0000%) and K5:L5 with 2. Fill **H5:L5** down to the last offer.

The template's J5 wraps step 6 in checks. It returns 0 when the payments add up to the cash price, and shows Check for a text, blank, negative or fractional entry, a down payment at or above the price, or a rate RATE cannot find. **M5** then says OK, No interest, Costs less than cash or Check.

Help pages: Microsoft's [RATE function](https://support.microsoft.com/en-us/office/rate-function-9f665657-4a7e-4bb7-a030-83fc59e748ce) and Google's [RATE](https://support.google.com/docs/answer/3093257). To go the other way, from a rate to a payment, see [How to Calculate a Loan Payment in Excel and Google Sheets](https://tidytabs.blogspot.com/2026/10/how-to-calculate-loan-payment-in-excel.html).

## Example

The sample offers are fictional.

| Offer | Cash price | Down | Fee | Payment | Payments | At signing | Extra vs cash | Monthly rate | Rate x 12 | Effective yearly | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Kiln, 24 months | $2,400.00 | | | $109.00 | 24 | No | $216.00 | 0.7012% | 8.41% | 8.75% | OK |
| Kiln, 24 months, 1st at signing | $2,400.00 | | | $109.00 | 24 | Yes | $216.00 | 0.7644% | 9.17% | 9.57% | OK |
| Booth tent | $899.00 | $99.00 | | $72.00 | 12 | No | $64.00 | 1.2043% | 14.45% | 15.45% | OK |
| Laptop | $1,500.00 | $300.00 | | $75.00 | 18 | No | $150.00 | 1.2704% | 15.24% | 16.36% | OK |
| Display case | $1,800.00 | | $49.00 | $158.00 | 12 | No | $145.00 | 1.2457% | 14.95% | 16.02% | OK |
| Sewing machine, 0% | $1,200.00 | | | $100.00 | 12 | No | $0.00 | 0.0000% | 0.00% | 0.00% | No interest |
| Card reader, pay in 4 | $299.00 | | | $74.75 | 4 | Yes | $0.00 | 0.0000% | 0.00% | 0.00% | No interest |
| Vinyl cutter | $600.00 | | | $95.00 | 6 | No | -$30.00 | -1.4461% | -17.35% | -16.04% | Costs less than cash |

The two kiln rows cost the same $216.00 extra, but paying the first $109.00 at signing raises the rate from 8.41% to 9.17% x 12.

The Display case's $49.00 fee matters too. Leave it out and RATE gives 9.70% x 12; with it, 14.95%.

The summary in **O4:P8** shows 8 offers listed, Kiln, 24 months as the lowest yearly rate with interest at 8.75%, 2 no-interest offers and 0 rows showing Check.

## Troubleshooting

### RATE returns #N/A
The payment was typed as a positive number: `=RATE(24,109,2400)` gave #N/A in LibreOffice. Put a minus sign before the payment. An extreme rate also fails: `=RATE(36,-1000,100)` gave #N/A even with a guess of 5, and the template shows Check.

### The rate looks too low
The upfront fee was left out (Display case 0.8086% instead of 1.2457% a month), or the first payment is due at signing but type was left at 0 (kiln 0.7012% instead of 0.7644%). Subtract the fee from the price and set type to 1.

### A 0% plan shows a tiny number
`=RATE(12,-100,1200)` returned 1.7E-10 in LibreOffice, not exactly 0. The template returns 0 and No interest when the payments add up to the price.

### The rate is negative
The payments add up to less than the cash price. The Vinyl cutter's 6 x $95.00 is $570.00 on a $600.00 price, so the file shows Costs less than cash.

### A one-payment plan due at signing shows Check
Nothing is financed, so there is no rate. RATE still returned -0.999999999999998 there in LibreOffice instead of an error, so the template shows Check.

## Template

[Download the .xlsx](templates/tidy-tabs-payment-plan-interest-rate.xlsx)

The file has two tabs. **Offers** holds rows 5 to 54: Offer, Cash price, Down payment, Upfront fee, Monthly payment, Number of payments and First payment at signing (inputs in **A:G**), then Total paid, Extra vs cash, Monthly rate, Rate x 12, Effective yearly rate and Status (formulas in **H:M**), with the summary in **O4:P8**. It prints Landscape on US Letter. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Status turns green for No interest and red for Check. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
