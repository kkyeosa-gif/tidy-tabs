---
title: How to Calculate a Tiered Discount Based on Order Total in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 228 workbooks and 1,150,260 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Spend-and-save tiers in Excel and Google Sheets: COUNTIF finds the tier, INDEX pulls the dollars off and next threshold, and one cell shows what to add.
image_prompts: A boutique checkout counter during a holiday weekend sale with a small clear acrylic sign holder holding a blank card, a stack of folded knit scarves tied with ribbon beside a kraft paper shopping bag, and a few blank plastic gift cards fanned out on the counter, warm indoor light, realistic photo, no screens, no readable text, no logos
image_alt: Checkout counter with a blank sign holder, ribbon-tied scarves and fanned gift cards
threads: Spend-and-save tiers without nested IFs: =COUNTIF(tiers,"<="&subtotal) is the tier number, and INDEX(tiers,tier+1) is the next threshold.\nA $148.25 order gets $25.00 off and is $1.75 short of $40.00 off.\nThe sample sheet flags every order within $20.00 of a bigger discount: 7 of the 13 in the promo.
---
Count how many thresholds the order has passed with `=COUNTIF(Tiers!$A$5:$A$14,"<="&E5)`, then pull that tier's dollars off with `=INDEX(Tiers!$B$5:$B$14,F5)` and the next threshold with `=INDEX(Tiers!$A$5:$A$14,F5+1)`. With $10.00 off $50.00, $25.00 off $100.00 and $40.00 off $150.00, a $148.25 order gets $25.00 off and is $1.75 away from $40.00 off.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (228 workbooks, 1,150,260 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets were not opened for this post.

![Fifteen sample orders with tier reached, dollars off and amount to the next tier](images/2026-11-26-team-01-spend-and-save-tiered-discount-amount-to-reach-next-tier-countif-index-excel-google-sheets/spend-and-save-tiered-discount-template.png) *Orders tab of a fictional holiday promo: 15 orders from 11/26/2026 to 12/02/2026 with the eligible subtotal, tier reached, dollars off, the next threshold and the amount to add, green Yes where an order is within $20.00 of a bigger discount, orange where it falls outside the promo dates, and the summary with $345.00 off on $1,811.38 of merchandise. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

This is a spend-and-save promo: the more an order spends, the more dollars come off. The dollars off come off the merchandise subtotal only. Shipping and sales tax are not in the sheet, and this post gives no tax advice.

The steps are the same in Excel and Google Sheets.

1. On the **Tiers** tab, type one tier per row from row 5, lowest first: Spend at least in **A**, Dollars off in **B**. Up to 10 tiers fit in rows 5 to 14. Leave no blank rows between tiers.
2. Type the promo start date in **B1** and the end date in **B2**. Both days are included.
3. Check **E5:E14** on the Tiers tab. Each row must say OK: thresholds and dollars off both go up, the dollars off stay below the threshold, and no amount has more than 2 decimals. **D** shows the discount as a percent of the threshold, so you can see which tier is the best value. **H8** is the tab's CHECK cell.
4. On the **Orders** tab, type one order per row from row 5: Order # in **A**, Order date in **B**, Merchandise subtotal in **C**, and anything that does not count toward the promo (gift cards, items already on sale) in **D**. Leave D blank if everything counts.
5. **E5** gives the eligible subtotal with `=IF(OR(N5="",LEFT(N5,5)="Check"),"",ROUND(C5-N(D5),2))`. The ROUND keeps the amount in whole cents before it meets the thresholds. It is a safeguard: subtracting 20.10 from 120.10 gave exactly 100 in LibreOffice, so no miss was seen without it. See the ROUND help from [Microsoft](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) or [Google](https://support.google.com/docs/answer/3093440).
6. **F5** finds the tier with `=IF(E5="","",IF(N5="OK",COUNTIF(Tiers!$A$5:$A$14,"<="&E5),0))`. Because the thresholds go up, the number of thresholds at or below the order is its tier number. The `"<="` means an order of exactly $100.00 reaches the $100.00 tier. See the COUNTIF help from [Microsoft](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) or [Google](https://support.google.com/docs/answer/3093480).
7. **G5** pulls the dollars off with `=IF(F5="","",IF(F5=0,0,INDEX(Tiers!$B$5:$B$14,F5)))`. **H5** is the total after discount, `=IF(G5="","",ROUND(C5-G5,2))`, and **I5** the discount as a percent of the whole order, `=IF(G5="","",IF(C5=0,0,G5/C5))`. See the INDEX help from [Microsoft](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) or [Google](https://support.google.com/docs/answer/3098242).
8. **J5** gives the next threshold with `=IF(OR(N5<>"OK",F5>=Tiers!$H$4),"",INDEX(Tiers!$A$5:$A$14,F5+1))`. It stays blank at the top tier, since `Tiers!$H$4` counts the tiers.
9. **K5** shows how much more would reach that tier, `=IF(J5="","",ROUND(J5-E5,2))`, and **L5** what that tier takes off, `=IF(J5="","",INDEX(Tiers!$B$5:$B$14,F5+1))`.
10. Type a nudge amount in **B1** on the Orders tab ($20.00 in the sample). **M5** says Yes when the amount to add is at or below it: `=IF(K5="","",IF(K5<=$B$1,"Yes",""))`.
11. **N5** shows OK, Outside promo dates, or a Check message for a text date, a text or negative amount, a not-eligible amount above the subtotal, or a broken Tiers tab.
12. Read the summary in **P4:Q13**: orders in the promo, orders that reached a tier, merchandise in the promo, total dollars off, discount % of promo merchandise (`=IF(Q6=0,"",Q7/Q6)`), orders within the nudge amount, outside promo dates, rows showing Check, rows typed below row 504, and a final CHECK cell in **Q13**.

For a single free shipping cutoff instead of tiers, see [How to Check a Free Shipping Threshold and Show the Amount to Add](https://tidytabs.blogspot.com/2026/10/how-to-check-free-shipping-threshold.html). For a per-unit price that drops with quantity, see [How to Look Up Wholesale Price Tiers by Quantity](https://tidytabs.blogspot.com/2026/10/how-to-look-up-wholesale-price-tiers-by.html).

## Example

The orders below are made up: a holiday weekend promo from 11/27/2026 to 12/01/2026 with four tiers ($10.00 off $50.00, $25.00 off $100.00, $40.00 off $150.00, $75.00 off $250.00) and a $20.00 nudge amount.

| Order | Date | Subtotal | Not eligible | Eligible | Tier | Off | After | Next at | Add | Next off | Nudge | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1101 | 11/26/2026 | $112.40 | | $112.40 | 0 | $0.00 | $112.40 | | | | | Outside promo dates |
| 1102 | 11/27/2026 | $49.99 | | $49.99 | 0 | $0.00 | $49.99 | $50.00 | $0.01 | $10.00 | Yes | OK |
| 1103 | 11/27/2026 | $50.00 | | $50.00 | 1 | $10.00 | $40.00 | $100.00 | $50.00 | $25.00 | | OK |
| 1104 | 11/27/2026 | $99.99 | | $99.99 | 1 | $10.00 | $89.99 | $100.00 | $0.01 | $25.00 | Yes | OK |
| 1105 | 11/27/2026 | $100.00 | | $100.00 | 2 | $25.00 | $75.00 | $150.00 | $50.00 | $40.00 | | OK |
| 1106 | 11/28/2026 | $120.10 | $20.10 | $100.00 | 2 | $25.00 | $95.10 | $150.00 | $50.00 | $40.00 | | OK |
| 1107 | 11/28/2026 | $136.50 | | $136.50 | 2 | $25.00 | $111.50 | $150.00 | $13.50 | $40.00 | Yes | OK |
| 1108 | 11/28/2026 | $148.25 | | $148.25 | 2 | $25.00 | $123.25 | $150.00 | $1.75 | $40.00 | Yes | OK |
| 1109 | 11/29/2026 | $175.00 | $50.00 | $125.00 | 2 | $25.00 | $150.00 | $150.00 | $25.00 | $40.00 | | OK |
| 1110 | 11/29/2026 | $236.80 | | $236.80 | 3 | $40.00 | $196.80 | $250.00 | $13.20 | $75.00 | Yes | OK |
| 1111 | 11/30/2026 | $262.15 | | $262.15 | 4 | $75.00 | $187.15 | | | | | OK |
| 1112 | 11/30/2026 | $38.00 | | $38.00 | 0 | $0.00 | $38.00 | $50.00 | $12.00 | $10.00 | Yes | OK |
| 1113 | 12/01/2026 | $84.60 | | $84.60 | 1 | $10.00 | $74.60 | $100.00 | $15.40 | $25.00 | Yes | OK |
| 1114 | 12/01/2026 | $310.00 | $25.00 | $285.00 | 4 | $75.00 | $235.00 | | | | | OK |
| 1115 | 12/02/2026 | $160.00 | | $160.00 | 0 | $0.00 | $160.00 | | | | | Outside promo dates |

Order 1109 spent $175.00 but $50.00 of it was a gift card, so it counts as $125.00 and stays in the $25.00 tier. Order 1106 lands on $100.00 exactly after its $20.10 exclusion and reaches the $25.00 tier, which is 20.8% of its $120.10 subtotal.

The summary shows 13 orders in the promo, 11 that reached a tier, $1,811.38 of merchandise and $345.00 off, 19.0%. Seven orders stopped within $20.00 of a bigger discount, 2 fell outside the promo dates, and CHECK is OK.

## Troubleshooting

### Every order says Check: Tiers tab
A row on the Tiers tab is out of order, blank in the middle of the list, typed as text, or the promo dates are reversed. Find the row whose Status in column **E** is not OK (or **H6**, Promo dates) and fix it; the orders recalculate on their own.

### An order gets no discount
Its Status says Outside promo dates (1101 on 11/26/2026 and 1115 on 12/02/2026 in the sample), or the not-eligible amount in **D** took it under the first threshold. Check the dates in **Tiers!B1:B2** and the amount in D.

### Nudge? is Yes on every row
**B1** on the Orders tab was typed as text, such as "20" with an apostrophe. Retype it as a number; CHECK in **Q13** stays Check until you do.

### An order exactly on a threshold gets the lower tier
Check the comparison in your copy of the tier formula. With `"<"&` instead of `"<="&`, COUNTIF counts only thresholds below the order, so an order of exactly $100.00 would not count the $100.00 threshold and would land one tier lower. The template uses `"<="`, so Spend at least $100.00 includes $100.00.

## Template

[Download the .xlsx](templates/tidy-tabs-spend-and-save-tiers.xlsx)

**Orders** takes 500 orders (rows 5 to 504) with the eligible subtotal, tier, dollars off, total after discount, next tier, amount to add, nudge flag, Status and the summary in P4:Q13, which also catches rows typed below row 504. **Tiers** holds up to 10 tiers, the promo dates and its own CHECK cell. **How to use** explains each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Nudge Yes cells are light green, Outside promo dates is light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
