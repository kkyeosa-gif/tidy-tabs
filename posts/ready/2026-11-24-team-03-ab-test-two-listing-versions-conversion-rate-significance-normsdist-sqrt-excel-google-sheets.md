---
title: How to Check If an A/B Test Result Is Significant in Excel and Google Sheets
labels: excel-formulas, small-business, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 23 workbooks and 11,684 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. NORMSDIST and NORMSINV results were checked against Python's math.erf and NormalDist only.
search_description: Get z from the pooled rate with SQRT and a two-tailed p-value with NORMSDIST to see if version B's conversion rate beats A, plus visits needed per version.
image_prompts: Two printed product photos of the same handmade ceramic mug lying side by side on a small studio table, one an overhead flat lay and one a styled lifestyle shot, with the real mug standing between them, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Two printed product photos of one mug, flat lay and lifestyle, side by side
threads: Is 3.04% vs 2.01% a real difference or just noise?\nGet z from the pooled rate with SQRT, then p with =2*(1-NORMSDIST(ABS(z))).\nThe sample photo test, 37 orders from 1,840 visits vs 58 from 1,910, gives p 0.046. Likely real at 95%, not at 99%.
---
Get z with `=(rate_b-rate_a)/SQRT(pooled*(1-pooled)*(1/visits_a+1/visits_b))`, where pooled is all orders divided by all visits, and the two-tailed p-value with `=2*(1-NORMSDIST(ABS(z)))`; the result is likely a real difference when p is at or below 0.05 for 95% confidence. In the made-up listing photo test, 37 orders from 1,840 visits (2.01%) against 58 from 1,910 (3.04%) gives z 2.00 and p 0.046, a likely real difference at 95% but not at 99%.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (23 workbooks, 11,684 cell comparisons, 0 mismatches), plus the edge cases below. NORMSDIST and NORMSINV results were checked against Python's math.erf and NormalDist only. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Five sample listing tests with conversion rates, z, p-values and result labels](images/2026-11-24-team-03-ab-test-two-listing-versions-conversion-rate-significance-normsdist-sqrt-excel-google-sheets/ab-test-significance-template.png) *Tests tab at 95% confidence with five fictional tests: the listing photo test (37 orders from 1,840 visits vs 58 from 1,910) gets p 0.0457, Likely a real difference, while the $24.00 vs $26.00 price test (p 0.4132) would need 13,959 visits per version. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

This is a pooled two-proportion z test with a normal approximation, two-tailed. It suits visit counts in the hundreds and up, such as your own Etsy listing stats. It does not fit revenue per visitor.

The steps are the same in Excel and Google Sheets.

1. On the **Tests** tab, type the confidence level in **B1** (95%), the power for the visits-needed column in **B2** (80%) and the fewest orders per version you want before trusting a test in **B3** (10).
2. Type one test per row from row 6: Test name in **A**, Visits A in **B**, Orders A in **C**, Visits B in **D** and Orders B in **E**.
3. In **F6** and **G6**, get each conversion rate with `=C6/B6` and `=E6/D6`. **H6** shows the lift with `=G6/F6-1`, left blank when Rate A is 0.
4. In **I6**, get the pooled rate with `=(C6+E6)/(B6+D6)`: all orders over all visits, as if A and B were the same.
5. In **J6**, get the standard error with `=SQRT(I6*(1-I6)*(1/B6+1/D6))`. See the [SQRT help from Microsoft](https://support.microsoft.com/en-us/office/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf).
6. In **K6**, get z with `=(G6-F6)/J6`.
7. In **L6**, get the two-tailed p-value with `=2*(1-NORMSDIST(ABS(K6)))`. See the [NORMSDIST help from Microsoft](https://support.microsoft.com/en-us/office/normsdist-function-463369ea-0345-445d-802a-4ff0d6ce7cac).
8. In **M6**, read the result with `=IF(L6<=1-$B$1,"Likely a real difference","Not enough evidence yet")`. At 95% confidence the cutoff is p 0.05.
9. **N6** estimates the visits each version needs to catch a gap this size at the power in **B2**: `=ROUNDUP((NORMSINV(1-(1-conf)/2)*SQRT(2*p*(1-p))+NORMSINV(power)*SQRT(rA*(1-rA)+rB*(1-rB)))^2/(rB-rA)^2,0)`, where p is the average of the two rates. It shows "Rates are equal" when they match and "Over 1,000,000" above that count.
10. **O6** is the Status. It shows Check for blank, text, decimal or negative counts, 0 visits, more orders than visits, no orders at all, every visit ordered or no test name, and F:N stay blank on those rows. It shows "Caution: few orders" when either version has fewer orders than **B3**.

**What the result does not say:** "Not enough evidence yet" is not proof that A and B convert the same. The 10-order caution is a setting you can change, not a statistical rule.

**Two cautions, not measured for this post:** a one-tailed p is half the two-tailed one, so a one-tailed calculator looks more confident than this sheet. Checking a test every day and stopping the first day p dips under 0.05 also overstates results.

## Example

The sample tests are made up. Settings: 95% confidence, 80% power, caution below 10 orders.

| Test | A visits / orders | B visits / orders | Rate A | Rate B | Lift | z | p-value | Result | Visits needed | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Listing photo: flat lay vs lifestyle | 1,840 / 37 | 1,910 / 58 | 2.01% | 3.04% | 51.0% | 1.998 | 0.0457 | Likely a real difference | 3,669 | OK |
| Price: $24.00 vs $26.00 | 1,200 / 30 | 1,180 / 36 | 2.50% | 3.05% | 22.0% | 0.818 | 0.4132 | Not enough evidence yet | 13,959 | OK |
| Newsletter subject line | 400 / 8 | 420 / 15 | 2.00% | 3.57% | 78.6% | 1.362 | 0.1731 | Not enough evidence yet | 1,721 | Caution: few orders |
| Free shipping banner | 950 / 19 | 960 / 19 | 2.00% | 1.98% | -1.0% | -0.033 | 0.9740 | Not enough evidence yet | Over 1,000,000 | OK |
| Shop photos: old vs new (one year) | 52,000 / 500 | 50,000 / 570 | 0.96% | 1.14% | 18.6% | 2.797 | 0.0052 | Likely a real difference | 51,246 | OK |

For the listing photo row, the pooled rate is 2.53% and the standard error 0.0051.

The summary in **Q5:R12** shows settings OK, 5 tests listed, 2 likely real differences, 3 with not enough evidence yet, 1 caution, 0 rows showing Check, 0 typed below row 55 and CHECK OK.

At 99% in **B1**, only the shop photos test stays a likely real difference.

## Troubleshooting

### z and p show #DIV/0!
Neither version had an order, so the standard error is 0. With the formulas copied without the Status guard, LibreOffice Calc showed #DIV/0!; the template shows "Check: no orders in either version" instead.

### A rate over 100% and p 0 say "Likely a real difference"
Orders were typed in the visits column, such as 520 orders on 500 visits. Without the guard that gave a 104% rate, p 0 and an error in visits needed. The template shows "Check: more orders than visits"; swap the two numbers back.

### Visits needed says Over 1,000,000 or Rates are equal
The two rates are almost or exactly the same, as in the 950 / 19 vs 960 / 19 banner test.

### The result flips when you change the confidence level
A p-value near the cutoff passes at one level and fails at the next. The listing photo test (p 0.0457) is a likely real difference at 95% and not enough evidence yet at 99%.

## Template

[Download the .xlsx](templates/tidy-tabs-ab-test-significance.xlsx)

The file has two tabs. **Tests** holds up to 50 tests in rows 6 to 55: you type **B1:B3** and **A:E**, and **F:O** hold the formulas above. The summary in **Q5:R12** checks the settings, counts results and Check rows, and ends in a CHECK cell. **How to use** explains each column. It prints Landscape on US Letter with row 5 repeated on every page.

Yellow cells are inputs and gray cells are formulas. "Likely a real difference" turns light green, "Caution: few orders" light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
