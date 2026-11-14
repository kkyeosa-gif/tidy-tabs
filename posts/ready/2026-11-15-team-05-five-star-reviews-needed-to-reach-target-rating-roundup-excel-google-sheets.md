---
title: How to Calculate How Many 5-Star Reviews You Need for a Target Rating in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 18 workbooks and 1,584 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Find how many 5-star reviews lift your average to a target with =ROUNDUP(reviews*(target-current)/(5-target),0), and how many 1-star reviews it can absorb. Free template.
image_prompts: An open cloth-bound guest book with blank cream pages on the front table of a craft fair booth, a fountain pen resting in the gutter and a small ceramic vase of dried lavender beside it, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Open guest book with blank pages, a fountain pen and dried lavender at a booth
threads: 4.72 stars over 183 reviews. Reaching 4.80 takes 74 more 5-star reviews.\n=ROUNDUP(reviews*(target-current)/(5-target),0) does the math, and it has to be ROUNDUP: 73.2 means 74.\nA 5.0 target is not possible once any lower review exists.
---
Find the number of new 5-star reviews it takes to reach a target average with `=ROUNDUP(reviews*(target-current)/(5-target),0)`, and the number of 1-star reviews your average can take before it drops below the target with `=ROUNDDOWN(reviews*(current-target)/(target-1),0)`. In the made-up sample, a 4.72 average over 183 reviews needs 74 more 5-star reviews to reach 4.80 (73.2 rounded up, ending at 4.8006), and the download shows Not possible for a target of 5.0 instead of an error.

The math assumes your rating is a plain average of every review. Your platform may calculate or round the rating it shows differently, so treat the result as the plain-average answer.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (18 workbooks, 1,584 cell comparisons, 0 mismatches), plus the edge cases below. ROUNDUP, ROUNDDOWN and SUMPRODUCT also exist in Excel and Google Sheets, but neither app was opened for this post; how they round results that land a hair above a whole number is the open question.

![Shop ratings with current average, review count, target and the 5-star reviews still needed](images/2026-11-15-team-05-five-star-reviews-needed-to-reach-target-rating-roundup-excel-google-sheets/five-star-reviews-needed-template.png) *Seven fictional shops and listings: 74 five-star reviews take the candle shop from 4.72 to 4.80, and a 5.00 target shows Not possible (the Average after column and the summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*

## Steps

Use one row per shop or product, starting in row 5 of the **Calculator** tab.

1. Type the shop or product name in **A5**.
2. Type the current average in **B5** (4.72), the number of reviews in **C5** (183) and the target average in **D5** (4.80).
3. In **E5**, type `=ROUNDUP(C5*(D5-B5)/(5-D5),0)` for the 5-star reviews needed. Your reviews sit C5 x (D5 - B5) stars below the target in total, and each new 5-star review sits (5 - D5) above it. Use ROUNDUP, not ROUND: 73.2 means 74.
4. In **F5**, type `=(B5*C5+5*E5)/(C5+E5)` to see the average after those reviews. It should land at or just above the target.
5. In **G5**, type `=ROUNDDOWN(C5*(B5-D5)/(D5-1),0)` for the 1-star reviews the average can take and still stay at or above the target. Each 1-star review sits (D5 - 1) below the target, so this time you round down.
6. Fill **E5:G5** down to the last row.

If you have the count of reviews at each star level, use it instead of the shown average. On the **From star counts** tab, type the counts for 5, 4, 3, 2 and 1 stars in **B5:B9**, then get the average with `=SUMPRODUCT(A5:A9,B5:B9)/SUM(B5:B9)` and the total with `=SUM(B5:B9)`. Row 11 of the Calculator reads both. This is the same math as a [weighted average cost](https://tidytabs.blogspot.com/2026/10/how-to-calculate-weighted-average-cost.html).

The template wraps these formulas in checks. **H5** sets the Status first, then E5 returns 0 when you are already at the target and Not possible for a target of 5.00. G5 shows Already below when the average is under the target.

The file also wraps each division in `ROUND(...,6)` before ROUNDUP or ROUNDDOWN. A quotient that should be a whole number can come out a tiny bit above it, such as 60.0000000000003 for the Soap listing. LibreOffice still returned 60 without the guard, and Excel and Sheets were not checked, so treat it as a safety step rather than a fix for any app.

Help pages: Microsoft's [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) and [ROUNDDOWN function](https://support.microsoft.com/en-us/office/rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53), and Google's [ROUNDUP](https://support.google.com/docs/answer/3093443) and [ROUNDDOWN](https://support.google.com/docs/answer/3093442).

## Example

The shops, listings and star counts are fictional.

| Shop or product | Current | Reviews | Target | 5-star needed | Average after | 1-star before below | Status |
|---|---|---|---|---|---|---|---|
| Candle shop | 4.72 | 183 | 4.80 | 74 | 4.8006 | Already below | Below target |
| Soap listing | 4.70 | 30 | 4.90 | 60 | 4.9000 | Already below | Below target |
| Tote bag listing | 4.60 | 100 | 4.80 | 100 | 4.8000 | Already below | Below target |
| Mug listing | 4.85 | 40 | 4.80 | 0 | 4.8500 | 0 | At or above target |
| Print shop | 4.95 | 240 | 4.90 | 0 | 4.9500 | 3 | At or above target |
| Jewelry shop | 4.90 | 50 | 5.00 | Not possible | | Already below | Not possible |
| Whole candle shop (from star counts) | 4.7377 | 183 | 4.80 | 57 | 4.8000 | Already below | Below target |

The Print shop at 4.95 can take 3 one-star reviews and stay at 4.9012. A 4th drops it to 4.8852. The Mug listing is above its target but has no room: its cushion is 0.

The last row comes from star counts of 152, 21, 6, 1 and 3 (5 stars down to 1). That is 183 reviews averaging 4.7377, so the whole candle shop needs 57 five-star reviews, not the 74 the rounded 4.72 suggests.

The summary in **J5:K9** shows 7 rows listed, 4 below target, 2 at or above target, 1 not possible and 0 rows showing Check.

## Troubleshooting

### #DIV/0! for a target of 5.0
`=ROUNDUP(100*(5-4.9)/(5-5),0)` divides by zero. No number of 5-star reviews brings an average with any lower review up to exactly 5.00, so the template shows Not possible instead.

### The answer seems one review short
You used ROUND instead of ROUNDUP. The Candle shop needs 73.2 reviews; 73 leaves it at 4.7998, just under 4.80, and 74 reaches 4.8006.

### The answer changes when you use star counts
The shown average is rounded. At 120 reviews, an average shown as 4.7 that is really 4.74 needs 60 five-star reviews to reach 4.80 using 4.7, but only 36 using 4.74. The candle shop's 4.72 gives 74 and its star-count average of 4.7377 gives 57. Use star counts when you have them.

### A row shows Check
One of the inputs cannot work: 0 reviews, a review count typed as text or with decimals, an average above 5 or below 1, or a target of 1 or less (or above 5). Retype the number in the yellow cell.

### The star-count average shows Check or No reviews
A count on the From star counts tab is text, negative or not a whole number (Check), or every count is 0 (No reviews). Row 11 of the Calculator then shows Check too.

## Template

[Download the .xlsx](templates/tidy-tabs-reviews-needed-for-rating.xlsx)

The file has three tabs. **Calculator** holds rows 5 to 24: Shop or product, Current average, Number of reviews and Target average (inputs in **A:D**; row 11 reads the star counts), then 5-star reviews needed, Average after those, 1-star reviews before it drops below target and Status (formulas in **E:H**), with the summary in **J5:K9**. **From star counts** takes the number of reviews at each star level and returns the total and the average. **How to use** explains each column. The Calculator and From star counts tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Status turns green for At or above target and red for Not possible and Check. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
