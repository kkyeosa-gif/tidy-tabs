---
title: How to Calculate Dimensional Weight and Billable Weight in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with Python math, including the 12 sample boxes at divisors 139 and 166, an exactly divisible box (139 x 1 x 1 at divisor 139), fractional sizes, a tie between the two weights, an actual weight of 0, and zero, negative, text or blank sizes and divisors. No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Calculate dimensional weight and billable weight in Excel and Google Sheets with ROUNDUP and MAX. Free template with a box list and a divisor cell.
image_prompts: A kraft cardboard shipping box on a small digital postal scale beside a metal tape measure, a roll of packing tape and a second flat mailer box in a home workshop, soft daylight, no screens, no readable text, no logos
image_alt: Cardboard box on a postal scale next to a tape measure and packing tape
threads: A 12 x 10 x 8 inch box that weighs 4.2 lb gets billed as 7 lb.\n=ROUNDUP(B4*B5*B6/B8,0) gives the dimensional weight, =MAX(B11,B12) picks the larger one.\nDivisor is your carrier's call, so check it. The sample uses 139.
---
Divide length x width x height (inches) by your carrier's divisor and round up with `=ROUNDUP(B4*B5*B6/B8,0)`, then bill the larger of that and the rounded-up actual weight with `=MAX(B11,B12)`. A 12 x 10 x 8 inch box with a divisor of 139 comes to 7 lb, so a 4.2 lb package is billed at 7.

Here B4:B6 are the three sizes, B7 is the actual weight in pounds, B8 is the divisor, B11 is the dimensional weight from the first formula and B12 is the actual weight rounded up with `=ROUNDUP(B7,0)`. The divisor is not a universal number. Confirm the current one with your carrier before you rely on any result.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Calculator sheet with a 12 x 10 x 8 inch box, 4.2 lb actual weight, divisor 139 and a billable weight of 7 lb](images/2026-10-25-team-02-dimensional-weight-billable-weight-roundup-max-excel-google-sheets/dimensional-weight-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: a light box that takes up a lot of room can be billed as if it were heavier. Compare the volume-based weight with the scale weight and use the bigger one.

1. In **B4**, **B5** and **B6**, type the length, width and height in inches, such as 12, 10 and 8.
2. In **B7**, type the actual weight in pounds, such as 4.2.
3. In **B8**, type the divisor from your carrier's current dimensional weight rule. The sample uses 139.
4. In **B10**, type `=B4*B5*B6` for the volume in cubic inches: 960.
5. In **B11**, type `=ROUNDUP(B4*B5*B6/B8,0)`. This is the dimensional weight: 960 cubic inches / 139 = 6.91, rounded up to 7.
6. In **B12**, type `=ROUNDUP(B7,0)`. This is the actual weight rounded up: 5.
7. In **B13**, type `=MAX(B11,B12)`. This is the billable weight: the larger of the dimensional weight (7) and the actual weight rounded up (5).
8. In **B14**, type `=IF(B11>B12,"Dimensional","Actual")` to see which weight applies.

How the formulas read:

- `ROUNDUP(...,0)` always rounds up to a whole pound. 6.91 becomes 7, and 4.2 becomes 5. A normal `ROUND` would turn 6.91 into 7 but 4.2 into 4, which can under-bill.
- `MAX` returns the larger of its two numbers.
- The `IF` in step 8 says "Actual" when the two weights are equal. Change `>` to `>=` if you prefer a tie to say "Dimensional".

Whether your carrier rounds up, which services use the divisor and what it is today are carrier rules. This post does not state them.

Help pages for the functions: [ROUNDUP in Excel](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7), [MAX in Excel](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098) and [ROUNDUP in Google Sheets](https://support.google.com/docs/answer/3093443). Those pages describe the functions, not our test.

### What the template adds

The built formulas are longer than the bare ones. The dimensional weight in B11 is wrapped as `ROUNDUP(ROUND(B4*B5*B6/B8,8),0)`. `ROUND(...,8)` first clears tiny floating point error, so a box that divides exactly, such as 139 x 1 x 1 at divisor 139, gives 1 and not 2.

Each result in B10:B14 also checks the inputs. If a size or the divisor is blank, text, 0 or negative, every result cell shows "Check inputs" instead of an error.

### In Excel

The formulas work as typed. ROUNDUP, MAX and IF are old functions.

### In Google Sheets

The same formulas should work as typed.

## Example

Sample data is fictional: 12 boxes from a small gift shop, with sizes in inches and weights in pounds, all at a divisor of 139.

| Box | L x W x H (in) | Actual (lb) | Rounded | Dimensional (lb) | Billable (lb) | Applies |
|---|---|---|---|---|---|---|
| Candle set, small | 8 x 6 x 6 | 2.4 | 3 | 3 | 3 | Actual |
| Candle set, large | 12 x 10 x 8 | 4.2 | 5 | 7 | 7 | Dimensional |
| Ceramic mug pair | 9 x 9 x 6 | 3.1 | 4 | 4 | 4 | Actual |
| Tote bag order | 14 x 12 x 4 | 1.8 | 2 | 5 | 5 | Dimensional |
| Soap gift box | 7 x 5 x 4 | 1.2 | 2 | 2 | 2 | Actual |
| Print, framed 16x20 | 22 x 18 x 3 | 6.5 | 7 | 9 | 9 | Dimensional |
| Pillow cover bundle | 16 x 14 x 10 | 2.0 | 2 | 17 | 17 | Dimensional |
| Jewelry mailer | 6 x 4 x 2 | 0.4 | 1 | 1 | 1 | Actual |
| Tea tin case | 10 x 10 x 10 | 8.7 | 9 | 8 | 9 | Actual |
| Planter, medium | 12 x 12 x 12 | 9.0 | 9 | 13 | 13 | Dimensional |
| Sticker roll tube | 24 x 4 x 4 | 1.1 | 2 | 3 | 3 | Dimensional |
| Wall art, large | 30 x 20 x 4 | 5.5 | 6 | 18 | 18 | Dimensional |

The total billable weight is 91 lb. Seven boxes are billed by dimensional weight and five by actual weight. The Tea tin case shows the other side: it is dense, so its 9 lb actual weight beats the 8 lb dimensional weight.

Change the divisor to 166 and the total drops to 81 lb, still seven Dimensional and five Actual. The 12 x 10 x 8 box becomes 960 / 166 = 5.78, which rounds up to 6. These are the template's math at two divisors, not a statement about what any carrier uses.

## Troubleshooting

### Every result says Check inputs
A size or the divisor is blank, text, 0 or negative. Retype it as a plain positive number. An actual weight of 0 is allowed and bills the dimensional weight.

### The dimensional weight is one pound too high
Check for a box that should divide exactly, such as 139 x 1 x 1 inches at divisor 139. The answer should be 1. The template wraps the division in `ROUND(...,8)` before `ROUNDUP` to guard against tiny floating point error, and the exact case gave 1 in LibreOffice. If your result is higher than hand math says, add the same wrap or recheck the divisor.

### The total on the Packages tab does not match the rows
The check cell **K6** says CHECK. A listed box has a blank or text size or weight, or numbers sit on a row with no box name. Fill in the row or clear the stray numbers.

### Results are lower than the carrier's quote
The divisor in the sheet may differ from the one your carrier uses for that service, or the carrier measures the box differently. Confirm the current divisor and measuring rules on the carrier's own site.

### Fractional sizes look off
12.5 x 9.5 x 7.25 inches is 861 cubic inches (to the nearest whole), so the dimensional weight is 7 at divisor 139. Type the real decimals. Rounding each size yourself first can change the answer.

## Template

[Download the .xlsx](templates/tidy-tabs-dimensional-weight.xlsx)

The file has three tabs. **Calculator** holds the yellow inputs in **B4:B8** (length, width, height, actual weight and divisor) and the gray results in **B10:B14**: volume, dimensional weight, actual weight rounded up, billable weight and which one applies. **Packages** holds 12 fictional boxes in yellow in **A2:E100**, with dimensional weight, billable weight and the applying weight in columns F to H. The summary in **J1:K6** shows the box counts, total billable pounds, the divisor in use and a check. **How to use** has short notes.

The template has no carrier rates or rules. For shipping prices by weight, use your carrier's current price table. The boxes are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Packages tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
