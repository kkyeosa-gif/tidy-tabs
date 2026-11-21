---
threads_url: https://www.threads.com/@tin_ylab/post/DgATmFCm8HO
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-estimate-project-hours-with.html
title: How to Estimate Project Hours With a Three-Point (PERT) Estimate in Excel and Google Sheets
labels: excel-formulas, small-business, freelance, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 19 workbooks and 2,565 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Get expected hours per task with =(O+4*M+P)/6, combine spreads with =SQRT(SUMSQ(spreads)) and turn the total into a quote range. Free template.
image_prompts: A freelancer's home studio desk seen from above with hand-drawn paper website wireframe sketches made only of boxes and lines, a row of blank index cards laid out in three short columns, and a small wind-up kitchen timer beside them, warm lamp light, realistic photo, no screens, no readable text, no logos
image_alt: Paper wireframe sketches, blank index cards in three columns and a wind-up timer
threads: Adding up every task's spread to get your project buffer?\nIn a made-up six-task web project that gives 8.5 hours. =SQRT(SUMSQ(spreads)) gives 4.14.\nEach task's expected hours come from =(O+4*M+P)/6, so 8, 12 and 20 is 12.67, not 12.
---
Type a best case (O), most likely (M) and worst case (P) for each task, get its expected hours with `=(O+4*M+P)/6` and its spread with `=(P-O)/6`, then combine the spreads for the whole project with `=SQRT(SUMSQ(spreads))`. In the made-up six-task web project, the expected total is 45.83 hours with a project spread of 4.14 hours, so at $85.00 an hour the quote range runs $3,895.83, $4,247.73 with one spread added and $4,599.62 with two; simply adding the task spreads would give 8.5 hours instead.

The weighting counts the most likely guess four times and the best and worst case once each. One task guessed at 8, 12 and 20 hours comes out at 12.67 expected hours, not 12, because the worst case sits further from the middle than the best case.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (19 workbooks, 2,565 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (SUM, SUMIF, SUMSQ, SQRT, ROUND, COUNT) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six web project tasks with best, likely and worst hours beside totals and quotes](images/2026-11-17-team-03-three-point-estimate-project-hours-quote-range-pert-sumsq-excel-google-sheets/three-point-estimate-hours-and-quotes-template.png) *A fictional six-task web project at $85.00 an hour: 45.83 expected hours with a 4.14 hour spread gives quotes of $3,895.83, $4,247.73 and $4,599.62 (the Status column is hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Estimate** tab, one task per row starting in row 5.

1. Type your hourly rate in **B1** (the sample uses $85.00).
2. Type the task name in **A5**, then the best case hours in **B5**, the most likely hours in **C5** and the worst case hours in **D5**. Best case must be the smallest number and worst case the largest. Decimals such as 0.25 for a quarter hour work.
3. In **E5**, type `=(B5+4*C5+D5)/6` for the expected hours.
4. In **F5**, type `=(D5-B5)/6` for the spread, one sixth of the gap between best and worst case.
5. Fill **E5:F5** down to the last task.
6. Add the expected hours with `=SUM(E5:E44)`. That is the expected total in **J8**.
7. Combine the spreads with `=SQRT(SUMSQ(F5:F44))` in **J9**. SUMSQ squares each spread and adds the squares, and SQRT takes the square root of that sum. This is the usual way to combine spreads of tasks that do not depend on each other.
8. Add one spread with `=J8+J9` in **J10** and two with `=J8+2*J9` in **J11**. Treat those extra hours as a buffer, not a guarantee. The sheet makes no probability claim.
9. Turn hours into dollars with `=ROUND(J8*$B$1,2)` in **J13**, and the same on J10 and J11 in **J14** and **J15**.

The template wraps these in checks. **G5** shows a Status of OK or Check, and E5 and F5 stay blank on a Check row, so every total counts OK rows only. The quotes stay blank until B1 holds a number above 0.

Which quote you send, and at what rate, is your call. The sheet does the math; it is not pricing advice. To set the rate itself, see [how to calculate a freelance hourly rate](https://tidytabs.blogspot.com/2026/10/how-to-calculate-freelance-hourly-rate.html). After the job, [find your real hourly rate on a fixed-price project](https://tidytabs.blogspot.com/2026/11/how-to-find-your-real-hourly-rate-on.html) to see how the estimate held up.

Help pages: Microsoft's [SUMSQ function](https://support.microsoft.com/en-us/office/sumsq-function-e3313c02-51cc-4963-aae6-31442d9ec307) and [SQRT function](https://support.microsoft.com/en-us/office/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf).

## Example

The project, tasks, hours and rate are fictional.

| Task | Best case (O) | Most likely (M) | Worst case (P) | Expected hours | Spread | Status |
|---|---|---|---|---|---|---|
| Discovery call and brief | 1 | 2 | 4 | 2.17 | 0.50 | OK |
| Wireframes | 4 | 6 | 10 | 6.33 | 1.00 | OK |
| Design mockups | 8 | 12 | 20 | 12.67 | 2.00 | OK |
| Build pages | 12 | 16 | 30 | 17.67 | 3.00 | OK |
| Revisions | 2 | 4 | 12 | 5.00 | 1.67 | OK |
| Launch and handoff | 1 | 2 | 3 | 2.00 | 0.33 | OK |

The summary in **I4:J18** shows:

| Summary | Value |
|---|---|
| Tasks counted | 6 |
| Best case total | 28.00 hours |
| Most likely total | 42.00 hours |
| Worst case total | 79.00 hours |
| Expected total | 45.83 hours |
| Project spread | 4.14 hours |
| Expected + 1 spread | 49.97 hours |
| Expected + 2 spreads | 54.11 hours |
| Task spreads added (do not use) | 8.50 hours |
| Quote at expected hours | $3,895.83 |
| Quote at + 1 spread | $4,247.73 |
| Quote at + 2 spreads | $4,599.62 |
| Rows showing Check | 0 |
| CHECK | OK |

The cells hold more decimals (45.8333 and 4.13991); the quotes are rounded to cents from the full values.

## Troubleshooting

### A task is missing from the total
Its Status is Check. The hours are out of order (a 5 / 3 / 8 row has a best case above the most likely), the best case is negative, a number is blank, or a number was typed as text such as "2". Fix the row and it comes back into every total.

### The project buffer looks too big
You added the task spreads. In the sample, `=SUM(F5:F44)` gives 8.50 hours (shown in **J12** for comparison), while `=SQRT(SUMSQ(F5:F44))` in **J9** gives 4.14.

### The quotes are blank
The rate in **B1** is blank, 0 or text such as "$85.00". Type the rate as a plain number, 85, and format the cell as currency.

### CHECK says Check but no row shows Check
Something is typed below row 44, which the totals do not read (**J17** counts it), or no task is listed yet. Move the task into an empty row between 5 and 44.

## Template

[Download the .xlsx](templates/tidy-tabs-three-point-project-estimate.xlsx)

The file has two tabs. **Estimate** has the hourly rate in **B1** and rows 5 to 44: Task, Best case hours (O), Most likely hours (M) and Worst case hours (P) (inputs in **A:D**), then Expected hours, Spread and Status (formulas in **E:G**). The summary in **I4:J18** holds the totals, spreads, quotes and checks shown in the example above. **How to use** explains each formula and when a row shows Check. The Estimate tab prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. A Status of Check turns red, and so does the CHECK cell when it is not OK. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
