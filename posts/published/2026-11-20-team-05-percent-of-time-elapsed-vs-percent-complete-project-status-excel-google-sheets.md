---
threads_url: https://www.threads.com/@tin_ylab/post/DgDZ6BxlgA6
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-compare-percent-of-time-elapsed.html
title: How to Compare Percent of Time Elapsed With Percent Complete in Excel and Google Sheets
labels: excel-formulas, freelancers, project-planning
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 41 workbooks and 8,569 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Get percent of time elapsed with =MIN(1,MAX(0,(asof-start)/(due-start))) and compare it with percent done to mark Ahead, On track or Behind.
image_prompts: A freelance designer's home studio wall with a large paper wall calendar showing several days crossed out in red marker, a cork board beside it with printed design proofs pinned up and a few marked with a check, and a sand hourglass half run through on the desk below, soft afternoon light, realistic photo, no screens, no readable text, no logos
image_alt: Wall calendar with crossed-out days beside pinned design proofs and an hourglass
threads: 50% done sounds fine until you look at the calendar.\nOn 11/20, a made-up 11/02 to 11/30 job has used 64.29% of its time, so it is 14.29 points behind.\n=MIN(1,MAX(0,(asof-start)/(due-start))) keeps that share between 0% and 100%.
---
Find the share of a project's calendar time used with `=MIN(1,MAX(0,(asof-start)/(due-start)))`, then subtract it from your percent done to see if you are ahead or behind. In the made-up tracker on 11/20/2026, a job from 11/02 to 11/30 that is 50% done has used 64.29% of its time, 14.29 points behind, so it shows Behind with a 10-point tolerance; the download also marks Overdue, Done, a one-day job and due dates before the start (Check: due before start).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (41 workbooks, 8,569 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (MIN, MAX, IF, AND, OR, ISNUMBER, ROUND, COUNTIF, SUMPRODUCT) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Eight client design projects comparing percent done with percent of time elapsed](images/2026-11-20-team-05-percent-of-time-elapsed-vs-percent-complete-project-status-excel-google-sheets/percent-time-elapsed-vs-complete-template.png) *Eight fictional client projects as of 11/20/2026 with a 10% tolerance: Logo refresh is 50% done with 64.29% of its time gone, so Behind; Holiday catalog is Overdue; Menu boards is Done; and the Brochure row with its dates typed backwards shows Check: due before start (the summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Projects** tab, one project per row starting in row 5.

1. Type the date you are checking on in **B1**, such as 11/20/2026. It is a typed date, not TODAY, so the sheet only moves when you change it.
2. Type your tolerance in **B2** as a percent, such as 10%: within 10 points of time elapsed counts as On track.
3. In **A5:E5**, type the project, the client, the start date, the due date and your own percent done as a percent (50%, not 50).
4. In **F5**, type `=MIN(1,MAX(0,($B$1-C5)/(D5-C5)))`. It divides the days from start to the as-of date by the days from start to due. MAX(0,...) stops it at 0% before the start and MIN(1,...) at 100% after the due date.
5. In **G5**, type `=E5-F5`. A negative number means the work is behind the calendar; a positive one means it is ahead.
6. In **H5**, type `=MAX(0,D5-$B$1)` for days left, so an overdue job shows 0, not a negative count.
7. In **I5**, turn the gap into a status: Done at 100% done, Overdue when the as-of date is past the due date and the work is not done, Behind when G5 is below minus the tolerance, Ahead when it is above the tolerance, and On track otherwise. A gap of exactly 10 points either way is On track.
8. Fill **F5:I5** down. The template has these formulas in rows 5 to 54, with guards for blank rows, text dates and one-day jobs.

Pick one day-count rule and keep it. The sheet counts the calendar days between the dates and treats the as-of date as the start of that day. Counting both end days, `(asof-start+1)/(due-start+1)`, gave 65.52% instead of 64.29% for the logo job in LibreOffice. Mixing the two rules makes two sheets disagree. This version counts calendar days, not business days.

Microsoft's [MIN function](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152) and [MAX function](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098) pages cover the two functions that do the clamping. To show the dates as bars instead, see the Tidy Tabs post on [making a simple Gantt chart with conditional formatting](https://tidytabs.blogspot.com/2026/10/how-to-make-simple-gantt-chart-with.html).

## Example

The projects and clients are fictional. As-of date 11/20/2026, tolerance 10%.

| Project | Start | Due | Done | Time elapsed | Done minus elapsed | Days left | Status |
|---|---|---|---|---|---|---|---|
| Logo refresh | 11/02/2026 | 11/30/2026 | 50% | 64.29% | -14.29% | 10 | Behind |
| Website rebuild | 11/10/2026 | 12/10/2026 | 20% | 33.33% | -13.33% | 20 | Behind |
| Holiday catalog | 11/01/2026 | 11/15/2026 | 80% | 100.00% | -20.00% | 0 | Overdue |
| Menu boards | 11/01/2026 | 11/15/2026 | 100% | 100.00% | 0.00% | 0 | Done |
| Spring lookbook | 12/01/2026 | 12/20/2026 | 0% | 0.00% | 0.00% | 30 | On track |
| Flyer edit | 11/20/2026 | 11/20/2026 | 0% | 100.00% | -100.00% | 0 | Behind |
| Social posts | 11/18/2026 | 11/22/2026 | 50% | 50.00% | 0.00% | 2 | On track |
| Brochure (dates typed backwards) | 11/30/2026 | 11/10/2026 | 30% | | | | Check: due before start |

The Flyer edit is due today, so its time is fully used but it is not Overdue yet: it shows Behind until the as-of date passes 11/20/2026. The summary in **K4:L12** counts 8 projects listed, 1 Done, 1 Overdue, 3 Behind, 2 On track, 0 Ahead and 1 row showing Check, and CHECK shows Check because of the Brochure row.

## Troubleshooting

### Time elapsed shows more than 100% or a negative percent
The formula is not clamped. In LibreOffice the plain `(asof-start)/(due-start)` gave 135.71% for the overdue Holiday catalog and -57.89% for the Spring lookbook, which has not started. Wrap it in `MIN(1,MAX(0,...))`.

### #DIV/0! on a one-day job
When the start date equals the due date, the plain formula divides by zero, as it did for the Flyer edit. The template checks for that case first and shows 0% before that day and 100% on or after it.

### Every row shows Check: tolerance
**B2** holds 10 instead of 10%, or it is blank. The tolerance must be from 0% to under 100%, so retype it as a percent.

### A row shows Check: percent done
Percent done was typed as 50 instead of 50%, left blank, or is above 100% or below 0%. Retype it as a percent.

### A row shows Check: due before start
The start and due dates are swapped, like the Brochure row (11/30/2026 to 11/10/2026). Retype them in the right columns.

## Template

[Download the .xlsx](templates/tidy-tabs-project-time-elapsed-tracker.xlsx)

The file has two tabs. **Projects** has the As-of date in **B1** and the Tolerance (points) in **B2**, then rows 5 to 54 with Project, Client, Start date, Due date and Percent done (inputs in **A:E**) and Time elapsed %, Done minus elapsed, Days left and Status in **F:I**. Status reads Done, Overdue, Behind, On track, Ahead, or a Check label: as-of date, tolerance, dates, due before start or percent done. The summary in **K4:L12** counts each status, rows showing Check and entries typed below row 54, plus the CHECK. **How to use** explains each formula. The Projects tab prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Behind turns orange, Overdue and every Check turn red, and Done and Ahead turn green. Percent done and the tolerance are yours; the sheet only compares numbers. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
