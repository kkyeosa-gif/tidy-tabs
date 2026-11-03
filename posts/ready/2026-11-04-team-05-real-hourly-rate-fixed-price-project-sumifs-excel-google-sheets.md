---
title: How to Find Your Real Hourly Rate on a Fixed-Price Project in Excel and Google Sheets
labels: excel-formulas, freelancers, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 22 workbooks and 2,398 comparisons with 0 mismatches. Checked: an empty time log, no projects, target rates of $0.00, -$5.00, text, $45.00 and $120.50, a rate exactly at the target, a rate of 59.996 that rounds to the target, half-cent rates, a text price, a blank price, a misspelled project in the log, a trailing space, lower case, text hours, odd decimal hours, and five random sets of up to 20 projects and 400 time entries. Excel and Google Sheets were not used.
search_description: Find your real hourly rate on fixed-price projects in Excel and Google Sheets. Free template totals hours with SUMIFS and flags low-rate projects.
image_prompts: A mechanical kitchen timer beside an open sketchbook with blank pencil roughs and a jar of colored pencils on a freelancer's desk by a window, soft daylight, no screens, no readable text, no logos
image_alt: Kitchen timer next to an open sketchbook and a jar of colored pencils
threads: You quoted $2,400.00 flat. What did you actually earn per hour?\nDivide the price by logged hours: =ROUND(B6/C6,2), with hours from SUMIFS on a time log. In the sample, 52 hours made it $46.15 an hour, below a $60.00 target.\nThe sheet also shows hours the price covers at your target.
---
Divide the fixed price by the hours you logged with `=ROUND(B6/C6,2)`, where `C6` is `=ROUND(SUMIFS(Time!$C$5:$C$404,Time!$B$5:$B$404,A6),2)`. A $2,400.00 project that took 52 hours paid $46.15 an hour.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. SUMIFS, ROUND, COUNTIF and IF exist in current Excel and Google Sheets, but neither app was opened for this post.

![Project table with fixed price, logged hours, real hourly rate, status and a summary](images/2026-11-04-team-05-real-hourly-rate-fixed-price-project-sumifs-excel-google-sheets/real-hourly-rate-fixed-price-projects-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Time** is a log with the date in **A**, the project in **B** and the hours as a decimal in **C** (1.5 for an hour and a half), from row 5. **Projects** lists each project once from **A6** with its fixed price in **B**.

1. On **Time**, log each work session. Spell the project the same way as on Projects.
2. On **Projects**, type your target hourly rate in **B2**, such as $60.00.
3. In **C6**, type `=ROUND(SUMIFS(Time!$C$5:$C$404,Time!$B$5:$B$404,A6),2)` for the hours logged.
4. In **D6**, type `=IF(C6>0,ROUND(B6/C6,2),"")` for the real hourly rate.
5. In **E6**, type `=ROUND(B6/$B$2,2)`. This is the hours the price covers at your target rate.
6. In **F6**, type `=ROUND(E6-C6,2)`. A negative number means you worked more hours than the price covers.
7. In **G6**, type `=IF(C6<=0,"No hours",IF(D6>=$B$2,"At or above target","Below target"))`.
8. Fill **C6:G6** down.

For all projects together, divide the total of the prices by the total hours: `=ROUND(SUM(B6:B25)/SUM(C6:C25),2)`.

This divides the price by hours and nothing else. It does not subtract software, materials, fees or taxes, and it is not tax or pricing advice.

Help pages: Microsoft's [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/03/2026; the pages were not read for this post.

## Example

Sample data is fictional: six design projects, 30 time entries from 09/08/2026, a target of $60.00 an hour.

| Project | Fixed price | Hours logged | Real hourly rate | Status |
|---|---|---|---|---|
| Bakery logo and labels | $900.00 | 13.50 | $66.67 | At or above target |
| Studio website refresh | $2,400.00 | 52.00 | $46.15 | Below target |
| Product photos, 40 items | $640.00 | 14.25 | $44.91 | Below target |
| Menu redesign | $780.00 | 11.50 | $67.83 | At or above target |
| Packaging mockups | $1,200.00 | 17.25 | $69.57 | At or above target |

All six projects together: $6,370.00 for 115.25 hours, or $55.27 an hour. At $60.00 an hour, the $2,400.00 website price covers 40.00 hours. The 52 hours logged are 12.00 hours over.

## Troubleshooting

### A project shows No hours
Nothing is logged under that name, so there is no rate to show. In the test, a project with a price and no time entries showed No hours and a blank rate, not an error.

### Hours are missing from a project
The name on Time does not match the name on Projects. **J11** adds up the hours logged under names that are not on the list. In the test, a misspelled name put 7.25 hours there and the **CHECK** cell said Check. A trailing space also made a different project; upper and lower case did not matter.

### The rate shows exactly $60.00 but the status says At or above
The status compares the rate rounded to cents. In the test, $599.96 over 10 hours (59.996) showed $60.00 and counted as At or above target, while $599.90 over 10 hours showed $59.99 and counted as Below.

### Hours typed as text are ignored
An entry such as `1.5 hrs` is text, so SUMIFS skips it. In the test, one text entry made the **CHECK** cell say Check. Type 1.5.

### B3 shows a message
The target rate must be a number above 0. In the test, 0, -5 and text each gave a message in **B3** and Check settings in the statuses, with no #DIV/0!.

## Template

[Download the .xlsx](templates/tidy-tabs-real-hourly-rate-fixed-price.xlsx)

The file has three tabs. **Projects** has the target in **B2**, a settings check in **B3**, 20 project rows from row 6 (yellow names and prices in **A:B**, formulas in **C:G**) and a summary in **J5:J12**: projects listed, total price, total hours, overall rate, projects below target, the lowest rate, hours on projects not listed and a **CHECK** cell. **Time** has 400 rows from row 5. **How to use** has notes on each formula.

**CHECK** says OK when every logged hour belongs to a listed project, every project has a price and hours, and the hours are numbers.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
