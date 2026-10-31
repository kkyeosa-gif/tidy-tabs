---
title: How to Track Sales Leads and Your Win Rate in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model of stage counts, quote values, win rate, weighted pipeline and best source. Edge cases run: no Won or Lost leads, a stage not on the list, a blank stage, a trailing space in "Won ", a renamed stage, a source not on the list, and no Lost leads. The Stage and Source dropdowns are data validation lists; their menus were not opened, only the formulas were checked. Excel and Google Sheets were not used.
search_description: Track sales leads and calculate win rate with COUNTIF in Excel and Google Sheets. Free lead tracker template with weighted pipeline, best source and a check cell.
image_prompts: A clipboard with a hand-sketched four-column board of sticky notes beside a cardboard box of business cards and a half-full mug on a freelancer's wooden desk, soft window light, no screens, no readable text, no logos
image_alt: Sticky notes sorted into columns on a clipboard next to a box of business cards
threads: Win rate is won divided by won plus lost, not by every lead.\n=IF(B8+B9=0,"",B8/(B8+B9)) leaves open leads out.\nIt also stays blank before your first close, so no divide-by-zero.
---
Divide won leads by won plus lost leads: `=IF(B8+B9=0,"",B8/(B8+B9))`, where B8 and B9 hold the counts from `=COUNTIF(Leads!$C$2:$C$201,A8)`. Open leads are left out, and the blank result avoids a divide-by-zero before your first close.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIF, SUMIF, SUMPRODUCT and INDEX/MATCH exist in Excel and Google Sheets, but neither app was opened for this post. The dropdown menus in the file were not exercised.

![Lead tracker summary with stage counts, weighted pipeline value, win rate and wins by source](images/2026-11-01-team-04-sales-lead-pipeline-conversion-rate-excel-google-sheets/sales-lead-tracker-win-rate-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The layout matches the template. **Leads** has one row per lead from row 2: lead in A, source in B, stage in C, quote in D and date in E. **Lists** holds the stages in A2:A6 with a chance to win in B2:B6, and your sources in D2:D6. **Summary** does the math.

Keep **Won** and **Lost** as the last two stages on **Lists**. The win rate reads them from rows 8 and 9 of **Summary**.

The steps are the same in Excel and Google Sheets, except for the dropdown in step 2.

1. On **Lists**, type your stages in A2:A6 in order: New, Contacted, Quoted, Won, Lost. Type a chance to win next to each, such as 50% for Quoted.
2. On **Leads**, add a dropdown to **C2:C201** that reads its choices from **Lists!A2:A6**. Do the same for **B2:B201** with **Lists!D2:D6**.
3. On **Summary**, type `=Lists!A2` in **A5** and fill it down to **A9**.
4. In **B5**, type `=COUNTIF(Leads!$C$2:$C$201,A5)`.
5. In **C5**, type `=SUMIF(Leads!$C$2:$C$201,A5,Leads!$D$2:$D$201)`.
6. Select **B5:C5** and drag the fill handle down to row 9.
7. In **B13**, type the win rate formula: `=IF(B8+B9=0,"",B8/(B8+B9))`. Format it as a percent.
8. In **B16**, type `=SUMPRODUCT(C5:C7,D5:D7)`, where D holds each stage's chance (`=Lists!B2` and so on).

### In Excel

For step 2, select the cells, then choose **Data > Data Validation**, set Allow to List and point Source at the range. Microsoft's page: [Create a drop-down list](https://support.microsoft.com/en-us/office/create-a-drop-down-list-7693307a-59ef-400a-b769-c5402dce407b).

### In Google Sheets

For step 2, select the cells, choose **Insert > Dropdown** and pick "Dropdown (from a range)". Google's page: [Create an in-cell dropdown list](https://support.google.com/docs/answer/186103).

## Why closed leads only

A lead you are still chasing has not won or lost yet. Counting it in the bottom of the fraction makes your win rate look worse than it is.

The weighted pipeline in step 8 covers the three open stages only. Won is money you have, not pipeline. The chances are your own guesses, not data, so treat the result as a rough number.

For the best source, **Summary B17** uses `=IF(MAX(C22:C26)=0,"n/a",INDEX(A22:A26,MATCH(MAX(C22:C26),C22:C26,0)))`. It returns the source with the most wins. A tie returns the first one in the list.

Help pages: Microsoft's [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) and Google's [COUNTIF page](https://support.google.com/docs/answer/3093480). They describe the function, not our test.

## Example

Sample data is fictional: 25 made up leads from 08/03/2026 on.

| Stage | Leads | Quote value | Chance |
|---|---|---|---|
| New | 3 | $7,300.00 | 10% |
| Contacted | 4 | $8,050.00 | 25% |
| Quoted | 6 | $9,500.00 | 50% |
| Won | 7 | $9,950.00 | 100% |
| Lost | 5 | $9,350.00 | 0% |
| All leads | 25 | $44,150.00 | |

Win rate is 7 / (7 + 5), or 58.3%. There are 13 open leads worth $24,850.00. The weighted pipeline is $730.00 + $2,012.50 + $4,750.00, or $7,492.50. The check cell says OK.

By source, Google search has 3 wins and 0 losses, so it is the best source. Website form has 2 wins and 1 loss (66.7%). Referral and Instagram each have 0 wins and 2 losses. Five closed leads per source is a thin sample, so read this as a hint.

## Troubleshooting

### Win rate is blank
No lead is Won or Lost yet, so B8 + B9 is 0 and the formula returns an empty cell on purpose. In the tests with no closed leads, best source showed "n/a" and the check cell stayed OK. Mark a lead Won or Lost and the percent appears.

### A lead is missing from every stage row
Its stage text does not match the list. In the tests, "Won " with a trailing space was not counted as Won, and a made-up stage ("Negotiating") landed in the "Blank or not in the stage list" row. The check cell showed CHECK. Retype the stage exactly or pick it from the list.

### Win rate shows 100%
No lead is Lost. In the tests, with no Lost leads the formula returned 100%. That is arithmetic, not a trend.

### Stage counts dropped after renaming a stage
You renamed a stage on **Lists** but not on the leads. Those leads moved to the "not in the stage list" row, and the check cell showed CHECK. Update the stage on each lead.

### The check cell says CHECK
It shows CHECK for a blank stage or one not on the list, a source not on the list, a quote stored as text, or missing chances. The "Blank or not in the stage list" row on **Summary** shows how many leads need fixing.

## Template

[Download the .xlsx](templates/tidy-tabs-lead-tracker.xlsx)

The file has four tabs. **Leads** has 25 fictional leads in **A2:E201** (200 slots), with Source and Stage dropdowns set up as data validation lists. **Lists** has the five stages with chances and five sources. **Summary** has the stage table, win rate, open leads, open quote value, weighted pipeline, best source, a source table and a check cell. **How to use** has short notes.

Limits: more than 200 leads, or removing and adding stages beyond five, were not tested. The dropdowns were not opened, and the file does not show an error message if you type a value that is not on the list, so the check cell is your safety net. The file was recalculated only in LibreOffice Calc 24.2.7.2. All tabs are set to portrait on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
