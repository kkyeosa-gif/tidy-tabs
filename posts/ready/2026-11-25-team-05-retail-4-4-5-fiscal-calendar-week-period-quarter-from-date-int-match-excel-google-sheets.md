---
title: How to Get the 4-4-5 Fiscal Week and Period for Any Date in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 64 workbooks and 647,616 cell comparisons, 0 mismatches, plus the edge cases listed in the post (every day from 01/29/2023 to 01/29/2028 for 4-4-5, 4-5-4 and 5-4-4). Not tested in Excel or Google Sheets.
search_description: Get the 4-4-5 fiscal week with INT((date-start)/7)+1, the period with MATCH on a first-week table and the quarter with INT. Handles 53-week years.
image_prompts: A large paper wall calendar hanging in a small retail stockroom with a few weeks blocked off by colored sticky flags in groups of four, four and five, beside a metal clipboard on a hook and a stack of folded shopping bags on a shelf below, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Paper wall calendar with colored flags grouping weeks in a retail stockroom
threads: =MONTH() says 01/30/2027 is January. A 4-4-5 calendar says FY2026 P12 W52.\nWeek is =INT((date-start)/7)+1, then MATCH it against each period's first week (1, 5, 9, 14... for 4-4-5).\nIn the sample, 11/27/2026 is week 43, period 10, quarter 4.
---
Find the fiscal year start with `=MATCH(date,starts,1)`, get the fiscal week with `=INT((date-start)/7)+1`, then the period with `=MATCH(week,first_weeks,1)` from a table of each period's first week (1, 5, 9, 14... for 4-4-5, with a second column for 53-week years) and the quarter with `=INT((period-1)/3)+1`. With a made-up FY2026 starting Sunday 02/01/2026, 11/27/2026 is week 43, period 10, quarter 4.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (64 workbooks, 647,616 cell comparisons, 0 mismatches), including every day from 01/29/2023 to 01/29/2028 for all three patterns. Excel and Google Sheets were not opened for this post.

![Sixteen sample dates converted to fiscal year, week, period, quarter and week label](images/2026-11-25-team-05-retail-4-4-5-fiscal-calendar-week-period-quarter-from-date-int-match-excel-google-sheets/4-4-5-fiscal-calendar-template.png) *Dates tab with a fictional 4-4-5 calendar starting Sunday 02/01/2026: 11/27/2026 is FY2026 P10 W43 in quarter 4 while its calendar month is 11, and 01/28/2024 falls in week 53 of FY2023. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

Type your own fiscal year start dates; the dates below are examples, not the NRF's or any company's calendar. Where week 53 goes is your calendar's rule, and the sheet does not pick one.

The steps are the same in Excel and Google Sheets.

1. On the **Calendar** tab, pick your pattern in **B1**: 4-4-5, 4-5-4 or 5-4-4.
2. In **B2**, type the period that gets week 53 in a 53-week year, a whole number from 1 to 12. The sample uses 12.
3. From row 6, type one fiscal year per row: a label in **A** (FY2026) and the first day in **B** (02/01/2026). Add one more row with the next year's start, so the last year has an end.
4. **C6** counts weeks with `=IF(OR(NOT(ISNUMBER(B6)),NOT(ISNUMBER(B7))),"",(B7-B6)/7)`. It should say 52 or 53; **E** shows OK or a Check message.
5. The period table in **G6:K17** builds itself from **B1**. **H** reads the weeks per period from the pattern with `=VALUE(MID($B$1,2*MOD(G6-1,3)+1,1))`, **I** adds them into each period's first week, and **J** does the same for a 53-week year, where periods after **B2** start one week later.
6. On the **Dates** tab, type or paste your dates in **A5:A1004**.
7. **B5** finds the fiscal year row: `=MATCH(A5,Calendar!$B$6:$B$25,1)`. MATCH with 1 returns the last start on or before the date. See the [MATCH help from Microsoft](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) or the [MATCH help from Google](https://support.google.com/docs/answer/3093378).
8. **D5** gets the week: `=IF(C5="","",INT((A5-INDEX(Calendar!$B$6:$B$25,B5))/7)+1)`. See the [INT help from Microsoft](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef).
9. **E5** gets the period by matching the week against the right first-week column: `=IF(D5="","",IF(INDEX(Calendar!$C$6:$C$25,B5)=53,MATCH(D5,Calendar!$J$6:$J$17,1),MATCH(D5,Calendar!$I$6:$I$17,1)))`.
10. **F5** gets the quarter: `=IF(E5="","",INT((E5-1)/3)+1)`.
11. **G:K** show the week and period start dates, a label like FY2026 P10 W43, the calendar month and a Status.

Starts on a Saturday or a Monday work too; the sheet counts whole weeks from the start you type.

## Example

The sample data is made up. Starts (all Sundays): FY2023 01/29/2023, FY2024 02/04/2024, FY2025 02/02/2025, FY2026 02/01/2026, FY2027 01/31/2027 and FY2028 01/30/2028 as the end marker. Pattern 4-4-5, week 53 in P12. FY2023 has 53 weeks (last day 02/03/2024); the other years have 52.

| Date | Fiscal year | Week | Period | Quarter | Week starts | Period starts | Calendar month |
|---|---|---|---|---|---|---|---|
| 02/01/2026 | FY2026 | 1 | 1 | 1 | 02/01/2026 | 02/01/2026 | 2 |
| 02/28/2026 | FY2026 | 4 | 1 | 1 | 02/22/2026 | 02/01/2026 | 2 |
| 03/01/2026 | FY2026 | 5 | 2 | 1 | 03/01/2026 | 03/01/2026 | 3 |
| 04/05/2026 | FY2026 | 10 | 3 | 1 | 04/05/2026 | 03/29/2026 | 4 |
| 11/27/2026 | FY2026 | 43 | 10 | 4 | 11/22/2026 | 11/01/2026 | 11 |
| 01/30/2027 | FY2026 | 52 | 12 | 4 | 01/24/2027 | 12/27/2026 | 1 |
| 01/31/2027 | FY2027 | 1 | 1 | 1 | 01/31/2027 | 01/31/2027 | 1 |
| 12/31/2023 | FY2023 | 49 | 12 | 4 | 12/31/2023 | 12/24/2023 | 12 |
| 02/03/2024 | FY2023 | 53 | 12 | 4 | 01/28/2024 | 12/24/2023 | 2 |
| 02/04/2024 | FY2024 | 1 | 1 | 1 | 02/04/2024 | 02/04/2024 | 2 |

In 4-4-5 the first weeks are 1, 5, 9, 14, 18, 22, 27, 31, 35, 40, 44 and 48, so P12 has 6 weeks in FY2023. For 4-5-4 they are 1, 5, 10, 14, 18, 23, 27, 31, 36, 40, 44, 49, and for 5-4-4 they are 1, 6, 10, 14, 19, 23, 27, 32, 36, 40, 45, 49.

Send week 53 to P1 instead (**B2** = 1) and P1 has 5 weeks: 02/26/2023 to 03/01/2023 become FY2023 P1 W5, and P2 starts at week 6.

For a year that starts on the first of a month instead of a Sunday, see [How to Get the Fiscal Year and Quarter From a Date](https://tidytabs.blogspot.com/2026/11/how-to-get-fiscal-year-and-quarter-from.html). For WEEKNUM and ISO weeks, see [How to Get the Week Number From a Date (US and ISO Weeks)](https://tidytabs.blogspot.com/2026/11/how-to-get-week-number-from-date-us-and.html).

## Troubleshooting

### The calendar month does not match the period
That is expected. `=MONTH()` puts 01/30/2027 in month 1, but in the sample it is FY2026 P12 W52. Report by the **Period** column.

### Dates land in the wrong period with 4-5-4 or 5-4-4
A first-week table typed for 4-4-5 (1, 5, 9...) and then used with 4-5-4 (1, 5, 10...) puts dates in the wrong period. Let **I** and **J** build from **B1** instead of typing the first weeks.

### Every date says Check: Calendar tab
Something on **Calendar** is off: starts out of order, a blank row inside the table, a start typed as text or with a time, a missing label, or a year that is not 52 or 53 whole weeks. Column **E** flags the row.

### Recent dates say Check: after the last fiscal year
The last year has no end yet. Type the next year's start date on the next row of **Calendar**.

### A date says Check: date is text
The date was pasted as text, such as "02/01/2026". Retype it as a real date.

## Template

[Download the .xlsx](templates/tidy-tabs-4-4-5-fiscal-calendar.xlsx)

The file has three tabs. **Calendar** holds the pattern in **B1**, the week 53 period in **B2**, up to 20 fiscal year rows in **A6:B25**, the period table in **G6:K17** and a summary in **M4:N9** that ends in a CHECK cell. **Dates** reads up to 1,000 dates in **A5:A1004** and fills week, period, quarter, start dates, label, calendar month and status, with a summary in **M4:N8**. **How to use** explains each formula. The data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. A 53-week year turns light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
