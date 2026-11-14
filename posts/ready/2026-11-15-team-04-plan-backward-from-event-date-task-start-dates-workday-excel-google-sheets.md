---
title: How to Plan Backward From an Event Date With WORKDAY in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 34 workbooks and 3,298 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Work back from a fixed event date with negative WORKDAY to get each task's start and finish, skipping weekends and your own days off, plus a Start by date.
image_prompts: A sheet tray of freshly poured candles in plain glass jars, each wick held upright by a wooden centering bar, on a home studio counter beside a small heat gun, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Freshly poured candles with wooden wick bars on a sheet tray beside a heat gun
threads: Work back from the event with a negative WORKDAY: =WORKDAY(finish,-(days-1),holidays) skips weekends and your own days off.\nCraft fair on Saturday 12/05/2026, five tasks, two shop days off: start by 11/16/2026.\nUse -1 for the last task, not 0. In LibreOffice a 0 step just returned the Saturday.
---
Set the last task's finish with `=WORKDAY(event,-1,holidays)`, its start with `=WORKDAY(finish,-(days-1),holidays)`, and each earlier task's finish with `=WORKDAY(next start,-1,holidays)`, so weekends and your own days off are skipped. In the made-up sample, a craft fair on Saturday 12/05/2026 with five tasks and shop days off on 11/26 and 11/27 gives a Start by date of 11/16/2026, and the download adds a Check for blank, zero or text day counts.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (34 workbooks, 3,298 cell comparisons, 0 mismatches), plus the edge cases below. WORKDAY also exists in Excel and Google Sheets with the same arguments, but neither app was opened; what this post says about them comes from Microsoft's and Google's help pages.

![Five candle-making tasks with work days, start and finish dates counted back from December 5](images/2026-11-15-team-04-plan-backward-from-event-date-task-start-dates-workday-excel-google-sheets/plan-backward-from-event-date-template.png) *Five tasks planned back from a Saturday 12/05/2026 event skip weekends and the shop's 11/26 and 11/27 days off, so the plan starts 11/16/2026. PDF export from LibreOffice Calc 24.2.*

## Steps

This is how you calculate a start date from an end date and a number of work days. A negative day count makes WORKDAY go back in time, skipping Saturdays, Sundays and any dates in your days-off list. The formulas are typed the same way in Excel and Google Sheets.

1. On a tab named **Holidays**, type your own days off in **A2:A20** as real dates (not text). These are your shop's closed days, not an official holiday list.
2. On the **Plan** tab, type the event date in **B1** and a buffer in **B2**: the number of work days you want free before the event. Use 0 for none.
3. In **A5** down, list the tasks in the order you will do them. In **B5** down, type the work days each one needs, as whole numbers of 1 or more. Leave no blank row between tasks.
4. In the last task's **Finish** cell, type `=WORKDAY($B$1,-1-$B$2,Holidays!$A$2:$A$20)`. With a buffer of 0, that is the work day before the event.
5. In each **Start** cell (column C), type `=WORKDAY(D5,-(B5-1),Holidays!$A$2:$A$20)`. The -1 inside counts the finish day itself: a 1-day task starts and finishes on the same day.
6. In every other **Finish** cell (column D), type `=WORKDAY(C6,-1,Holidays!$A$2:$A$20)`, the work day before the next task starts.
7. Read the first task's Start. That is your Start by date.

In the template, D5 picks between steps 4 and 6 with `IF(A6="",...)`, so one formula copies down the whole column. Every WORKDAY is also wrapped in `IFERROR(...,"Check")`.

Why -1 for the last task and not 0: in LibreOffice, `=WORKDAY(12/05/2026,0)` returned 12/05/2026 itself, a Saturday. A zero step does not move the date, even off a weekend.

Help pages: Microsoft's [WORKDAY function](https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33) and Google's [WORKDAY](https://support.google.com/docs/answer/3093059). Both say a negative day count goes back in time and a day count with a decimal is truncated.

WORKDAY counts work days only. A step that runs over weekends, such as curing time, is calendar time; the template has no calendar-day formula, so count those days as work days in B or plan that step yourself. This is a date calculator, not project-management advice.

## Example

The plan, the shop and the days off are fictional. The event is Saturday 12/05/2026, the buffer is 0, and the Holidays tab has 11/26/2026, 11/27/2026 and 12/25/2026 marked Shop closed.

| Task | Work days needed | Start | Finish | Status |
|---|---|---|---|---|
| Order supplies | 2 | 11/16/2026 | 11/17/2026 | OK |
| Pour candles | 5 | 11/18/2026 | 11/24/2026 | OK |
| Cure and label | 3 | 11/25/2026 | 12/01/2026 | OK |
| Photograph and price | 2 | 12/02/2026 | 12/03/2026 | OK |
| Pack the booth | 1 | 12/04/2026 | 12/04/2026 | OK |

Cure and label runs Wednesday 11/25, Monday 11/30 and Tuesday 12/01. Thanksgiving Thursday, the Friday day off and the weekend are skipped.

The summary shows Start by 11/16/2026, 13 work days in the plan, 5 tasks, 0 rows showing Check, a last finish of 12/04/2026, 0 days off typed as text, and CHECK OK.

Change the event or buffer and the whole chain moves:

- Event on Wednesday 12/09/2026: Pack the booth on 12/08, Start by 11/18/2026.
- Event on Sunday 12/06/2026: the same dates as Saturday, Start by 11/16/2026.
- Buffer of 3 with the 12/05/2026 event: Pack the booth on 12/01, Start by 11/11/2026.
- Event on 12/25/2026, which is also a day off: Pack the booth on 12/24, Start by 12/08/2026.

## Troubleshooting

### A task starts after it finishes
Its work days are 0. With Pour candles set to 0, its start is 11/25/2026 and its finish 11/24/2026. The template shows Check, and the Start by date moves to 11/23/2026. Use 1 or more.

### Two tasks overlap near the event
There is a blank row in the task list, so the task above it is treated as the last task. With a blank row after Pour candles, Pour candles runs 11/30 to 12/04 while the tasks below the blank still chain to the event. The template shows "Check: blank row". Delete the blank row.

### Every row says Check
A day off is typed as a word, such as "closed". In LibreOffice the plain WORKDAY formula returned #VALUE! for that; the template shows Check instead. Replace it with a date or clear the cell.

### The dates look right but CHECK says Check
A day off is typed as text that looks like a date, such as "11/27/2026". LibreOffice still used it as a day off, so the dates matched the sample, but the summary counts 1 day off typed as text. Retype it as a real date; this was not checked in Excel or Google Sheets.

### A 2.5-day task got 2 days
WORKDAY drops the decimal. In LibreOffice, `=WORKDAY(11/24/2026,-1.5)` returned 11/23/2026, the same as -1. The template flags any day count that is not a whole number with Check.

## Template

[Download the .xlsx](templates/tidy-tabs-plan-backward-from-event.xlsx)

The file has three tabs. **Plan** has the event date in **B1** and the buffer in **B2**, tasks in rows 5 to 34 with Task and Work days needed in **A:B**, then Start, Finish and Status in **C:E**. The summary in **G4:H10** shows Start by, Work days in the plan, Tasks listed, Rows showing Check, Last finish, Day off typed as text and CHECK. **Holidays** holds your days off in **A2:A20** with a note column. **How to use** explains each formula.

Yellow cells are inputs and gray cells are formulas. Status and CHECK turn red for Check. The Plan and Holidays tabs print Landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
