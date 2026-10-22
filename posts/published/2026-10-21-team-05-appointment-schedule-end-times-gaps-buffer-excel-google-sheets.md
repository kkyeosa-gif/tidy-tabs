---
threads_url: https://www.threads.com/@tin_ylab/post/DezQp13Ecd_
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-add-minutes-to-start-time-to.html
title: How to Add Minutes to a Start Time to Find an Appointment End Time in Excel and Google Sheets
labels: excel-formulas, small-business, scheduling
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an exact minute model, with 33 edited variants (back to back, overlap, past midnight, blank row, text start time, text minutes, buffer 0, 20, blank and text, 11 appointments, one appointment, empty schedule, other slot lengths). No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Add minutes to a start time with =C2+D2/1440 in Excel and Google Sheets. Free appointment schedule template with end times, gaps and buffer warnings.
image_prompts: A grooming table with a folded towel and a pet brush beside a coiled leash in a small shop, soft daylight, no screens, no readable text, no logos
image_alt: Grooming table with a folded towel, pet brush and coiled leash in a small shop
threads: Zero minutes between Mochi and Cooper.\n=C2+D2/1440 gives the end time and =ROUND((C3-E2)*1440,0) gives the gap. With a 15 minute buffer, 3 of 8 appointments get flagged.
---
Add the appointment length to the start time with `=C2+D2/1440`, where C2 is the start time and D2 the minutes. Divide the minutes by 1440 because a day has 1,440 minutes, and format the result as a time.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with an exact minute model. The Excel and Google Sheets steps use the same formulas as typed. Neither Excel nor Google Sheets was opened for this post.

![Appointment schedule with start times, minutes, end times, gaps and status for eight appointments](images/2026-10-21-team-05-appointment-schedule-end-times-gaps-buffer-excel-google-sheets/appointment-schedule-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

Excel and Sheets store a time as a fraction of a day. 9:00 AM is 0.375, and one minute is 1/1440. That is why minutes must be divided by 1440 before they are added.

1. In **A1:D1**, type the headings Client, Service, Start time and Minutes.
2. In **A2:D2**, type the first appointment. Type the start time as `9:00 AM` and the length as a plain number such as 90.
3. In **E2**, type `=C2+D2/1440`. Format **E2** as a time such as 1:30 PM.
4. In **F2**, type `=ROUND((C3-E2)*1440,0)`. This is the gap in minutes between this appointment's end and the next start.
5. In **L1**, type your buffer in minutes, for example 15.
6. In **G2**, type `=IF(F2<0,"Overlap",IF(F2<$L$1,"Short buffer","OK"))`.
7. Copy **E2:G2** down the list. Leave the last row's gap blank.

The ROUND in step 4 matters. Subtracting two times gives a fraction of a day, and a fraction can land a hair off a whole number. Rounding to whole minutes keeps a 15 minute gap from being judged shorter than a 15 minute buffer.

The template wraps these in blank and text tests, so an empty or mistyped row shows Check instead of a wrong time. The status formula in the file is longer than step 6 and adds Last, Check next row, Out of order and Past midnight.

### In Excel

The formulas work as typed. Microsoft describes the rounding on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) page and the time building block on its [TIME function](https://support.microsoft.com/en-us/office/time-function-9a5aff99-8f7d-4611-845e-747d0b8d5457) page. Those are the help pages, not our test.

### In Google Sheets

The same formulas should work as typed. Select the cells and choose **Format > Number > Time**. Google explains number formats on its [Format numbers in a spreadsheet page](https://support.google.com/docs/answer/56470). Sheets was not opened for this post.

## Example

Sample data is fictional: a dog groomer's day with a 15 minute buffer in **L1**.

| Client | Start | Minutes | End | Gap | Status |
|---|---|---|---|---|---|
| Bailey | 9:00 AM | 90 | 10:30 AM | 15 | OK |
| Mochi | 10:45 AM | 45 | 11:30 AM | 0 | Short buffer |
| Cooper | 11:30 AM | 90 | 1:00 PM | 15 | OK |
| Luna | 1:15 PM | 30 | 1:45 PM | 15 | OK |
| Rosie | 2:00 PM | 60 | 3:00 PM | 10 | Short buffer |
| Max | 3:10 PM | 90 | 4:40 PM | 5 | Short buffer |
| Daisy | 4:45 PM | 30 | 5:15 PM | 15 | OK |
| Buddy | 5:30 PM | 45 | 6:15 PM | blank | Last |

Three gaps fall below the buffer: Mochi to Cooper (0), Rosie to Max (10) and Max to Daisy (5). A gap of 0 means back to back, and it still shows Short buffer.

The day totals sit in **L2:L8**:

| Cell | Result |
|---|---|
| L2 Booked minutes | 480 |
| L3 First start | 9:00 AM |
| L4 Last end | 6:15 PM |
| L5 Open minutes between | 75 |
| L6 Appointments | 8 |
| L7 Rows to look at | 3 |
| L8 Check | OK |

The span from 9:00 AM to 6:15 PM is 555 minutes. Minus 480 booked leaves 75, which is also the sum of the gaps. **L8** shows OK when booked minutes plus gaps equals that span.

The Slots tab lists 30 minute slots from the opening time in **B1** to the closing time in **B3**. For this day it shows 19 slots, 18 Taken and 1 Shared slot (4:30 PM, where one appointment ends at 4:40 PM and the next starts at 4:45 PM). A slot with no appointment shows Open.

## Troubleshooting

### A late appointment ends at 12:30 AM
An appointment from 11:30 PM for 60 minutes passes midnight. The sheet shows the end as 12:30 AM and the status as Past midnight, and **L8** says Fix rows first. The template plans one day only, so move the start earlier or shorten the length.

### The gap is negative and the status says Overlap
The next client starts before this one ends. Mochi moved to 10:25 AM, 5 minutes before Bailey ends, gives a gap of -5. Change the start or the minutes until the gap is 0 or more.

### Back to back appointments show Short buffer
A gap of 0 is below any buffer above 0. If you want back to back booking with no warning, set the buffer in **L1** to 0, and every row shows OK.

### A row shows Check
A start time or minutes value is text, such as "noon" or "ninety", or the minutes are 0, negative or a decimal such as 37.5. The row above shows Check next row. Retype the time as `12:00 PM` and the minutes as a whole number. A buffer in **L1** that is blank or text makes every row show Check.

### Last appears in the middle of the day
A blank row sits in the list. The row before it reads as the last appointment, and **L8** shows CHECK because booked minutes plus gaps no longer equals the span. Fill or delete the blank row.

## Template

[Download the .xlsx](templates/tidy-tabs-appointment-schedule.xlsx)

The file has three tabs. **Schedule** holds the yellow inputs in **A2:D12** and **L1** (11 rows), the formulas in **E2:I12**, the day totals in **K1:L8** and the check in **L8**. **Slots** holds the opening time, slot length and closing time in **B1:B3** and up to 48 slots. **How to use** has short notes. Schedule prints landscape and the other tabs portrait, all on US Letter. Short buffer rows are amber and Overlap or Check rows are red.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets, and the colors were not viewed in a spreadsheet window. It works with times and minutes only, not booking rules, deposits or pay. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
