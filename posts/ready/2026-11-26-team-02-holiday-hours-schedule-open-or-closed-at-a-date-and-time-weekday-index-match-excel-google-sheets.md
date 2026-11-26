---
title: How to Make a Holiday Hours Schedule in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 280 workbooks and 408,520 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The MOD time-of-day rounding problem was reproduced in LibreOffice Calc (and in Python's float math) only.
search_description: Holiday hours schedule from weekly hours plus a list of holiday dates with WEEKDAY, INDEX and MATCH, then check if the shop is open at any date and time.
image_prompts: The glass front door of a small gift shop at dusk with a blank paper hours sign taped on the inside of the glass, a string of warm holiday lights around the door frame, and a folded wooden sandwich board sign leaning against the brick wall outside, realistic photo, no screens, no readable text, no logos
image_alt: Gift shop glass door with a blank paper sign and holiday lights at dusk
threads: Closed Thanksgiving, 8:00 AM on Black Friday, regular hours the rest of the week?\nKeep weekly hours plus a short list of dated exceptions. MATCH the date first, and any other day falls back to INDEX(hours,WEEKDAY(date)).\nRound the time to whole seconds before comparing: in LibreOffice, plain MOD put 704 of 1,440 minutes a hair early.
---
List your holiday dates with their own hours, then look each day up with `=IF(COUNTIF(Hours!$F$5:$F$44,A5)=0,0,MATCH(A5,Hours!$F$5:$F$44,0))`: a date on the list uses its row, and any other day uses its weekday's hours with `INDEX(Hours!$B$5:$B$11,WEEKDAY(A5))`. In the made-up shop, 11/23/2026 to 01/03/2027 prints as 42 days with 34 open, Thanksgiving closed and Black Friday 8:00 AM to 8:00 PM.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (280 workbooks, 408,520 cell comparisons, 0 mismatches), plus the edge cases below. The MOD time-of-day rounding problem was reproduced in LibreOffice Calc and in Python's float math only. Excel and Google Sheets were not opened for this post.

![Holiday schedule from 11/23 to 12/18 beside fourteen open or closed time checks](images/2026-11-26-team-02-holiday-hours-schedule-open-or-closed-at-a-date-and-time-weekday-index-match-excel-google-sheets/holiday-hours-schedule-template.png) *Left: Schedule tab of a fictional shop from 11/23/2026 to 12/18/2026, with closed Mondays and Thanksgiving in orange, Black Friday and Small Business Saturday special hours in green, and the summary (42 days, 34 open, 253.00 hours). Right: the Check a time tab with 14 sample date-times, such as 11/27/2026 7:59 AM Not open yet and 8:00 PM Closed for the day. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

The sheet knows no holiday dates or rules. It uses only the dates you type. Hours are same-day only: a shop that closes after midnight is not supported.

The steps are the same in Excel and Google Sheets.

1. On the **Hours** tab, type your regular hours in rows 5 to 11 (Sunday to Saturday): Opens in **B**, Closes in **C**. Leave both blank for a day you are closed. Column **D** shows Open, Closed or a Check message.
2. In **F5:I44**, list your dated exceptions: Date in **F**, Note in **G** (Thanksgiving, Black Friday), Opens in **H**, Closes in **I**. Leave both times blank for a date you are closed. Column **J** checks each row, and **M7** is the tab's CHECK cell.
3. On the **Schedule** tab, type the first date in **B1** and the number of days to list in **B2** (1 to 62). **A5** lists the dates with `=IF($K$4<>"OK","",IF(ROWS($A$5:A5)>$B$2,"",$B$1+ROWS($A$5:A5)-1))`, and **B5** names the day with `=IF(A5="","",TEXT(A5,"dddd"))`.
4. **H5** finds the date on your exception list, or 0 when it is not there: `=IF(A5="","",IF(COUNTIF(Hours!$F$5:$F$44,A5)=0,0,MATCH(A5,Hours!$F$5:$F$44,0)))`. The COUNTIF test keeps MATCH from returning #N/A. See the MATCH help from [Microsoft](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) or [Google](https://support.google.com/docs/answer/3093378).
5. **D5** picks the opening time: `=IF(A5="","",IF(H5>0,IF(ISNUMBER(INDEX(Hours!$H$5:$H$44,H5)),INDEX(Hours!$H$5:$H$44,H5),""),IF(ISNUMBER(INDEX(Hours!$B$5:$B$11,WEEKDAY(A5))),INDEX(Hours!$B$5:$B$11,WEEKDAY(A5)),"")))`. `WEEKDAY` gives 1 for Sunday through 7 for Saturday, which is the row order of the weekly table. **E5** is the same with **Hours!$I** and **Hours!$C**. See the [WEEKDAY help from Microsoft](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a) and the [INDEX help from Google](https://support.google.com/docs/answer/3098242).
6. **C5** pulls the note, **F5** gives hours open with `=IF(OR(D5="",E5=""),"",ROUND((E5-D5)*24,2))`, and **G5** shows Regular hours, Special hours, Closed or Closed (holiday).
7. Read the summary in **J4:K10**: days listed, open days (`=COUNTIF(G5:G66,"*hours")`), closed days (`=COUNTIF(G5:G66,"Closed*")`), days from your date list, total hours and CHECK.
8. To answer "are we open?", go to **Check a time** and type a date and time in **A5**, such as 11/27/2026 7:59 AM. **B5** takes the date with `=IF(H5<>"OK","",INT(A5))` and **C5** the time with `=IF(H5<>"OK","",ROUND(MOD(A5,1)*86400,0)/86400)`, rounded to whole seconds. See the MOD help from [Microsoft](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) and the [ROUND help from Google](https://support.google.com/docs/answer/3093440).
9. **J5** turns the time into the minute of the day with `=IF(B5="","",INT(ROUND(C5*86400,0)/60))`, and **F5** compares minutes: `=IF(B5="","",IF(D5="",IF(I5>0,"Closed (holiday)","Closed today"),IF(J5<ROUND(D5*1440,0),"Not open yet",IF(J5<ROUND(E5*1440,0),"Open","Closed for the day"))))`. Open means opening time <= time < closing time, so at 8:00 PM sharp a shop that closes at 8:00 PM shows closed.
10. **G5** shows the minutes until it opens (Not open yet) or closes (Open).

The sheet does not update your website or Google Business Profile. Copy the hours over by hand.

For cutoff times on orders, see [How to Calculate a Ship Date With a Daily Cutoff Time](https://tidytabs.blogspot.com/2026/11/how-to-calculate-ship-date-with-daily.html). For time math with minutes, see [How to Add Minutes to a Start Time to Find an Appointment End Time](https://tidytabs.blogspot.com/2026/10/how-to-add-minutes-to-start-time-to.html).

## Example

The shop and its hours are made up. Weekly hours: Sunday 11:00 AM to 4:00 PM, Monday closed, Tuesday to Friday 10:00 AM to 6:00 PM, Saturday 9:00 AM to 5:00 PM. The date list holds eight holiday dates. **B1** is 11/23/2026 and **B2** is 42.

| Date | Day | Note | Opens | Closes | Hours open | Status |
|---|---|---|---|---|---|---|
| 11/26/2026 | Thursday | Thanksgiving | | | | Closed (holiday) |
| 11/27/2026 | Friday | Black Friday | 8:00 AM | 8:00 PM | 12.00 | Special hours |
| 11/28/2026 | Saturday | Small Business Saturday | 9:00 AM | 7:00 PM | 10.00 | Special hours |
| 11/29/2026 | Sunday | | 11:00 AM | 4:00 PM | 5.00 | Regular hours |
| 11/30/2026 | Monday | | | | | Closed |
| 12/21/2026 | Monday | Holiday week | 10:00 AM | 6:00 PM | 8.00 | Special hours |
| 12/24/2026 | Thursday | Christmas Eve | 10:00 AM | 2:00 PM | 4.00 | Special hours |
| 12/25/2026 | Friday | Christmas Day | | | | Closed (holiday) |
| 12/31/2026 | Thursday | New Year's Eve | 10:00 AM | 3:00 PM | 5.00 | Special hours |
| 01/01/2027 | Friday | New Year's Day | | | | Closed (holiday) |

The summary shows 42 days listed, 34 open, 8 closed (five Mondays plus three holidays), 8 days from the date list, 253.00 hours open and CHECK OK.

On **Check a time**, 14 sample date-times give these results (6 open):

| Date and time | Open? | Minutes |
|---|---|---|
| 11/26/2026 12:00 PM | Closed (holiday) | |
| 11/27/2026 7:59 AM | Not open yet | 1 |
| 11/27/2026 8:00 AM | Open | 720 |
| 11/27/2026 7:59 PM | Open | 1 |
| 11/27/2026 8:00 PM | Closed for the day | |
| 11/30/2026 11:00 AM | Closed today | |
| 12/21/2026 5:30 PM | Open | 30 |
| 12/24/2026 2:00 PM | Closed for the day | |
| 12/26/2026 8:45 AM | Not open yet | 15 |

Seconds are rounded first and then the minute is cut down, so 7:59:30 AM still counts as minute 479 and shows Not open yet, and 8:00:59 PM shows Closed for the day.

## Troubleshooting

### A time shows one minute early or "Not open yet" right at opening
The time was taken from the date and time with plain `=MOD(A5,1)`. In LibreOffice Calc, 704 of the 1,440 minutes of 11/27/2026 came out a hair under the minute (10:00 AM gave minute 599), and Python's float math gives the same 704. A first draft of the template with plain MOD showed 12/21/2026 5:30 PM as 5:29 PM in a LibreOffice PDF render. This was seen in LibreOffice and Python only, not checked in Excel or Google Sheets. Round to whole seconds first, as **C5** does.

### Every Schedule row says Check: Hours tab
The Hours tab has a problem: a time pair with one time blank, a time typed as text, a date listed twice, or a date typed below row 44. Find the Check message in column **D** or **J** of the Hours tab, or **M6** for rows below the table, and fix that row.

### A holiday shows regular hours
The date on the Hours tab has a time attached or is text, so MATCH cannot find it. Column **J** says "Check: date has a time" or "Check: date blank or text". Retype it as a plain date.

### A day that closes after midnight shows Check
Closing at or before opening shows "Check: closes before it opens". The sheet handles same-day hours only, so enter 11:59 PM as the closing time if you need the day to run late.

### Check a time says Check: no date, time only
You typed a time like 2:00 PM without a date. Type the full date and time, such as 12/24/2026 2:00 PM.

## Template

[Download the .xlsx](templates/tidy-tabs-holiday-hours-schedule.xlsx)

**Hours** holds the weekly hours and up to 40 dated exceptions (rows 5 to 44) with a CHECK cell that also catches dates typed below row 44. **Schedule** lists up to 62 days with the summary. **Check a time** takes 100 date-times (rows 5 to 104) and flags rows typed below row 104. **How to use** explains each formula. The data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Closed days are light orange, special hours and Open are light green, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
