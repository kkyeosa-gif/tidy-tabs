---
threads_url: https://www.threads.com/@tin_ylab/post/DfsjWRwktT1
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-convert-unix-timestamp-to-date.html
title: How to Convert a Unix Timestamp to a Date and Time in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 21 workbooks and 21,147 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Turn a 10-digit or 13-digit Unix timestamp into a date and time with =B4/86400+DATE(1970,1,1), plus a time zone offset. Free Excel template.
image_prompts: A plain analog wall clock above a home office shelf holding a short stack of folded payout statements and a flip desk calendar, all print blurred and unreadable.
image_alt: Wall clock over a shelf with folded payout statements and a desk calendar
threads: A column of numbers like 1793630100 where the payment date should be?\nThat is seconds since 01/01/1970 UTC. =B4/86400+DATE(1970,1,1) makes it a date, and =(B4-5*3600)/86400+DATE(1970,1,1) gives Eastern standard time.\n13 digits means milliseconds: divide by 86400000, or the date lands far past the year 9999.
---
Convert a Unix timestamp in B4 to a date and time with `=B4/86400+DATE(1970,1,1)` and format the cell as mm/dd/yyyy h:mm AM/PM; that gives UTC, and `=(B4-5*3600)/86400+DATE(1970,1,1)` gives US Eastern standard time. For a 13-digit timestamp in milliseconds divide by 86400000 instead; in the made-up sample, 1793630100 becomes 11/02/2026 9:35 AM Eastern.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with a Python model (21 workbooks, 21,147 comparisons, 0 mismatches). Excel and Google Sheets have the same functions, but neither app was opened for this post.

![Ten and thirteen digit Unix timestamps turned into Eastern dates, times and status flags](images/2026-11-13-team-03-convert-unix-timestamp-to-date-and-time-seconds-milliseconds-excel-google-sheets/unix-timestamp-to-date-template.png) *Twelve payment timestamps in seconds and milliseconds converted at -5 hours from UTC; one is cut off and one is text. PDF export from LibreOffice Calc 24.2.*

## Steps

A Unix timestamp counts seconds since 01/01/1970 12:00 AM UTC. A spreadsheet date counts days, so you divide by 86,400 (the seconds in a day) and add the date 01/01/1970. The formulas are the same in Excel and Google Sheets.

1. Paste the timestamps into column **B**, starting in **B4**.
2. Count the digits. 10 digits (such as 1793630100) are seconds. 13 digits (such as 1793725200000) are milliseconds.
3. For seconds, type `=B4/86400+DATE(1970,1,1)` in **D4**. For milliseconds, type `=B4/86400000+DATE(1970,1,1)`.
4. To show your own time zone, add the offset in hours times 3,600 before dividing. With the offset in **B1**: `=(B4+$B$1*3600)/86400+DATE(1970,1,1)`. Type -5 for Eastern standard time, -6 for Central, -7 for Mountain, -8 for Pacific.
5. Format **D4** as a date and time with the custom format `mm/dd/yyyy h:mm AM/PM`. In Excel this is under **Format Cells > Number > Custom** ([Microsoft's guide](https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309)). In Google Sheets it is **Format > Number > Custom date and time** ([Google's guide](https://support.google.com/docs/answer/56470)).
6. To group by day, put `=INT(D4)` in **E4** and format it as mm/dd/yyyy. For the time alone, use `=D4-INT(D4)` in **F4**.
7. Fill **D4:F4** down to the last timestamp.

Worked example for 1793630100: divided by 86,400 it is 20759.6076. DATE(1970,1,1) is serial 25569, so the sum is 46328.6076, which is 11/02/2026 2:35 PM UTC. Minus 5 hours, that is 9:35 AM Eastern.

Microsoft's [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) page describes DATE; it is not a record of our test.

### Daylight saving time

**B1** is one fixed offset. It does not switch for daylight saving time. In 2026, US daylight saving time ended on 11/01/2026, so -5 is right for Eastern from then until it starts again in March; summer dates need -4. For a list that spans the change, convert in UTC (B1 = 0) or split the rows. For more on offsets, see [How to Convert Time Zones for Client Calls in Excel and Google Sheets](https://tidytabs.blogspot.com/2026/10/how-to-convert-time-zones-for-client.html).

## Example

Sample data is fictional: 12 made-up payments, converted with B1 = -5.

| Payment ID | Timestamp | Unit | Date and time (local) | Status |
|---|---|---|---|---|
| pay_0001 | 1793630100 | Seconds | 11/02/2026 9:35:00 AM | OK |
| pay_0002 | 1793650500 | Seconds | 11/02/2026 3:15:00 PM | OK |
| pay_0003 | 1793725200000 | Milliseconds | 11/03/2026 12:00:00 PM | OK |
| pay_0004 | 1793768399 | Seconds | 11/03/2026 11:59:59 PM | OK |
| pay_0005 | 1793768400 | Seconds | 11/04/2026 12:00:00 AM | OK |
| pay_0006 | 1793901045123 | Milliseconds | 11/05/2026 12:50:45 PM | OK |
| pay_0007 | 1794087000 | Seconds | 11/07/2026 4:30:00 PM | OK |
| pay_0008 | 1794157200 | Seconds | 11/08/2026 12:00:00 PM | OK |
| pay_0009 | 1794520800 | Seconds | 11/12/2026 5:00:00 PM | OK |
| pay_0010 | "1794575700" (text) | Seconds | 11/13/2026 8:15:00 AM | OK |
| pay_0011 | "2026-11-13T13:15:00Z" (text) | Check | Check | Check |
| pay_0012 | 17945208 (cut off) | Seconds | 07/27/1970 11:46:48 AM | Before 2000 |

pay_0004 and pay_0005 are one second apart and fall on different dates, because local midnight sits between them. pay_0006 also carries 0.123 seconds that the format does not show.

The summary block in **I3:J9** shows 12 timestamps listed, 10 OK, 1 Before 2000, 1 Check, the earliest OK time 11/02/2026 9:35:00 AM, the latest 11/13/2026 8:15:00 AM, and OK for the B1 offset.

## Troubleshooting

### The hours are off by 5 (or 4)
The short formula gives UTC. pay_0001 is 2:35 PM UTC and 9:35 AM at -5. Add the offset as in step 4.

### The date is thousands of years ahead, or an error shows
A 13-digit milliseconds value was divided by 86,400. In LibreOffice, pay_0003 gave a serial of about 20,786,277, far past the year 9999. Divide by 86400000 instead.

### The date is in 1970
Digits were cut off. The 8-digit value 17945208 gave 07/27/1970 11:46:48 AM. The template marks it Before 2000. Copy the full timestamp again.

### A row shows Check
The value is not a plain timestamp: ISO date text such as 2026-11-13T13:15:00Z, a minus sign, commas or spaces inside the number, or a value past 2100. In LibreOffice, text made only of digits (pay_0010) converted fine in both the template and the short formula, while the ISO text gave an error in the short formula. How Excel and Sheets treat text timestamps was not tested.

### Every row shows Check
**B1** is text, blank, or outside -12 to 14. Type a plain number such as -5.

## Template

[Download the .xlsx](templates/tidy-tabs-unix-timestamp-to-date.xlsx)

The file has two tabs. **Timestamps** has the offset in **B1**, payment IDs and timestamps in **A4:B203**, and formula columns **C:G** (Unit, Date and time, Date, Time, Status). Unit reads values below 100,000,000,000 as seconds and larger ones as milliseconds, and flags anything past 01/01/2100. The summary is in **I3:J9**. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Status is red for Before 2000 and Check. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
