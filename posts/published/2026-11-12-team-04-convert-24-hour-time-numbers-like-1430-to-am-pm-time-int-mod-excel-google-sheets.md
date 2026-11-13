---
threads_url: https://www.threads.com/@tin_ylab/post/DfqaA4zDmi2
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-convert-24-hour-times-like-1430.html
title: How to Convert 24-Hour Times Like 1430 to AM and PM Times in Excel and Google Sheets
labels: excel-formulas, small-business, data-cleanup, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 8 workbooks and 4,840 comparisons with 0 mismatches. Checked: 0, 1, 5, 59, 60, 99, 100, 159, 160, 959, 1000, 1159, 1200, 1259, 1300, 2300, 2359, 2360, 2399, 2400, 2401, 1859, 1899, 9999, 10000, 123456, "0" to "00000", "0745", "745", "07 45", " 0745 ", "7:45", "14:30", "12a", "abc", "-5", -5, 905.5, 0.5, a real time, 1E3 as a number and as text, full width digits, "+930", ".930", "9.30", "930.", spaces only, blank, "0060", "1260", and 6 random sets of 200 rows. Probes of the short formula: 1460 gives 3:00 PM, 2400 gives 12:00 AM, "14:30" gives 12:00 AM, "0745" gives 7:45 AM. Excel and Google Sheets were not used. Not run: TEXT with "h:mm AM/PM" in a non-US language setting, more rows than the template holds.
search_description: Turn 1430 into 2:30 PM with =TIME(INT(B4/100),MOD(B4,100),0) and an h:mm AM/PM format, and flag bad values like 1460. Free Excel template.
image_prompts: A long curl of thermal receipt tape on a farmers market booth table next to a gray metal cash box and a small wind-up kitchen timer, morning light under a canopy, print unreadable, realistic photo, no screens, no readable text, no logos
image_alt: Curled receipt tape beside a gray cash box and a kitchen timer on a market table
threads: =TIME(INT(B4/100),MOD(B4,100),0) turns 1430 into 2:30 PM. Format the cell as h:mm AM/PM.\nThe catch: it also turns 1460 into 3:00 PM and 2400 into 12:00 AM without a word.\nThe template pads each value and shows Check for those instead (made-up tickets).
---
Turn a time exported as a number such as 1430 into a real time with `=TIME(INT(B4/100),MOD(B4,100),0)`, then format the cell as **h:mm AM/PM**: 1430 shows 2:30 PM and 905 shows 9:05 AM. INT(B4/100) is the hour and MOD(B4,100) the minutes. The short formula turns 1460 into 3:00 PM and 2400 into 12:00 AM without a warning, so the template pads and checks each value and shows Check instead.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every value below was recalculated and compared with a Python model (8 workbooks, 4,840 comparisons). The formulas use only IF, OR, TRIM, LEN, LEFT, RIGHT, MID, FIND, ISNUMBER, SUMPRODUCT, COUNTIF, MIN, MAX, TIME, INT, MOD and TEXT, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Thirteen sales tickets with exported times, converted times, a text column and a status](images/2026-11-12-team-04-convert-24-hour-time-numbers-like-1430-to-am-pm-time-int-mod-excel-google-sheets/24-hour-time-to-am-pm-template.png) *Exported times such as 1430 and 0745 turned into real times; 1460, 2400 and 14:30 are flagged. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. Card readers, time clocks and point of sale exports often write 24-hour (military) time as a plain number, so this works for any column like that.

1. Paste the exported times into column **B**, starting at **B4**.
2. In **C4**, type `=TIME(INT(B4/100),MOD(B4,100),0)`.
3. Format **C4** as **h:mm AM/PM**. In LibreOffice the new cell showed a time on its own. If your cell shows a decimal such as 0.604166, apply the format yourself.
4. Fill **C4** down to the last row.
5. Optional: for a label or message, type `=TEXT(C4,"h:mm AM/PM")` in **D4**. That result is text, so keep using column C for sorting and math.

How it works: dividing 1430 by 100 gives 14.3, and `INT` keeps the 14, the hour. `MOD(B4,100)` is the remainder after dividing by 100, which is 30, the minutes. `TIME(14,30,0)` builds the time from those two parts.

TIME returns a fraction of a day. The value behind 2:30 PM is 0.6041666..., which is why the cell format matters.

Lost leading zeros are not a problem. A number cannot keep the 0 in 0745, so it arrives as 745, and both give 7:45 AM. Short values work too: 30 is 12:30 AM and 5 is 12:05 AM.

Help pages: Microsoft's [TIME function](https://support.microsoft.com/en-us/office/time-function-9a5aff99-8f7d-4611-845e-747d0b8d5457) and [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3). They describe the functions, not our test.

### What the template adds

The short formula does not check anything. The template's **C4** first pads the value to 4 characters with `RIGHT("000"&TRIM(B4),4)`, so 905 becomes 0905. It then checks that all 4 characters are digits, that the hour is 0 to 23 and the minutes are 0 to 59, and only then calls `TIME` with the first 2 and last 2 characters. Anything else shows Check.

Column **E** shows OK or Check, with Check in red. Values with a colon (14:30), seconds (143015), decimals (905.5), minus signs, letters or more than 4 characters all show Check.

## Example

Sample data is fictional: 13 made-up sales tickets from a market day.

| Ticket | Time as exported | Time | Status |
|---|---|---|---|
| T-101 | 0745 (text) | 7:45 AM | OK |
| T-102 | 905 | 9:05 AM | OK |
| T-103 | 1200 | 12:00 PM | OK |
| T-104 | 1430 | 2:30 PM | OK |
| T-105 | 0 | 12:00 AM | OK |
| T-106 | 2359 | 11:59 PM | OK |
| T-107 | 1205 | 12:05 PM | OK |
| T-108 | " 815 " (text with spaces) | 8:15 AM | OK |
| T-109 | 30 | 12:30 AM | OK |
| T-110 | 1745 | 5:45 PM | OK |
| T-111 | 1460 | Check | Check |
| T-112 | 2400 | Check | Check |
| T-113 | 14:30 (text) | Check | Check |

The summary block shows 13 times listed, 10 converted, 3 rows showing Check, an earliest time of 12:00 AM and a latest time of 11:59 PM. Column D repeats each converted time as text, such as 2:30 PM, and stays blank for the Check rows.

Compare T-104 with T-111. With the short formula, 1430 gives 2:30 PM and 1460 gives 3:00 PM, which looks just as believable. The template shows Check for 1460 because 60 is not a valid minute.

## Troubleshooting

### 1460 shows 3:00 PM
The minutes part is 60, and `TIME` carries 60 minutes into the next hour. The short formula gives no warning. The template shows Check, so go back to the source and find the real time.

### 2400 shows 12:00 AM
`TIME` wraps 24 hours to 0, so 2400 becomes midnight at the start of the day. The template shows Check. If your export uses 2400 for the end of the day, type 0 or decide yourself how to treat it.

### A pasted 14:30 shows 12:00 AM
Text that already has a colon is not an HHMM number. In LibreOffice, the text "14:30" divided by 100 gave a time of 0:00:36, which displays as 12:00 AM. The template shows Check for it. A cell that already holds a real time also shows Check; it does not need converting.

### A time with spaces around it
A value such as " 815 " pasted as text still works in the template, because `TRIM` removes the spaces before the check. The sample row T-108 shows 8:15 AM.

## Template

[Download the .xlsx](templates/tidy-tabs-24-hour-time-numbers.xlsx)

The file has two tabs. **Times** has the exported times in **B4:B203** (numbers or text), the real time in **C**, the 12-hour text in **D**, the status in **E**, and a summary in **G3:H7** with times listed, converted, rows showing Check, and the earliest and latest time. **How to use** explains each column. The page is portrait, US Letter.

Yellow cells are inputs and gray cells are formulas. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
