---
threads_url: https://www.threads.com/@tin_ylab/post/DfIEvNylLNg
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-hours-worked-past.html
title: How to Calculate Hours Worked Past Midnight With an Unpaid Break in Excel and Google Sheets
labels: excel-formulas, time-tracking, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model (14 workbooks, 869 cell comparisons). Checked: end exactly at midnight, break equal to or one minute over the shift, start equal to end, times with seconds, a date part on the times, text and blank cells, and an invalid rate. Excel and Google Sheets were not used.
search_description: Calculate hours worked past midnight with an unpaid break in Excel and Google Sheets with MOD. Free shift template with decimal hours, pay, Review flag.
image_prompts: A paper shift schedule clipped to a small clipboard with a stack of paper wristbands and a brass stopwatch on a worn wooden counter, warm evening light, no screens, no readable text, no logos
image_alt: Brass stopwatch and paper wristbands beside a clipboard with a shift schedule on a wooden counter
threads: Your night shift hours come out negative?\nA plain end minus start fails when the end is after midnight.\nTry =MOD(D5-C5,1)-E5/1440 and format the cell as [h]:mm. 6:00 PM to 2:00 AM with a 30 minute break gives 7:30.
---
Subtract the times with `=MOD(D5-C5,1)-E5/1440`, where C5 is Start, D5 is End and E5 is the break in minutes, and format the result as `[h]:mm`. `MOD` adds back the missing day when the end is after midnight, so 6:00 PM to 2:00 AM counts as 8 hours before the break.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Shift table with hours, decimal hours, pay and a plain subtraction column showing negative times](images/2026-10-30-team-03-hours-worked-overnight-shift-unpaid-break-mod-excel-google-sheets/overnight-shift-hours-template.png) *caption: Overnight shift template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

Times are fractions of a day. 6:00 PM is 0.75 and 2:00 AM is about 0.083. Subtracting the start from the end gives a negative number, because the end is a smaller fraction than the start. `MOD(x,1)` keeps only the fraction of a day, so a negative result wraps around to the right positive length.

The break is typed in minutes. A day has 1,440 minutes, so `E5/1440` turns minutes into a time you can subtract.

The layout matches the template: Start in **C5:C64**, End in **D5:D64**, Break minutes in **E5:E64**, hourly rate in **B1**.

### In Excel

1. In **F5**, type `=MOD(D5-C5,1)-E5/1440`.
2. Select **F5**, press **Ctrl+1**, choose **Custom** and type `[h]:mm`. The square brackets keep the hours from resetting at 24.
3. In **G5**, type `=ROUND(F5*24,2)` for decimal hours.
4. In **H5**, type `=ROUND(G5*$B$1,2)` for pay at the rate in B1.
5. In **I5**, type `=D5-C5`. This is the plain subtraction, kept next to the fix so you can see the problem.
6. Select **F5:I5** and drag the fill handle down.

The template's F5 is a longer version, `=IF(J5="OK",MAX(0,MOD(D5-C5,1)-N(E5)/1440),"")`. `N()` makes a blank break count as 0, and `MAX(0,...)` removes a tiny negative leftover when the break equals the shift.

Microsoft's pages: [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) and [Create a custom number format](https://support.microsoft.com/en-us/office/create-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4). They describe the function and the format, not our test.

### In Google Sheets

1. Type the same formulas in **F5**, **G5**, **H5** and **I5**.
2. Select **F5**, choose **Format > Number > Custom number format**, and enter `[h]:mm`.
3. Fill **F5:I5** down.

Google's page: [MOD function](https://support.google.com/docs/answer/3093497). It describes the function, not our test.

## Example

Sample names, shifts and pay are fictional, at $18.00 an hour. The last column is plain `=D5-C5`.

| Start | End | Break | Hours (h:mm) | Decimal | Pay | Plain end - start | Review |
|---|---|---|---|---|---|---|---|
| 6:00 PM | 2:00 AM | 30 | 7:30 | 7.50 | $135.00 | -16:00 | OK |
| 5:30 PM | 11:30 PM | 30 | 5:30 | 5.50 | $99.00 | 6:00 | OK |
| 8:00 PM | 12:00 AM | 0 | 4:00 | 4.00 | $72.00 | -20:00 | OK |
| 10:00 PM | 10:00 PM | 30 | blank | blank | blank | 0:00 | Review: start equals end (0 or 24 hours?) |
| 5:00 PM | 5:20 PM | 30 | blank | blank | blank | 0:20 | Review: break is longer than the shift |
| 4:00 PM | 11:59 PM | 0 | 7:59 | 7.98 | $143.64 | 7:59 | OK |
| 6:00 PM | 12:30 AM | (blank) | 6:30 | 6.50 | $117.00 | -17:30 | OK |

The summary is in **L5:M10**. For all 14 sample rows, it shows 12 shifts with hours, 77.48 decimal hours, $1,394.64 in pay, 2 rows to review and 8 rows where the plain subtraction is negative. **M10** says OK.

The negative times in the last column show as a minus sign in LibreOffice. How Excel and Google Sheets display a negative time was not checked here.

## Troubleshooting

### The row says Review: start equals end
A start of 10:00 PM and an end of 10:00 PM could be a 0 hour shift or a 24 hour shift. The sheet does not guess, so the row gets no hours and no pay. The same happened when the times held a date part exactly 24 hours apart. Fix the time, or enter a shift under 24 hours.

### The row says Review: break is longer than the shift
The break minutes are bigger than the time between Start and End. A break equal to the shift (480 minutes on 6:00 PM to 2:00 AM) is allowed and gives 0:00 hours and $0.00. One minute more triggers the flag. Check whether the break was typed in hours by mistake.

### An 11:00 PM to 8:00 PM shift counts 21 hours
When the end is earlier than the start, `MOD` treats it as the next day. That is the point of the formula, but it also means a typo, such as 8:00 PM instead of 8:00 AM, produces 21:00 and no error. Scan the Plain column and the hours for values that look wrong.

### Hours and pay are blank for a row that looks fine
Start, End or Break may be text, such as "9:00" typed with a leading apostrophe. The row shows a Review message and no pay. Retype the cell as a time. A number stored as text usually left-aligns.

### The pay column says Check rate
The rate in **B1** is text, blank or negative. **B2** shows the reason. Type a number of 0 or more. A rate of 0 gives $0.00 pay, and **M10** shows CHECK until the rate is valid.

## Template

[Download the .xlsx](templates/tidy-tabs-overnight-shift-hours.xlsx)

The file has two tabs. **Shifts** holds a yellow hourly rate in **B1** and yellow input cells in **A5:E64**: Date, Name, Start, End and Break minutes. Sixty rows are available and 14 are filled with fictional night-market and event crew. Gray cells are formulas for hours, decimal hours, pay, the plain subtraction and the Review flag. **How to use** has short notes.

It assumes every shift is under 24 hours and one row is one shift. The pay column is only hours times a typed rate. It says nothing about overtime rules or what must be paid.

Calculations were checked in LibreOffice Calc 24.2.7.2 only. It prints on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
