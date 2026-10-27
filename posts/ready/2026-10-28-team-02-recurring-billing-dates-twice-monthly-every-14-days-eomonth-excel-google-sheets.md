---
title: How to List Twice a Month and Every 14 Days Billing Dates in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and every date, weekday and weekend flag was compared with Python calendar and datetime for 13 start dates in both modes. Checked: 01/31, 02/15, 02/28, 12/31 and 02/29/2028 starts, 1 and 26 dates, and bad inputs (text date, Weekly, count 27 or 0). Excel and Google Sheets were not used.
search_description: List twice a month and every 14 days billing dates in Excel and Google Sheets with EOMONTH. Free template with a start date, a mode dropdown, weekday names and a weekend flag.
image_prompts: A paper desk calendar with the 15th and the last day of the month marked by small round stickers, a stack of blank envelopes, and a wooden date stamp on a studio desk, soft window light, no screens, no readable text, no logos
image_alt: Desk calendar with round stickers on two dates beside blank envelopes and a date stamp
threads: Invoice on the 15th and the last day, and February keeps throwing you off?\nEOMONTH(B7,0) finds the last day of any month, so 28, 30 and 31 need no special case.\nEvery 14 days is just B7+14. The template lists up to 26 dates and flags weekends.
---
For the 15th and last day of each month, fill down `=IF(DAY(B7)=15,EOMONTH(B7,0),DATE(YEAR(B7),MONTH(B7)+1,15))`, where B7 is the date above. For every 14 days, use `=B7+14` instead.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every date below was recalculated and compared with Python. The Excel and Google Sheets steps use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Dates sheet listing 24 twice-monthly billing dates with weekday names and weekend rows shaded pink beside a summary box](images/2026-10-28-team-02-recurring-billing-dates-twice-monthly-every-14-days-eomonth-excel-google-sheets/recurring-billing-dates-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: each date comes from the date above it. After the 15th comes the last day of the same month. After the last day comes the 15th of the next month. `EOMONTH(date,0)` returns the last day of the month, so February and 31 day months need no special case.

Dates only. The sheet does not say when a client must pay, and it does not move a date off a weekend. It flags weekend dates and leaves the choice to you.

### In Excel

1. In **C1**, type the start date, such as 11/15/2026. Format it as a date.
2. In **C2**, type the schedule: `Twice a month` or `Every 14 days`. To get a dropdown, select **C2** and use **Data > Data Validation > List**.
3. In **C3**, type how many dates you want, such as 24.
4. In row 6, type the headings: No., Date, Weekday, Weekend?
5. In **B7**, type the first date. For twice a month, the first date is the first 15th or month end on or after the start: `=IF(DAY($C$1)<=15,DATE(YEAR($C$1),MONTH($C$1),15),EOMONTH($C$1,0))`. For every 14 days, it is `=$C$1`.
6. In **B8**, type the next date. Twice a month: `=IF(DAY(B7)=15,EOMONTH(B7,0),DATE(YEAR(B7),MONTH(B7)+1,15))`. Every 14 days: `=B7+14`.
7. In **C7**, type `=TEXT(B7,"ddd")` for the weekday name. In **D7**, type `=IF(WEEKDAY(B7,2)>5,"Weekend","")`.
8. Select **B8:D8** and drag the fill handle down. Copy **C7:D7** down to match.

Microsoft's page for the month-end function: [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628). It describes the function, not our test.

### In Google Sheets

1. Type the same cells and formulas as in steps 1 to 8 above, exactly as written, and format the date cells with **Format > Number > Date**.
2. For the dropdown in **C2**, choose **Data > Data validation**, add a rule, and pick **Dropdown**.

Google's help pages: [EOMONTH](https://support.google.com/docs/answer/3093543) and [Add a drop-down list](https://support.google.com/docs/answer/186103). They describe the menus and function, not our test.

### The version in the template

The template folds both rules into one formula keyed to the dropdown, and leaves a row blank past the count you ask for. **B8** is:

`=IF(A8="","",IF($C$2="Every 14 days",B7+14,IF(DAY(B7)=15,EOMONTH(B7,0),DATE(YEAR(B7),MONTH(B7)+1,15))))`

**C4** says OK or explains what to fix. The summary sits in **G7:G11**: first date, last date, dates listed, dates on a Saturday or Sunday, and a CHECK that says OK when the count matches and every date is later than the one above.

## Example

Sample data is fictional: a start date of 11/15/2026, the Twice a month schedule and 24 dates.

| No. | Date | Weekday | Weekend? |
|---|---|---|---|
| 1 | 11/15/2026 | Sun | Weekend |
| 2 | 11/30/2026 | Mon | |
| 3 | 12/15/2026 | Tue | |
| 4 | 12/31/2026 | Thu | |
| 5 | 01/15/2027 | Fri | |
| 6 | 01/31/2027 | Sun | Weekend |
| 7 | 02/15/2027 | Mon | |
| 8 | 02/28/2027 | Sun | Weekend |

The summary for all 24 rows: **G7** is 11/15/2026, **G8** is 10/31/2027, **G9** is 24, **G10** is 7 weekend dates, and **G11** is OK.

Switch **C2** to Every 14 days and the list becomes 11/15/2026, 11/29/2026, 12/13/2026, 12/27/2026, 01/10/2027, 01/24/2027. Every one of those is a Sunday, so all of them are flagged Weekend.

## Troubleshooting

### The list starts on 11/30 instead of my start date
In Twice a month mode, the list starts on the first 15th or month end on or after the start date. A start of 11/16/2026 gives 11/30/2026, and a start of 11/10/2026 gives 11/15/2026. A start that already is the 15th or the last day is kept.

### The list is empty and C4 shows a message
**C4** names the problem. In the test, a start date typed as text gave "Type a date in C1". Typing Weekly in **C2** gave "Pick Twice a month or Every 14 days in C2". A count of 27, 0 or 12.5 gave "Dates to list must be a whole number from 1 to 26". Fix the cell and the list returns.

### February dates look wrong
They are correct. From a start of 01/31/2026 the list runs 01/31, 02/15, 02/28, 03/15, 03/31. From 02/29/2028 it runs 02/29, 03/15, 03/31. `EOMONTH` picks the real last day.

### Every 14 days drifts off the month end
That mode adds 14 days and nothing else. From 01/31/2026 you get 01/31, 02/14, 02/28, 03/14, 03/28. For dates tied to the 15th and month end, use the Twice a month schedule.

### A date I need falls on a weekend
The Weekend? column flags it. The sheet does not shift dates to a weekday, because that is your decision. The template does not use `WORKDAY` and has no holiday list.

## Template

[Download the .xlsx](templates/tidy-tabs-recurring-billing-dates.xlsx)

The file has two tabs. **Dates** holds three yellow inputs: the start date in **C1**, the schedule dropdown in **C2** and the count in **C3**, from 1 to 26. Gray cells are formulas: the list in **A7:D32** and the summary in **G7:G11**. Weekend rows turn red. **How to use** has short notes.

The file was recalculated only in LibreOffice Calc 24.2.7.2, and the colors were seen in the PDF render, not in a spreadsheet window. It prints on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
