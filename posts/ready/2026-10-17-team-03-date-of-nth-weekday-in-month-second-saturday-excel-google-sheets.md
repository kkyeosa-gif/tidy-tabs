---
title: How to Find the Date of the 2nd Saturday (or Any Nth Weekday) of a Month in Excel and Google Sheets
labels: excel-formulas, dates, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there. The Nth weekday and last weekday formulas were run for 2026, 2027 and 2028, all 12 months, all 7 weekdays and n = 1 to 5 (1,260 combinations) with 0 mismatches against the Python calendar module, and the Schedule tab was checked for 96 more cells. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Find the 2nd Saturday of any month with =DATE(B2,B3,1)+MOD(7-WEEKDAY(DATE(B2,B3,1),1),7)+7 in Excel and Google Sheets. Free schedule.
image_prompts: Folded canvas market tent leaning beside a folding table holding a stack of paper flyers and a wall calendar with several Saturdays circled in red marker, early morning light, no readable text, no logos
image_alt: Market tent, paper flyers and a wall calendar with circled Saturdays on a folding table
threads: Does your market run on the 2nd Saturday?\nOne formula finds that date for any month, and the same pattern works for any Nth weekday.\nIn the sample, November 2026 lands on 11/14, but a 5th Saturday that month shows No such date.
---
With the year in B2 and the month number in B3, `=DATE(B2,B3,1)+MOD(7-WEEKDAY(DATE(B2,B3,1),1),7)+7` returns the 2nd Saturday. Change the 7 inside MOD to another weekday number (1 is Sunday) and the last `+7` to `+7*(n-1)` for any other Nth weekday.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Schedule of the 2nd Saturday of each month in 2026](images/2026-10-17-team-03-date-of-nth-weekday-in-month-second-saturday-excel-google-sheets/nth-weekday-of-month-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The formula finds the first day of the month, jumps forward to the first matching weekday, then adds whole weeks. On a tab named **Dates**, labels go in column **A** and values in column **B**.

1. In **A1:B1** type the headers Item and Value. Type these labels in **A2:A10**: Year, Month (1 to 12), Weekday, Which one (1 to 5), Weekday number (1 = Sunday), Inputs OK?, Date of the Nth weekday, Date of the last such weekday, Check: result falls on that weekday.
2. Enter 2026 in **B2**, 11 in **B3**, Saturday in **B4** and 2 in **B5**.
3. Type Weekday list in **E1** and the days Sunday to Saturday in **E2:E8**. Sunday is 1 and Saturday is 7, which matches how `WEEKDAY(date,1)` counts.
4. Select **B4**, then add a dropdown that reads its list from **E2:E8**.
5. In **B6**, type `=IFERROR(MATCH(B4,$E$2:$E$8,0),"Check weekday")`. This turns the name into 1 to 7.
6. In **B7**, type this guard so a bad month cannot roll into the next year: `=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B5),ISNUMBER(B6)),IF(AND(B2>=1900,B2<=9999,B2=INT(B2),B3>=1,B3<=12,B3=INT(B3),B5>=1,B5<=5,B5=INT(B5)),"Yes","No"),"No")`
7. In **B8**, type the Nth weekday:
   `=IF(B7<>"Yes","Check inputs",IF(MONTH(DATE(B2,B3,1)+MOD(B6-WEEKDAY(DATE(B2,B3,1),1),7)+7*(B5-1))=B3,DATE(B2,B3,1)+MOD(B6-WEEKDAY(DATE(B2,B3,1),1),7)+7*(B5-1),"No such date"))`
8. In **B9**, type the last such weekday of the month: `=IF(B7<>"Yes","Check inputs",EOMONTH(DATE(B2,B3,1),0)-MOD(WEEKDAY(EOMONTH(DATE(B2,B3,1),0),1)-B6,7))`
9. Format **B8:B9** as a date, MM/DD/YYYY.

The MONTH test in **B8** matters. When the 5th Saturday would spill into the next month, the date no longer matches the month you asked for, so the cell says No such date.

### In Excel

The formulas work as typed. Microsoft describes them on its [WEEKDAY function](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a) and [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) pages. For the dropdown, use **Data > Data Validation**, choose **List** and select **E2:E8**. Microsoft describes this on its [Apply data validation to cells](https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) page.

### In Google Sheets

The same formulas should work as typed. For the dropdown, the usual menu path is **Data > Data validation** with a range of **E2:E8**. Neither the formulas nor that menu path were opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a market that opens on the 2nd Saturday of each month.

| Year | Month | Weekday | n | Result | Last Saturday |
|---|---|---|---|---|---|
| 2026 | 11 | Saturday | 2 | 11/14/2026 | 11/28/2026 |
| 2026 | 11 | Saturday | 5 | No such date | 11/28/2026 |
| 2028 | 2 | Tuesday | 5 | 02/29/2028 | 02/29/2028 |

The first row is the template as shipped. For the other two rows, change **B2** to **B5** to the values shown. The last column shows the last of that weekday, so the Tuesday row shows the last Tuesday. The Tuesday row is also a leap year check: the 5th Tuesday of February 2028 exists and is 02/29/2028.

The **Schedule** tab repeats the pattern for 12 rows, using the weekday and n from **Dates**. With the default 2nd Saturday and year 2026 it shows 01/10/2026, 02/14/2026, 03/14/2026, 04/11/2026, 05/09/2026, 06/13/2026, 07/11/2026, 08/08/2026, 09/12/2026, 10/10/2026, 11/14/2026 and 12/12/2026.

## Troubleshooting

### The month is typed as 13 and the date jumps to next year
Microsoft's [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) page says a month above 12 adds to the year, so `DATE(2026,13,1)` returns 01/01/2027, so a plain formula gives a believable but wrong date. The guard in **B7** catches a month of 13 or 0: in the recalculated copy, **B7** said No and B8 and B9 showed Check inputs. An n of 6, 0 or 2.5 gives the same message.

### The result says No such date
The 5th Saturday of 11/2026 does not exist, because November 2026 has only four Saturdays. In 2026, the 5th Saturday is missing in months 2, 3, 4, 6, 7, 9, 11 and 12. Use the last-weekday cell in **B9** when you mean the final one, which returned 11/28/2026 here.

### The first of the month is already the weekday you want
This works. For Thursday, n = 1 and 10/2026, the result is 10/01/2026, because the MOD part returns 0 days to move.

### The weekday shows Check weekday
**B6** found no match for the text in **B4**. Choose the name from the dropdown instead of typing it, and check for a stray space or a misspelling.

### The year is typed as text
A year entered as '2026 with a leading apostrophe is text. In the recalculated copy it gave Check inputs and the Schedule went blank. Retype it as a plain number.

## Limits

This is date math only. It does not know about holidays or a market's own rules, such as skipping a month. Years 1900 to 9999 work. The weekday must be spelled as in the dropdown list, and an n above 4 often gives No such date.

## Template

[Download the .xlsx](templates/tidy-tabs-nth-weekday-of-month.xlsx)

The file has three tabs. **Dates** holds the yellow input cells **B2:B5**, the formulas in **B6:B9** and a check cell in **B10** that confirms both results fall on the chosen weekday. **Schedule** has the headers Month, Date and Weekday and lists 12 months in rows 2 to 13, formatted mmmm yyyy and MM/DD/YYYY. **How to use** has short fill-in notes. All tabs print on US Letter, portrait.

The sample dates are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
