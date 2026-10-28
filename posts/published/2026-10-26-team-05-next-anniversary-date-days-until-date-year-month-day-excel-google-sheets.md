---
threads_url: https://www.threads.com/@tin_ylab/post/DfBGbSDDaAv
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-next-anniversary-of-date.html
title: How to Find the Next Anniversary of a Date and the Days Until It in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with Python datetime. Checked: all 366 month and day pairs against six as-of dates, 02/29 in a non leap year and in 2028, a date equal to the as-of date, 12/31 and 01/01 as-of dates, a window of 0, a first order date after the as-of date, a date typed as text, an as-of date typed as text, a date with a time, and an empty list. Excel, Google Sheets and TODAY were not used.
search_description: Find the next anniversary of a date and the days until it in Excel and Google Sheets. Free template with a typed as of date, a coming up flag and the 02/29 rule.
image_prompts: A wall calendar with a few days circled in pencil beside a small stack of handwritten thank you cards and a ceramic mug of pens on a wooden desk, soft window light, no screens, no readable text, no logos
image_alt: Wall calendar with penciled circles next to a stack of blank thank you cards
threads: Which customers hit a one year mark this month?\nOne formula builds each next anniversary: DATE(YEAR($B$1)+(...<$B$1),MONTH(B5),DAY(B5)). Subtract B1 for the days left.\nA 02/29 date lands on 03/01 in a year with no leap day.
---
Find the next yearly occurrence of a date with `=DATE(YEAR($B$1)+(DATE(YEAR($B$1),MONTH(B5),DAY(B5))<$B$1),MONTH(B5),DAY(B5))`, where B1 is the date you count from and B5 is the original date. Subtract B1 from the result for the days until it. A 02/29 date lands on 03/01 in a year with no leap day.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Customers sheet listing fictional businesses with first order dates, next anniversary, days until and green Yes rows](images/2026-10-26-team-05-next-anniversary-date-days-until-date-year-month-day-excel-google-sheets/next-anniversary-and-days-until-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: build the anniversary in the as-of year. If that date is before the as-of date, it has already passed, so use the next year instead. The comparison in the formula returns TRUE or FALSE, and adding it to the year adds 0 or 1.

The as-of date is typed, not `TODAY()`, so the sheet never changes on its own and you can check any number by hand. If you want it live, type `=TODAY()` in B1 yourself.

### In Excel

1. In **B1**, type the date you count from, such as 10/26/2026. In **B2**, type the number of days that counts as coming up, such as 30.
2. In row 4, type the headings: Name, First order date, Next anniversary, Days until, Years completed, Coming up.
3. Type names in **A5** and down, and the original dates in **B5** and down. Format column B as a date.
4. In **C5**, type the formula from the top of this post. Format the column as a date.
5. In **D5**, type `=C5-$B$1`. Format it as a number with no decimals.
6. In **E5**, type `=YEAR(C5)-YEAR(B5)`. This is the number of years completed on that anniversary.
7. In **F5**, type `=IF(D5<=$B$2,"Yes","")`.
8. Select **C5:F5** and drag the fill handle down the list.

How the formula reads:

- `DATE(YEAR($B$1),MONTH(B5),DAY(B5))` is this year's anniversary.
- `<$B$1` asks whether that date is before the as-of date. TRUE adds 1 to the year.

To color the rows, select the list and use **Home > Conditional Formatting > New Rule > Use a formula** with `=$F5="Yes"`. Microsoft's pages: [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349) and [Create a conditional formatting rule](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-fadb8b0e-6ae7-4b10-a7b0-e52e1ed9d9a1). These pages describe the menus and function, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 7 above, typed exactly as written, and format the date columns with **Format > Number > Date**.
2. Select **C5:F5**, then drag the fill handle down the list.
3. For color, select the list and choose **Format > Conditional formatting**. Under "Format cells if", pick **Custom formula is** and type `=$F5="Yes"`.

Google's help page: [Use conditional formatting rules in Google Sheets](https://support.google.com/docs/answer/78413). That page describes the menus, not our test.

### The version in the template

The template's C5 adds a guard so a blank row stays blank and a bad date shows a message instead of an error:

`=IF(B5="","",IF(AND(ISNUMBER(B5),ISNUMBER($B$1)),DATE(YEAR($B$1)+(DATE(YEAR($B$1),MONTH(B5),DAY(B5))<$B$1),MONTH(B5),DAY(B5)),"Check date"))`

The summary sits in **H5:I8**: **I5** counts the Yes rows, **I6** gives the nearest name, **I7** gives its days until, and **I8** is a check that says OK or CHECK.

## Example

Sample data is fictional: 20 businesses with made up first order dates, counting from 10/26/2026 with a 30 day window.

| Name | First order date | Next anniversary | Days until | Years completed | Coming up |
|---|---|---|---|---|---|
| Birch Street Bakery | 02/29/2020 | 03/01/2027 | 126 | 7 | |
| Lakeview Dental Office | 10/26/2019 | 10/26/2026 | 0 | 7 | Yes |
| Pine Hill Yoga | 10/25/2018 | 10/25/2027 | 364 | 9 | |
| Cedar & Co. Realty | 12/31/2021 | 12/31/2026 | 66 | 5 | |
| Oak Lane Florist | 11/10/2022 | 11/10/2026 | 15 | 4 | Yes |
| Harbor Coffee | 11/25/2023 | 11/25/2026 | 30 | 3 | Yes |
| Maple Print Shop | 11/26/2023 | 11/26/2026 | 31 | 3 | |

The summary for all 20 rows: **I5** is 8 coming up, **I6** is Lakeview Dental Office, **I7** is 0 days, and **I8** is OK.

Three rows show the rules. Lakeview's date is the same day as B1, so it is 0 days and counts as coming up. Pine Hill Yoga is one day earlier in the calendar, so its anniversary has passed and it waits until 10/25/2027. Harbor Coffee is exactly 30 days out and is in, while Maple Print Shop at 31 days is out.

Birch Street Bakery started on 02/29/2020. In 2027 there is no February 29, so the formula returns 03/01/2027.

## Troubleshooting

### Next anniversary shows Check date
The first order date is text, not a real date, so it left-aligns. In the test, a date typed as text showed Check date with no error value. Retype it as a date. If B1 is the text one, every row shows Check date.

### A 02/29 date shows 03/01
There is no February 29 in that year. `DATE` rolls 02/29 to 03/01 in a non leap year. When the as-of date is 02/29/2028 and the original is 02/29/2020, the result is 02/29/2028, 0 days.

### Days until is 0
The anniversary falls on the as-of date, so the row counts as coming up.  Same-day rows are meant to count, because that is the day to act.

### Years completed is 0 or negative, and I8 says CHECK
The first order date is after the as-of date. In the test, a date of 12/05/2027 against 10/26/2026 gave a next date of 12/05/2026 and -1 years. I8 turns to CHECK so you can find the row.

## Template

[Download the .xlsx](templates/tidy-tabs-next-anniversary.xlsx)

The file has two tabs. **Customers** holds the yellow as-of date in **B1** and the yellow window in **B2**. Yellow cells in **A5:B204** are for names and first order dates, with 20 fictional rows filled in. Gray cells are formulas: columns **C** to **F** and the summary in **H5:I8**. Rows with Yes turn green. **How to use** has short notes.

The file was recalculated only in LibreOffice Calc 24.2.7.2, and the green and red colors were seen in the PDF render, not in a spreadsheet window. It prints on US Letter, landscape for the Customers tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
