---
title: How to Track Progress to a Monthly Sales Goal in Excel and Google Sheets
labels: excel-formulas, sales-tracking, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was compared with Python math, including an as-of date on the last day, a goal already beaten and a new month with no sales. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Track progress to a monthly sales goal with =B4/B1 and the daily amount still needed with EOMONTH. Free .xlsx template with sample data.
image_prompts: Small craft market table with a handwritten-style paper tally card pinned under a wooden clip, a brass cash box with the lid half open, and a stack of kraft paper bags, soft morning light, no screens, no readable text, no logos
image_alt: Cash box with open lid beside a stack of kraft paper bags and a tally card
threads: $8,000.00 monthly goal, $2,699.25 sold by 10/12?\nDivide what is left by the days left: =MAX(0,B1-B4)/(EOMONTH(B2,0)-B2).\nThat comes to about $278.99 a day. The template also flags whether you are On pace or Behind.
---
Divide what you have sold by the goal with `=B4/B1`, then find the daily amount you still need with `=MAX(0,B1-B4)/(EOMONTH(B2,0)-B2)`, where B1 is the goal, B4 the sales so far and B2 the as-of date. With an $8,000.00 goal and $2,699.25 sold by 10/12/2026, you need about $278.99 a day for the rest of the month.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and compared with Python math. The Excel and Google Sheets steps use the same functions, and the Excel steps link to Microsoft's help pages. Neither Excel nor Google Sheets was opened for this post.

![Monthly sales goal sheet showing 33.7% of an $8,000.00 goal reached by 10/12/2026](images/2026-10-12-team-03-monthly-sales-goal-progress-excel-google-sheets/monthly-sales-goal-progress-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

You need two tabs: one that lists each sale with its date, and one that does the math. The sheet adds up only the sales dated in the month of your as-of date.

1. Name the first tab Daily sales. Type the headers Sale date in **A1** and Amount in **B1**.
2. Type one row per day, or per sale, from **A2** down. Use real dates such as 10/01/2026, not text. The formulas below read down to row 500.
3. Add a second tab named Goal. In **B1**, type your monthly goal, for example 8000.
4. In **B2**, type the as-of date. A typed date keeps the sample from changing. For a live sheet, type `=TODAY()` instead.
5. In **B3**, find the first day of the month: `=EOMONTH(B2,-1)+1`.
6. In **B4**, add up the month so far: `=SUMIFS('Daily sales'!$B$2:$B$500,'Daily sales'!$A$2:$A$500,">="&B3,'Daily sales'!$A$2:$A$500,"<="&B2)`.
7. In **B5**, type `=B4/B1` and format the cell as Percent. This is the share of the goal you have reached.
8. In **B6**, type `=MAX(0,B1-B4)`. This is the dollars left, and it never goes below $0.00.
9. In **B7**, type `=EOMONTH(B2,0)-B2`. This is the number of days left after the as-of date.
10. In **B8**, type `=IF(B6=0,"",IF(B7=0,"Last day",B6/B7))`. This is the amount you need per day.
11. In **B9**, type `=B4/DAY(B2)`. This is your average sales per day so far.
12. In **B10**, type `=B9*DAY(EOMONTH(B2,0))`. This is the month total if every day looked like your average.
13. In **B11**, type `=IF(B10>=B1,"On pace","Behind")`.

`SUMIFS` adds the amounts whose dates fall on or after the first of the month and on or before the as-of date. That is why a sale from the end of last month does not count.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628) pages. To format **B5** as Percent, select it and press **Ctrl+Shift+%**.

### In Google Sheets

The same formulas should work as typed, because Sheets has `SUMIFS`, `EOMONTH` and `DAY`. This was not opened in Google Sheets for this post. To format **B5** as Percent, select it and choose **Format > Number > Percent**.

## Example

Sample data below is fictional. It is a small shop's daily sales for the first part of October 2026.

| Sale date | Amount |
|---|---|
| 09/30/2026 | $250.00 |
| 10/01/2026 | $310.50 |
| 10/03/2026 | $185.00 |
| 10/04/2026 | $420.25 |
| 10/05/2026 | $96.00 |
| 10/07/2026 | $275.75 |
| 10/08/2026 | $142.50 |
| 10/09/2026 | $380.00 |
| 10/10/2026 | $210.00 |
| 10/11/2026 | $515.25 |
| 10/12/2026 | $164.00 |

The 09/30/2026 sale belongs to the previous month, so the Goal tab leaves it out. With a goal of $8,000.00 and an as-of date of 10/12/2026, it shows this:

| Cell | Meaning | Result |
|---|---|---|
| B3 | First day of month | 10/01/2026 |
| B4 | Sales so far | $2,699.25 |
| B5 | Share of goal | 33.7% |
| B6 | Dollars left | $5,300.75 |
| B7 | Days left | 19 |
| B8 | Needed per day | $278.99 |
| B9 | Average per day so far | $224.94 |
| B10 | Projected month total | $6,973.06 |
| B11 | Status | Behind |

At the current average of $224.94 a day, the month ends near $6,973.06, short of $8,000.00. Selling about $278.99 on each of the 19 remaining days closes the gap. Every value matched a Python check.

## Troubleshooting

### The daily amount shows "Last day" instead of a number
That happens when the as-of date is the last day of the month. With 10/31/2026, **B7** is 0, so **B8** shows Last day because there are no days left to divide by. In the recalculated file, **B9** also dropped to $87.07 and **B10** to $2,699.25, because the average is spread over all 31 days.

### The daily amount is blank
**B8** is blank when you have already reached the goal, since **B6** is $0.00. With a $2,000.00 goal, the recalculated file showed 134.96% in **B5**, $0.00 in **B6**, a blank **B8** and On pace in **B11**.

### Everything is $0.00 at the start of a new month
If you change the as-of date to 11/05/2026 before entering any November sales, **B3** becomes 11/01/2026 and **B4** is $0.00. **B7** is 25 and **B8** is $320.00, which is the whole goal spread over the days left. **B10** is $0.00 and **B11** says Behind. Add the November rows and it fills in.

### A sale is not counted
The date is probably text. The sheet matches sales by date, so dates must be real dates, not text. Retype the date so the cell is recognized as one. Text dates were not tested in any app. Also check that the row is within row 500.

### The projection looks too simple
**B10** is a straight-line average, not a forecast. It does not know about weekends, a holiday sale or a slow week. Use it as a quick gut check, not a promise.

## Template

[Download the .xlsx](templates/tidy-tabs-sales-goal-progress.xlsx)

The file has three tabs. **Daily sales** holds 11 sample rows in **A2:B12**. **Goal** has the yellow inputs in **B1** (goal $8,000.00) and **B2** (as-of date 10/12/2026), and the results in **B3:B11**. **How to use** has short fill-in notes, including the `=TODAY()` option. It covers one calendar month at a time.

The dates and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
