---
title: How to Make a Monthly Sales Chart With a Goal Line in Excel and Google Sheets
labels: excel-charts, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was built with a script, recalculated there and compared with a Python model, including edited copies (goal 5000, goal 0, a blank month, text typed as sales or as the goal, every month exactly $4,000). The chart XML was read back and the page was exported to PDF and viewed as an image. Excel and Google Sheets were not opened, and whether the chart imports there was not tested.
search_description: Make a monthly sales chart with a goal line in Excel and Google Sheets. Free template with a column and line chart, a goal cell and a months at or above goal count.
image_prompts: A wooden shop counter with a kraft paper gift box tied with twine, a spool of red baker's twine and a small brass bell, soft window light, no screens, no readable text, no logos
image_alt: Kraft gift box tied with twine beside a spool of red twine and a brass bell
threads: Want a goal line on your sales chart without drawing it by hand?\nType the goal once in B1, put =$B$1 beside every month, and chart Goal as a line.\nChange B1 and the whole line moves. In the sample, 7 of 12 months hit $4,000.00. Shop is made up.
---
Put your monthly goal in one cell, repeat it beside every month with `=$B$1`, then chart Sales as columns and Goal as a line. Change B1 and every Goal cell, and so the whole line, moves with it.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built, recalculated and exported to PDF to check the chart. The Excel and Google Sheets steps use menu paths from the Microsoft and Google help pages. Neither app was opened for this post, and the chart was not tested after import into either one.

![Column chart of 12 months of gift shop sales with a dashed red goal line at $4,000](images/2026-10-26-team-02-monthly-sales-column-chart-with-goal-line-excel-google-sheets/monthly-sales-chart-with-goal-line-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: a goal line is just a second series. Every row of the Goal column holds the same number, so the series draws as a flat line across the columns.

1. In **B1**, type your monthly goal, for example 4000.
2. In **A4:D4**, type the headings Month, Sales, Goal and Difference.
3. In **A5:A16**, type the month names, Jan to Dec.
4. In **B5:B16**, type the sales for each month. Leave a month blank if it has not happened yet.
5. In **C5**, type `=$B$1` and fill it down to **C16**. The dollar signs lock the reference, so every row points at B1.
6. In **D5**, type `=IF(AND(ISNUMBER(B5),ISNUMBER(C5)),B5-C5,"")` and fill it down. Sales minus goal: positive is above goal, negative is below. The `IF` keeps the cell blank if sales or goal is text.
7. Select **A4:C16**, the months, Sales and Goal, and insert a column chart.
8. Change the Goal series to a line.

### In Excel

1. Select **A4:C16**.
2. Go to **Insert > Chart** and pick a clustered column chart.
3. Right-click the Goal series and choose **Change Series Chart Type**, then pick Line for Goal.

Microsoft's page covers the chart basics: [Create a chart from start to finish](https://support.microsoft.com/en-us/office/create-a-chart-from-start-to-finish-0baf399e-dd61-4e18-8a73-b3fd5d5680c2). That is the help page, not our test.

### In Google Sheets

1. Select **A4:C16**.
2. Go to **Insert > Chart**. In the chart editor, choose Column chart (or Combo chart) as the chart type.
3. Under **Customize > Series**, pick Goal and set its type to Line.

Google's page covers the editor: [Add & edit a chart or graph](https://support.google.com/docs/answer/63824).

### What the template adds

The built template has the same formulas plus a summary in **A18:D25**, with labels in A and values in D:

- **D18** counts months at or above goal: `=COUNTIF(D5:D16,">=0")`.
- **D19** counts months with sales: `=COUNT(B5:B16)`.
- **D20** names the best month with `INDEX` and `MATCH`, and **D21** gives its sales.
- **D22** is total sales and **D23** is the goal for the months you entered: `=SUMPRODUCT(--ISNUMBER(B5:B16),C5:C16)`.
- **D24** is total sales minus that goal.
- **D25** says OK, or CHECK if the goal is not a number, a sales cell is text, or sales sit on a row with no month name.

## Example

Sample data is fictional: a small gift shop with a goal of $4,000.00 a month.

| Month | Sales | Goal | Difference |
|---|---|---|---|
| Jan | $3,184.50 | $4,000.00 | -$815.50 |
| Feb | $2,860.00 | $4,000.00 | -$1,140.00 |
| Mar | $3,642.25 | $4,000.00 | -$357.75 |
| Apr | $4,118.00 | $4,000.00 | $118.00 |
| May | $4,655.75 | $4,000.00 | $655.75 |
| Jun | $4,290.00 | $4,000.00 | $290.00 |
| Jul | $3,915.50 | $4,000.00 | -$84.50 |
| Aug | $4,430.25 | $4,000.00 | $430.25 |
| Sep | $3,748.00 | $4,000.00 | -$252.00 |
| Oct | $4,902.50 | $4,000.00 | $902.50 |
| Nov | $5,811.00 | $4,000.00 | $1,811.00 |
| Dec | $7,236.25 | $4,000.00 | $3,236.25 |

January: $3,184.50 - $4,000.00 = -$815.50. Seven months are at or above goal (Apr, May, Jun, Aug, Oct, Nov, Dec). The best month is Dec at $7,236.25. Total sales are $52,794.00 against a goal total of $48,000.00, a difference of $4,794.00. D25 says OK.

In the LibreOffice PDF render, the chart shows 12 dark green columns and a dashed red line at $4,000. With the goal changed to 5000, the line moved to $5,000 and 2 months were at or above goal. This is arithmetic on typed numbers. It does not forecast sales or suggest what your goal should be.

## Troubleshooting

### The goal line did not move when I changed B1
The Goal cells follow B1 only if B1 holds a number. If you typed something like "4k", the Goal column shows that text, Difference goes blank and D25 says CHECK. Retype B1 as a plain number such as 4000.

### A month has no column
A blank Sales cell draws no column, not a zero-height one. That is correct for a month that has not happened. The goal line still crosses that month. If the month has sales, type them as a number.

### The Difference column is blank
Sales or Goal on that row is text. The `IF(AND(ISNUMBER(...)))` guard blanks Difference instead of showing an error. Retype the sales as a number and D25 returns to OK. A sales entry like "abc" is skipped by the totals.

### The chart looks different after I open the file in Excel or Google Sheets
We did not test the import. After you open the file, check the chart's series. Sales should read B5:B16 as columns and Goal should read C5:C16 as a line, with A5:A16 as the months. If either series is wrong, rebuild the chart with the steps above.

## Template

[Download the .xlsx](templates/tidy-tabs-sales-vs-goal-chart.xlsx)

The file has two tabs. **Chart** holds the yellow inputs, the goal in **B1**, the months in **A5:A16** and the sales in **B5:B16**. Gray cells are formulas: Goal in **C**, Difference in **D** and the summary with its check in **A18:D25**. The chart sits beside the table, with axis titles and dollar labels, and the tab prints on one US Letter landscape page. **How to use** has short notes.

The shop and the numbers are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, and the chart was seen only in the PDF export, not in a spreadsheet window. To open it in Google Sheets, go to **File > Import > Upload** and select the file, then check both chart series.
