---
threads_url: https://www.threads.com/@tin_ylab/post/De9xkgcidos
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-make-simple-gantt-chart-with.html
title: How to Make a Simple Gantt Chart With Conditional Formatting in Excel and Google Sheets
labels: excel-formulas, conditional-formatting, project-planning
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The rule formula was copied into plain cells over all 20 x 28 grid cells and matched a Python date check with 0 mismatches, as did the weekend rule. The rules were read back from the saved .xlsx with openpyxl, and a PDF export (US Letter landscape) showed the bars and gray weekends. Edge cases tried: end before start, text dates, a start with no end, a one-day task, tasks cut at the chart edges, a chart start that crosses into 2027, and text in B1. Excel and Google Sheets were not opened, and the rule import into them was not checked.
search_description: Make a simple Gantt chart in Excel and Google Sheets with one conditional formatting formula that fills the cells between each task's start and end date. Free template.
image_prompts: A corkboard with rows of colored paper strips pinned in a staggered stair-step pattern, a pencil and ruler on a craft table, and a stack of folded tote bags nearby, soft window light, no screens, no readable text, no logos
image_alt: Corkboard with staggered colored paper strips above a ruler, pencil and folded tote bags
threads: Gantt chart without a Gantt add-in: one conditional formatting rule.\n=AND(F$3>=$B4,F$3<=$C4) over the grid fills each cell between a task's start and end.\nThe sample runs 28 days from 11/02/2026, and NETWORKDAYS shows why "Market day" has 0 workdays.
---
Build a row of dates across the top, then add one conditional formatting rule with the formula `=AND(F$3>=$B4,F$3<=$C4)` so every cell between a task's start and end date fills in. Apply it to the whole grid once and each task row becomes a bar.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and the rule formula was checked against a Python date test. The Excel and Google Sheets steps come from the Microsoft and Google help pages linked below. Neither app was opened for this post, and we did not check how either one imports the rule from the .xlsx.

![Gantt chart grid with green task bars across 28 November days and gray weekend columns](images/2026-10-25-team-03-simple-gantt-chart-conditional-formatting-excel-google-sheets/simple-gantt-chart-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: the grid cells hold nothing. The rule looks at the date above each cell and the start and end dates to its left, and fills the cell when the date falls between them.

1. In **A3:C3** type the headings Task, Start and End. Type each task in column A and its dates in B and C, formatted as MM/DD/YYYY.
2. In **D4** type `=C4-B4+1` and copy it down. This counts both the start day and the end day.
3. In **F3** type the first day of the chart, such as 11/02/2026.
4. In **G3** type `=F3+1`, then copy it right across to **AG3**. That gives 28 days.
5. Select the grid, **F4:AG23**, with F4 as the active cell.
6. Add a conditional formatting rule that uses a formula (menu paths below) and paste `=AND(F$3>=$B4,F$3<=$C4)`.
7. Pick a fill color and click OK.

Why the formula has dollar signs in odd places:

- `F$3` locks the row, so every cell in the grid reads its date from row 3 and the column still moves.
- `$B4` and `$C4` lock the column, so every cell in a row reads that row's start and end dates.
- The formula is written for the active cell, F4. Excel and Sheets adjust it for every other cell in the range.

If you select the grid starting from a different cell, rewrite the formula for that cell. A formula written for F4 and applied from F5 is shifted one row, and every bar lands one task too low.

### In Excel

Select the grid, then go to **Home > Conditional Formatting > New Rule > Use a formula to determine which cells to format**. Paste the formula, click **Format**, choose a fill on the **Fill** tab and click **OK**. Microsoft describes the same dialog on [Use conditional formatting to highlight information](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-in-excel-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f).

### In Google Sheets

Select the grid, then go to **Format > Conditional formatting**. Under **Format rules**, change the condition to **Custom formula is** and paste the formula. Choose a fill under **Formatting style** and click **Done**. Google describes this on [Use conditional formatting rules in Google Sheets](https://support.google.com/docs/answer/78413).

### Gray out weekends

Add a second rule on the same range with `=WEEKDAY(F$3,2)>5`. `WEEKDAY` with the second argument 2 numbers Monday as 1, so Saturday is 6 and Sunday is 7. Put the bar rule first so a task that covers a Saturday still shows green.

## Example

Sample data is fictional: eight tasks for getting ready for a fall market booth, starting 11/02/2026.

| Task | Start | End | Days | Workdays |
|---|---|---|---|---|
| Pick market and pay booth fee | 11/02/2026 | 11/03/2026 | 2 | 2 |
| Order tote bags and mugs | 11/02/2026 | 11/06/2026 | 5 | 5 |
| Print signs and price tags | 11/05/2026 | 11/09/2026 | 5 | 3 |
| Build display shelves | 11/09/2026 | 11/13/2026 | 5 | 5 |
| Pack inventory and cash box | 11/16/2026 | 11/18/2026 | 3 | 3 |
| Test tent setup in driveway | 11/17/2026 | 11/17/2026 | 1 | 1 |
| Market day | 11/21/2026 | 11/22/2026 | 2 | 0 |
| Count sales and restock | 11/23/2026 | 11/27/2026 | 5 | 5 |

The chart covers 11/02/2026 to 11/29/2026. "Print signs" starts on Thursday 11/05 and ends on Monday 11/09, so it has 5 days but only 3 workdays. "Market day" falls on a Saturday and Sunday, so it shows 2 days and 0 workdays. The template's **E** column uses `=NETWORKDAYS(B4,C4)`, which skips weekends but does not know about holidays.

Change the start date in **B1** to 12/28/2026 and the headings roll into 2027 without any edits. Change the end of "Order tote bags and mugs" to 11/10/2026 and the bar grows by four columns.

This is a planning picture, not a scheduler. Nothing moves when a task slips unless you change its dates.

## Troubleshooting

### No bars show up
Either the rule is not applied to the whole grid or the formula was written for the wrong cell. Open the rule list and check that **Applies to** is F4:AG23 and that the formula starts with `F$3` and `$B4`. Dates typed as text such as "Nov 2" also stop the bars, because the comparison needs real dates.

### The Days column says Check dates
The end date is before the start date, a date is text, or only one of the two is filled in. The template draws no bar for that row and shows CHECK in the summary cell **AJ6**. Retype the dates.

### A bar stops at the right edge
The task ends after the 28th day. The template cuts it off at the edge, and **AJ5** counts these tasks. Copy the last date cell further right if you need more days, and extend the rule's range to match.

### The bars are missing after importing into Google Sheets
We did not check whether Sheets keeps the rule when it imports the .xlsx. Look under **Format > Conditional formatting**. If the rule is gone, add it again with the steps above.

## Template

[Download the .xlsx](templates/tidy-tabs-simple-gantt-chart.xlsx)

The file has two tabs. **Plan** holds the yellow inputs: the first chart day in **B1** and up to 20 tasks in **A4:C23**, with eight filled in. Gray cells are formulas: Days in **D**, Workdays in **E**, the 28 date headings in **F3:AG3** and weekday names above them. The summary at **AI3:AJ6** shows the task count, the last day in the chart, how many tasks run past it and a check that says OK or CHECK. One rule fills the bars on **F4:AG23** and a second grays the weekends. **How to use** has short notes.

There is no "today" line, and the file does not use `TODAY()`, so it never changes by itself. The project and dates are fictional. The file was recalculated and checked only in LibreOffice Calc 24.2.7.2, and it prints on US Letter, landscape for the Plan tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file, then check that the conditional formatting came across.
