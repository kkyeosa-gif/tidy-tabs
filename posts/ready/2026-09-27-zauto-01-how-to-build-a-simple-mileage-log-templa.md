---
title: How to Build a Simple Mileage Log Template in Excel and Google Sheets
labels: mileage-log, freelance
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small notebook and a car key resting on a car dashboard, with a coffee shop receipt tucked in the notebook, no text visible on any surface
image_alt: Notebook and car key on a dashboard used for tracking business trips
search_description: Build a simple mileage log in Excel or Google Sheets with Date, Start, End, Purpose, and Miles columns, plus a running total formula.
threads: Freelancers: stop guessing your business mileage at tax time.\nA 5-column log (Date, Start, End, Purpose, Miles) with one SUM formula at the bottom is all you need.\nFill it in right after each drive, not in April.
---
Track business mileage in a five-column log: Date, Start Location, End Location, Purpose, and Miles, with a `SUM` formula totaling the Miles column. Fill it in right after each trip so you don't have to rebuild it from memory later.

> Works in: Excel (Microsoft 365, Excel 2019+) and Google Sheets. Steps come from Microsoft and Google help pages, not hands-on testing.

![Spreadsheet example with columns Date, Start, End, Purpose](images/2026-09-27-zauto-01-how-to-build-a-simple-mileage-log-templa/how-to-build-a-simple-mileage-log-template-in-exce-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

1. In row 1, type the headers: `Date`, `Start`, `End`, `Purpose`, `Miles`.
2. Widen the **Start** and **End** columns enough to fit an address or business name.
3. Under the **Miles** column, leave one row per trip and enter the mileage by hand, or subtract odometer readings if you track those in two extra columns.
4. Below your last row of data, add a totals row and enter `=SUM(E2:E200)` in the Miles column (adjust the range and column letter to match your sheet).
5. Format the **Date** column as a date so trips sort correctly.

### In Excel

*(Menu locations below match Excel for Windows. Excel for Mac and Excel for the web offer the same features under similar names, though the exact wording and placement can vary slightly.)*

1. Select the Date column, then go to **Home > Number Format** and pick **Short Date**.
2. To turn the Purpose column into a dropdown of common trip types (Client meeting, Supply run, Bank deposit), select the column, then **Data > Data Validation**, choose **List** under Allow, and type your options separated by commas. See Microsoft's guide on [creating a drop-down list](https://support.microsoft.com/en-us/office/create-a-drop-down-list-7693307a-59ef-400a-b769-c5402dda906d).
3. Freeze the header row with **View > Freeze Panes > Freeze Top Row** so headers stay visible as you scroll.

### In Google Sheets

1. Select the Date column, then **Format > Number > Date**.
2. For a Purpose dropdown, select the column, then **Data > Data validation**, choose **Dropdown**, and list your trip types. Details are in Google's [data validation help page](https://support.google.com/docs/answer/186103).
3. Freeze the header row with **View > Freeze > 1 row**.

## Example

Sample data below is fictional, for a freelance graphic designer based in Providence, RI.

| Date | Start | End | Purpose | Miles |
|---|---|---|---|---|
| 09/02/2026 | Home office | Client site, Warwick | Client meeting | 14.2 |
| 09/05/2026 | Home office | Staples, Cranston | Supply run | 6.8 |
| 09/09/2026 | Home office | Bank, downtown Providence | Bank deposit | 5.1 |
| 09/15/2026 | Home office | Client site, Warwick | Client meeting | 14.2 |
| **Total** | | | | **40.3** |

The Total row uses `=SUM(E2:E5)`. The IRS publishes its own standard mileage rate each year; check the current figure on [irs.gov](https://www.irs.gov/tax-professionals/standard-mileage-rates) before applying it to this log, since this post doesn't track or calculate a dollar deduction.

## Troubleshooting

### The SUM formula shows 0 or blank
Check that the Miles column contains numbers, not text. A trailing space or the word "miles" typed after the number turns the cell to text, and `
