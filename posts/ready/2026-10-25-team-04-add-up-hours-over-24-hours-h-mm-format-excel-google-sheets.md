---
title: How to Sum Hours Over 24 in Excel and Google Sheets
labels: excel-formulas, small-business, time-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with Python integer-minute math, including 40 random sheets (all matched), a time typed as text, a 36 hour entry, a negative time, 23:59 plus three 10:00 entries and seven 0:59 entries. Displayed values were read from a CSV export with cell content saved as shown. Excel and Google Sheets were not opened, and negative totals were not tested.
search_description: Total hours showing 03:45 instead of 27:45? Format the sum as [h]:mm in Excel and Google Sheets. Free template with decimal hours and pay.
image_prompts: A round analog wall clock above a vintage time-clock punch card rack with blank cards, a mug of pencils and a hand-wound kitchen timer on a workshop counter, soft light, no screens, no readable text, no logos
image_alt: Analog wall clock over a rack of blank punch cards and a kitchen timer
threads: Your 49 hour week shows up as 1:00?\nFormat the total as [h]:mm. Plain h:mm restarts at zero every 24 hours.\nIn the sample, 9:30 + 8:15 + 10:00 reads 3:45 without the brackets and 27:45 with them.
---
Format the total cell with the custom number format `[h]:mm`, because plain `h:mm` restarts at zero every 24 hours. 9:30, 8:15 and 10:00 add up to 27:45 with `[h]:mm` but display as 3:45 without the brackets.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel steps follow Microsoft's help page and the Google Sheets steps follow Google's. Neither Excel nor Google Sheets was opened for this post. The Google Sheets duration format menu differs from Excel's and was not tested.

![Hours sheet with three weeks of times and totals shown as h:mm, bracketed hours, decimal hours and pay](images/2026-10-25-team-04-add-up-hours-over-24-hours-h-mm-format-excel-google-sheets/total-hours-over-24-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The sum itself is right. A time is stored as a fraction of a day, so 27 hours 45 minutes is 1.15625 days. The `h:mm` format shows only the part left over after whole days, which is 3:45. The brackets in `[h]:mm` tell the cell to keep counting hours.

1. Type each day's time as H:MM, such as 9:30, in **B2:B8**. Format those cells as **h:mm**.
2. In the total cell, type `=SUM(B2:B8)`.
3. Select the total cell and open the custom number format box (below).
4. Type `[h]:mm` and confirm. The total now reads 27:45.

The stored number does not change. Only the format does, so you can leave your daily cells as plain `h:mm`.

### In Excel

1. Select the total cell.
2. Choose **Home > Number Format > More Number Formats > Custom**.
3. Type `[h]:mm` in the Type box and select **OK**.

Microsoft lists the codes on its [Available number formats in Excel](https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2) page. That page describes the formats, not our test.

### In Google Sheets

1. Select the total cell.
2. Choose **Format > Number > Custom number format**.
3. Type `[h]:mm` and select **Apply**.

Sheets also has **Format > Number > Duration**, which is a separate menu item. We did not test it and do not know how it differs from `[h]:mm` on your sheet. Google explains number formats on its [Format numbers in a spreadsheet page](https://support.google.com/docs/answer/56470). The custom format steps come from the template notes, not from opening Sheets.

### Other ways to show the total

- **Decimal hours:** `=ROUND(SUM(B2:B8)*24,2)` gives 27.75. Multiply by 24 because the sum is in days.
- **Pay:** `=ROUND(SUM(B2:B8)*24*$E$8,2)` with your hourly rate in **E8**.
- **Text:** `=TEXT(SUM(B2:B8),"[h]:mm")` gives "27:45" as text you can join into a sentence. You cannot add it afterward.
- **Whole minutes:** `=ROUND(SUM(B2:B8)*1440,0)` gives 1665.

## Example

Sample data is fictional: three weeks of hours for a freelancer billing $45.00 an hour. Week 1 has only three days entered.

| Week | Shown as h:mm (wrong) | Shown as [h]:mm | Decimal hours | Pay at $45.00 |
|---|---|---|---|---|
| 1 (9:30, 8:15, 10:00) | 3:45 | 27:45 | 27.75 | $1,248.75 |
| 2 | 19:00 | 43:00 | 43.00 | $1,935.00 |
| 3 | 1:00 | 49:00 | 49.00 | $2,205.00 |
| All three weeks | 23:45 | 119:45 | 119.75 | $5,388.75 |

Look at the "wrong" column: week 3 shows 1:00 for a 49 hour week. The all-weeks total looks like less than one day of work for 119 hours.

The values above come from the template recalculated in LibreOffice Calc. Opening the file in another app is your own check.

## Troubleshooting

### The total still shows 3:45
The format was not applied to that cell, or the cell is a copy that kept `h:mm`. Reapply `[h]:mm` to the total cell itself, not to the daily cells.

### A time is skipped in the sum
A time typed as text, such as "8:15" with a leading apostrophe, is ignored by `SUM`. In the template, week 1 dropped by 8:15 when we tested this, and the check in **G10** turned to CHECK. Retype the cell as a time.

### A day holds 36 hours
Typing 1.5 means 1.5 days, which is 36 hours. Type 36:00 instead. The template adds it (55:30 for the edited week) and shows CHECK in **G10**, because a daily cell should stay under 24 hours.

### The total shows ####
A negative time cannot be shown in the default 1900 date system, according to the template's notes. We did not test a negative total. Subtract the smaller time from the larger one instead.

### The total is off by a few minutes
Check the daily cells first. Whole minutes in **K** show the exact total, so compare it with your own count. In our tests a total of 23:59 plus three 10:00 entries came to 63:29, and seven 0:59 entries came to 6:53.

## Template

[Download the .xlsx](templates/tidy-tabs-total-hours-over-24.xlsx)

The file has two tabs. **Hours** holds the yellow date and time cells in **A2:B22** (three weeks, Monday through Sunday) and the yellow hourly rate in **E8**. Gray cells are formulas: the same sum shown as `h:mm` in column F and `[h]:mm` in column G, decimal hours in H, pay in I, the `TEXT` version in J and whole minutes in K. **G10** checks that every hours cell is a time under 24 hours. **G11** checks that the text totals match the minutes. **How to use** has short notes.

The hours and the rate are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file, then check the number formats in the total cells.
