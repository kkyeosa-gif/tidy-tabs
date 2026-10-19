---
title: How to Add the Same Cell Across Monthly Sheets in Excel and Google Sheets
labels: excel-formulas, small-business, monthly-reports
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and the Sales, Supplies and Fees totals were checked against a Python sum of the four tabs, plus 15 edited variants (an inserted tab, a tab after April, a misspelled tab name, a renamed tab, a text amount, an empty month, negatives). Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Add the same cell across monthly sheets with =Jan!B2+Feb!B2+Mar!B2+Apr!B2 or =SUM(Jan:Apr!B2) in Excel and Google Sheets. Free monthly rollup template.
image_prompts: Four manila folders fanned out on a worktable, each with a blank tab, next to a stack of paper receipts held by a binder clip and a rubber-banded bundle of envelopes, soft natural window light, no screens, no readable text, no logos
image_alt: Fanned manila folders and a clipped receipt stack on a worktable
threads: Add B2 from every monthly tab with =Jan!B2+Feb!B2+Mar!B2+Apr!B2.\nExcel can shorten it to =SUM(Jan:Apr!B2), but Google Sheets has no 3D reference like that.\nThe template has a check cell that says CHECK when the methods disagree.
---
Add cell B2 from each monthly tab with `=Jan!B2+Feb!B2+Mar!B2+Apr!B2`, which works in both apps. Excel can shorten it to `=SUM(Jan:Apr!B2)`, but Google Sheets does not support that 3D reference, and the Sheets side was not tested for this post.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every total below was recalculated and checked against Python. The formulas for Excel and Google Sheets are the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Summary tab with Sales, Supplies and Fees totals by three methods and a month table](images/2026-10-20-team-02-sum-same-cell-across-monthly-sheets-3d-sum-indirect-excel-google-sheets/sum-across-monthly-sheets-template.png) *PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

This assumes one tab per month and the same layout on every tab: Sales in **B2**, Supplies in **B3**, Fees in **B4**.

1. Name the tabs Jan, Feb, Mar and Apr. Put the same labels in **A2:A4** and the amounts in **B2:B4** on each one.
2. Add a tab named **Summary**. In **A2:A4**, type Sales, Supplies and Fees.
3. In **Summary!B2**, type `=Jan!B2+Feb!B2+Mar!B2+Apr!B2`. Each piece is a tab name, an exclamation point and a cell address.
4. In **B3**, type `=Jan!B3+Feb!B3+Mar!B3+Apr!B3`. In **B4**, do the same with row 4.
5. To add a month, type `+` and the new tab's cell at the end of the formula.

That is the version for Google Sheets. It has no 3D reference and no INDIRECT, so nothing depends on how the app reads tab order.

### In Excel

A 3D reference adds the same cell on every tab between two tab names. In **C2**, type `=SUM(Jan:Apr!B2)` and copy it down to **C4**. The range goes by tab position, left to right, so only tabs between Jan and Apr are included. Microsoft describes the function on its [SUM function](https://support.microsoft.com/en-us/office/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89) page. In the template this formula was run in LibreOffice only.

### In Google Sheets

Skip the 3D reference, since Sheets does not support it. In the template, column C would show an error there. Use the plus formula above. This was not run in Google Sheets.

### Pick tabs by typed name with INDIRECT

If you want to list the months and let a formula find the tabs, type each tab name in **Summary!A9:A20**. In **B9**, type:

`=IF($A9="","",IF(ISERROR(INDIRECT("'"&$A9&"'!B2")),"Check sheet name",INDIRECT("'"&$A9&"'!B2")))`

Use `!B3` in **C9** and `!B4` in **D9**, then fill down to row 20. Type `=SUM(B9:B20)` in **B21** and copy it across to **D21**. INDIRECT builds a cell address from text. Microsoft describes it on its [INDIRECT function](https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261) page, and Google on its [INDIRECT](https://support.google.com/docs/answer/3093377) page. It was run in LibreOffice only.

The `ISERROR` wrapper matters. In LibreOffice, `IFERROR` did not catch the error from a misspelled tab name, but `ISERROR` did.

## Example

The sample data is fictional: a small Etsy shop, January through April.

| Tab | Sales | Supplies | Fees |
|---|---|---|---|
| Jan | $3,150.00 | $820.40 | $252.00 |
| Feb | $2,875.50 | $640.25 | $230.04 |
| Mar | $3,980.00 | $1,105.80 | $318.40 |
| Apr | $4,410.75 | $960.10 | $352.86 |
| **Total** | **$14,416.25** | **$3,526.55** | **$1,153.30** |

The plus formula, the 3D reference and the INDIRECT table all gave these same totals in LibreOffice Calc. **B23** shows OK when the plus and INDIRECT totals agree. **B24** shows OK when the 3D totals match the plus totals, or "Not supported here" in an app with no 3D reference.

## Troubleshooting

### The INDIRECT row says Check sheet name
The typed name does not match a tab. A misspelling such as "Fbe" in **A10** made that row show Check sheet name, left it out of the total and turned **B23** to CHECK. Retype the name exactly as the tab shows it.

### I renamed a tab and one total dropped
Formulas that point at a tab follow a rename in Excel. Typed text does not. The INDIRECT row for the old name shows Check sheet name until you retype the name in column A. The rename was simulated by editing the file, not done in an app.

### A new tab is in one total but not the others
In LibreOffice, a tab dragged between Jan and Apr was included by the 3D SUM, while the plus formula and INDIRECT table skipped it. **B24** then said CHECK. Add the new tab to the plus formula and type its name in column A. Excel's behavior here was not run.

### A text amount changes the total
A cell holding the text "n/a" makes the plus formula show `#VALUE!`, while SUM and INDIRECT skip it. Text that looks like a number, such as $100 typed as text, was converted by the plus formula in LibreOffice but skipped by SUM, so the totals differed and the checks said CHECK. Retype the cell as a number.

### Summary shows an error in column C in Google Sheets
That is the 3D reference. Use column B instead.

## Template

[Download the .xlsx](templates/tidy-tabs-sum-across-monthly-sheets.xlsx)

The file has six tabs. **Jan**, **Feb**, **Mar** and **Apr** hold the sample amounts in yellow cells **B2:B4**. **Summary** has the plus version in **B2:B4**, the 3D SUM in **C2:C4**, the INDIRECT table in **A9:D21** for up to 12 tabs, and the two checks in **B23** and **B24**. **How to use** has short notes. Every tab prints portrait on US Letter.

The data is fictional and the file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file. Column C will not calculate there.

These are sums only, not bookkeeping or tax advice.
