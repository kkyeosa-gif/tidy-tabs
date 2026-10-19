---
threads_url: https://www.threads.com/@tin_ylab/post/DerwABdmvZT
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-highlight-dates-that-are-coming.html
title: How to Highlight Dates That Are Coming Up in the Next 30 Days in Excel and Google Sheets
labels: excel-formulas, small-business, conditional-formatting
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every Days left and Status value was checked against a Python date calculation, plus edge cases (expires on the as-of date, exactly 30 and 31 days out, yesterday, a blank date, a date typed as text, a window of 0, and a text as-of date). The conditional formatting rules were read back from the saved .xlsx and the colors were seen only in the LibreOffice PDF export. Nothing was applied by hand in Excel or Google Sheets, and the highlighting was not watched live in any app.
search_description: Highlight dates due in the next 30 days with a =$E2="Soon" conditional formatting rule in Excel and Google Sheets. Free expiring dates template.
image_prompts: Three small glass jars of hand lotion and a bag of soy wax flakes on a workshop shelf, each with a blank paper tag and a pencil beside them, soft natural window light, no screens, no readable text, no logos
image_alt: Lotion jars and a wax flake bag on a shelf with blank paper tags
threads: Expiring stock hiding in a long list?\nAdd a Status column (Expired / Soon / OK), then color rows with =$E2="Soon". Change the 30 day window in one cell and every row follows.
---
Highlight dates inside the next 30 days by adding a Status column and a conditional formatting formula rule such as `=$E2="Soon"` with an amber fill. Status in column E compares the expiration date with an as-of date you type in Summary B1 and a window of days in Summary B2.

The template reads the Status word instead of a cross-sheet date test, so no rule needs a reference to another sheet and the color always matches the printed word.

> Works in: LibreOffice Calc 24.2 (Linux), where the Days left and Status values were checked against Python. The rules were read back from the saved file and the colors seen in a LibreOffice PDF export only. Neither Excel nor Google Sheets was opened for this post, so the live highlighting was not watched in either app.

![Item list with expired rows in red, due soon rows in amber, later dates plain](images/2026-10-19-team-01-highlight-expiring-dates-next-30-days-conditional-formatting-excel-google-sheets/expiring-dates-status-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The template has two working tabs. **Items** lists one item per row: item in **A**, type in **B**, expiration date in **C**, days left in **D** and status in **E**. **Summary** holds the as-of date in **B1** and the warning window in days in **B2**.

1. On **Summary**, type your as-of date in **B1** (for example 10/19/2026) and the window in **B2** (30). The sheet does not use TODAY, so type a new date when you open it.
2. On **Items**, type each date in column **C** as a real date, not text.
3. In **D2**, type `=IF(OR(C2="",NOT(ISNUMBER(C2)),NOT(ISNUMBER(Summary!$B$1))),"",C2-Summary!$B$1)`. The part that does the work is `C2-Summary!$B$1`. A negative number means the date has passed.
4. In **E2**, type `=IF(C2="","",IF(OR(NOT(ISNUMBER(D2)),NOT(ISNUMBER(Summary!$B$2))),"Check",IF(D2<0,"Expired",IF(D2<=Summary!$B$2,"Soon","OK"))))`. An item that expires on the as-of date is Soon. One that expired the day before is Expired.
5. Select **D2:E2** and drag the fill handle down to row 60.
6. Select **A2:E60** and add a conditional formatting rule that uses a formula: `=$E2="Expired"`, with a red fill. Add a second rule, `=$E2="Soon"`, with an amber fill.
7. Optional: select **E2:E60** and add a third rule, `=$E2="Check"`, with a light red fill for dates typed as text.

The `$` before **E** locks the column, so every cell in a row follows that row's status.

### In Excel

Menu path from Microsoft help, not run by us. Select the range, then choose **Home > Conditional Formatting > New Rule > Use a formula to determine which cells to format**. Microsoft describes the dialog on its page about [using conditional formatting to highlight information](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f).

### In Google Sheets

Menu path from Google help, not run by us. Select the range, choose **Format > Conditional formatting**, then set the condition to **Custom formula is** and type the formula. Google documents the panel in [Use conditional formatting rules in Google Sheets](https://support.google.com/docs/answer/78413).

## Example

Sample data below is fictional. It is a small maker business with lotion, candle wax, booth banner orders and market vendor ID cards, with an as-of date of 10/19/2026 and a 30 day window.

| Item | Expires | Days left | Status |
|---|---|---|---|
| Oat milk lotion 8 oz | 10/18/2026 | -1 | Expired |
| Lavender hand lotion | 10/19/2026 | 0 | Soon |
| Soy wax flakes 10 lb | 11/18/2026 | 30 | Soon |
| Beeswax pellets 5 lb | 11/19/2026 | 31 | OK |
| Fall market banner | 10/31/2026 | 12 | Soon |
| Holiday market banner | 09/19/2026 | -30 | Expired |
| Vendor ID card, Ana | 10/24/2026 | 5 | Soon |
| Vendor ID card, Ben | 01/17/2027 | 90 | OK |
| Cedar candle wax 25 lb | 11/17/2026 | 29 | Soon |
| Vendor ID card, Eli | 10/07/2026 | -12 | Expired |

The full list has 20 items: 3 Expired, 9 Soon and 8 OK. The check cell **Summary B9** says OK when those counts add up to the number of real dates and match a separate COUNTIFS count.

Exactly 30 days out is Soon, 31 is OK, and the as-of date itself is Soon. In the PDF export, Expired rows are red with bold dark red text, Soon rows are amber and OK rows are plain. That PDF is the only place the colors were seen.

## Troubleshooting

### A row is not colored but its Status says Soon
The rule range or formula is off. Check that the range starts at row 2 and the formula reads `$E2`, not `$E$2` or `E2`. With `$E$2` every row would follow the first item.

### Status says Check and Days left is blank
The date in column **C** is text, usually left-aligned. Retype it as a date. In the recalculated copy, such a row was left out of the Expired, Soon and OK counts, and the check cell **B9** said CHECK.

### Every status says Check
The as-of date in **B1** or the window in **B2** is text or blank. Type a number in each. **B9** says Check settings until you do.

## Limits

Dates only: no advice on shelf life, safety or regulatory deadlines. Dates with a time of day are not handled specially. Rows 2 to 60 only.

## Template

[Download the .xlsx](templates/tidy-tabs-expiring-dates.xlsx)

The file has three tabs, all portrait on US Letter. **Items** has yellow input columns **A:C** with 20 sample rows (rows 2 to 21), formulas in **D:E** down to row 60 and the three color rules on **A2:E60**. **Summary** has the as-of date and window in yellow cells **B1:B2**, the counts in **B3:B8** and the check in **B9**. **How to use** has short notes.

The items are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
