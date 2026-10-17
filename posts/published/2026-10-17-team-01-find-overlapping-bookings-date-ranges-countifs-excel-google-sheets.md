---
threads_url: https://www.threads.com/@tin_ylab/post/Demma6ymrRw
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-overlapping-dates-and.html
title: How to Find Overlapping Dates and Double Bookings in Excel and Google Sheets
labels: excel-formulas, small-business, scheduling
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against a Python pairwise overlap test, including a same-day handoff, a booking inside another, identical dates, three overlapping bookings, an end date before the start, a blank row in the table and a start date typed as text. Excel and Google Sheets were not opened; the help link is from Microsoft Support.
search_description: Find double bookings with =COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,"<="&E2,$E$2:$E$40,">="&D2)-1 in Excel and Google Sheets. Free template.
image_prompts: Paper wall calendar with two sticky notes stuck on overlapping days and a pen resting on a wooden market stall counter, soft daylight, no screens, no readable text, no logos
image_alt: Wall calendar with two sticky notes on neighboring days beside a pen on a counter
threads: Double-booked the same booth twice?\nOne COUNTIFS per row counts the other bookings for that space whose dates collide with yours.\nIn the sample, 9 of 12 bookings come out as a conflict, and a booking that starts the day another ends counts unless you switch it off.
---
Count the other bookings for the same space that start on or before this booking's end date and end on or after its start date with `=COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,"<="&E2,$E$2:$E$40,">="&D2)-1`. Any result above 0 is a double booking, and the `-1` removes the row itself.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Bookings table with a Conflicts count and a Conflict or OK status for each row](images/2026-10-17-team-01-find-overlapping-bookings-date-ranges-countifs-excel-google-sheets/overlapping-bookings-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The sheet has one booking per row: ID in **A**, space in **B**, client in **C**, start date in **D** and end date in **E**. Two bookings collide only when they share a space, so the formula matches column **B** first.

1. On a tab named **Bookings**, type these headers in row 1: Booking ID, Space, Client, Start date, End date, Conflicts, Status.
2. Format columns **D** and **E** as dates (MM/DD/YYYY) and type real dates, not text.
3. In **F2**, type `=COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,"<="&E2,$E$2:$E$40,">="&D2)-1`. The ranges run to row 40 so you can add bookings below the sample rows.
4. In **G2**, type `=IF(F2>0,"Conflict","OK")`.
5. Select **F2:G2** and drag the fill handle down to row 40. Fill only as far as you have bookings, because the short formulas were not tested on empty rows.
6. For a summary, put `=COUNTIF(G2:G40,"Conflict")` in one cell and `=COUNTA(A2:A40)` in another.

The template's own **F2** and **G2** are longer. They return a blank for an empty row and Check for a missing or text date, and they read the handoff switch in **K2**.

The `$` signs lock the ranges so they stay put when you fill down.

Two bookings that only share a day count as a conflict, because of `<=` and `>=`. If a booth is free the afternoon one client leaves and the next arrives, use `"<"&E2` and `">"&D2` instead. The row itself is then no longer matched, so a one-day booking needs `-IF(E2>D2,1,0)` rather than `-1`.

### In Excel

The formulas work as typed. Microsoft describes the function on its [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) page. To list only the problems, select the header row and use **Data > Filter**, then filter column **G** to Conflict. Microsoft explains filters on its [Filter data in a range or table](https://support.microsoft.com/en-us/office/filter-data-in-a-range-or-table-01832226-31b5-4568-8806-38c37dcc180e) page.

### In Google Sheets

The same formulas should work as typed, and the **Data > Create a filter** menu item should give you the same Conflict view. Neither was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is the full 12-row booking list in the template for October 2026, with the same-day handoff switch in **K2** set to Yes.

| Booking ID | Space | Start date | End date | Conflicts | Status |
|---|---|---|---|---|---|
| BK-101 | Booth A | 10/03/2026 | 10/05/2026 | 1 | Conflict |
| BK-102 | Booth A | 10/05/2026 | 10/07/2026 | 1 | Conflict |
| BK-103 | Booth A | 10/12/2026 | 10/20/2026 | 1 | Conflict |
| BK-104 | Booth A | 10/14/2026 | 10/16/2026 | 1 | Conflict |
| BK-105 | Booth B | 10/03/2026 | 10/05/2026 | 0 | OK |
| BK-106 | Booth B | 10/10/2026 | 10/14/2026 | 1 | Conflict |
| BK-107 | Booth B | 10/13/2026 | 10/17/2026 | 2 | Conflict |
| BK-108 | Booth B | 10/16/2026 | 10/19/2026 | 1 | Conflict |
| BK-109 | Studio | 10/08/2026 | 10/09/2026 | 1 | Conflict |
| BK-110 | Studio | 10/08/2026 | 10/09/2026 | 1 | Conflict |
| BK-111 | Studio | 10/21/2026 | 10/23/2026 | 0 | OK |
| BK-112 | Studio | 10/24/2026 | 10/24/2026 | 0 | OK |

Nine of the 12 bookings have a conflict. BK-101 and BK-102 collide only because one ends on the day the other starts (10/05/2026). BK-104 sits inside BK-103, BK-107 overlaps both BK-106 and BK-108, and BK-109 and BK-110 have identical dates. BK-105 has the same dates as BK-101 but a different space, so it is OK. A booking is flagged on every row it touches, so a clash always shows at least twice.

The template adds a second count on each row with SUMPRODUCT, and a check cell that compares the two totals. It showed OK in the recalculated copy.

## Troubleshooting

### A conflict disappears when one date is text
COUNTIFS ignores a date stored as text, so the other bookings stop seeing it. In the recalculated copy, the start date of BK-102 typed as `'10/05/2026` made that row show Check, and BK-101 silently dropped to 0 and OK. The conflict count fell from 9 to 7. Retype the date so it right-aligns, or format the column as a date first.

### The end date is before the start date
That row cannot be tested, so the template shows Check and leaves the count blank. In the recalculated copy, changing BK-105 to 10/09/2026 through 10/04/2026 gave Check and a Check count of 1. The conflict count stayed at 9.

### Two bookings that touch on one day count as a conflict
This is the default. Set the same-day handoff switch in **K2** to No and BK-101 and BK-102 both became OK, dropping the conflict count from 9 to 7. A one-day booking added inside BK-103 and BK-104 with the switch on No showed 2, and both of those bookings showed 2 as well.

### A blank row in the middle of the table
It stays blank. In the recalculated copy, clearing row 7 left the others working: the conflict count fell to 8 and total bookings to 11.

## Limits

This is date math only. It does not read lease, contract or deposit terms. Times of day are not considered, so two bookings on the same date are treated as the same day. The space dropdown holds three fixed names (Booth A, Booth B, Studio); change the validation list for your own.

## Template

[Download the .xlsx](templates/tidy-tabs-overlapping-bookings.xlsx)

The file has two tabs. **Bookings** holds the 12 sample rows in yellow input cells (A2:E40, with a Space dropdown), the Conflicts, Status and cross-check formulas in **F** to **H** filled to row 40, the same-day handoff switch in **K2** and the summary and check cells in **K3** to **K6**. **How to use** has short fill-in notes. **Bookings** prints landscape and **How to use** prints portrait, both on US Letter.

The bookings are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
