---
title: How to Get the Season and Rate From a Date in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 22 workbooks and 44,748 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Tag any date with its season and nightly rate, even a Holiday season that runs past December 31, using a MONTH*100+DAY key and MATCH.
image_prompts: A small cabin rental check-in desk with a paper wall calendar above it marked in blocks of different colored highlighter, a set of cabin keys on a wooden key tag and a short stack of blank guest registration cards on the desk, warm daylight, realistic photo, no screens, no readable text, no logos
image_alt: Cabin keys on a wooden tag below a calendar marked with colored weeks
threads: A Holiday season from 12/15 to 01/05 is where a plain MATCH gives up with #N/A.\nTurn each date into =MONTH(date)*100+DAY(date), so Jan 2 is 102 in any year, and send keys below the first start to the last season.\n12/31/2026 and 01/01/2027 both come back Holiday at $249.00.
---
Turn each date into a month-day number with `=MONTH(date)*100+DAY(date)`, so Dec 15 is 1215 and Jan 2 is 102 in any year, then look up its season with `=IF(key<MIN(starts),INDEX(seasons,COUNT(starts)),INDEX(seasons,MATCH(key,starts,1)))`, which sends dates before the first start to the last season of the year. In the made-up cabin seasons, 12/31/2026 and 01/01/2027 both come back Holiday at $249.00, while the same MATCH without that IF gave #N/A for 01/01 to 01/05.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (22 workbooks, 44,748 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Season start table beside dates showing Holiday at $249.00 across December 31](images/2026-11-23-team-05-assign-season-and-rate-to-any-date-seasons-crossing-new-year-month-day-match-excel-google-sheets/season-and-rate-from-date-template.png) *Seasons tab with six fictional season starts beside the Dates tab, where 12/31/2026 and 01/01/2027 both fall in Holiday at $249.00, 01/06/2026 starts Off-season at $129.00, and each key such as 1215 is month times 100 plus day. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The sheet's rule: each season runs from its start month and day until the next start, every year. The last season of the year wraps into January and runs until the first start. Seasons and rates are yours; the sheet has no holiday calendar or lodging tax.

The steps are the same in Excel and Google Sheets.

1. On the **Seasons** tab, type one season per row from row 5 (up to 12 rows): Start month (1-12) in **A**, Start day in **B**, Season name in **C** and Rate $ in **D**, such as 12, 15, Holiday, $249.00.
2. Sort the rows by start month and day, January first. A season that crosses December 31, like Holiday from 12/15, goes last.
3. In **E5**, build the start key with `=IF(OR(NOT(ISNUMBER(A5)),NOT(ISNUMBER(B5))),"",A5*100+B5)`. December 15 becomes 1215, April 1 becomes 401.
4. On the **Dates** tab, type one date per row in **A** from row 5, as MM/DD/YYYY.
5. In **B5**, turn the date into the same kind of key with `=IF(OR(A5="",NOT(ISNUMBER(A5))),"",MONTH(A5)*100+DAY(A5))`. The year drops out: 07/04/2027 gives 704. See the MONTH help from [Microsoft](https://support.microsoft.com/en-us/office/month-function-579a2881-199b-48b2-ab90-ddba0eba86e8).
6. In **C5**, look up the season with `=IF(OR(B5="",Seasons!$I$7<>"OK"),"",IF(B5<MIN(Seasons!$E$5:$E$16),INDEX(Seasons!$C$5:$C$16,COUNT(Seasons!$E$5:$E$16)),INDEX(Seasons!$C$5:$C$16,MATCH(B5,Seasons!$E$5:$E$16,1))))`. MATCH with 1 finds the last start at or before the key. When the key is smaller than every start (01/01 to 01/05 here), the IF picks the last season in the list instead. The match type is explained in the MATCH help from [Microsoft](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [Google](https://support.google.com/docs/answer/3093378).
7. In **D5**, use the same formula on `Seasons!$D$5:$D$16` to get the rate.
8. Fill **B5:E5** down to row 504. In the template, **C** and **D** stay blank until the Seasons tab CHECK in **I7** says OK, and **E** shows a Status.

List a stay one night per row, so each night gets its own season's rate. For weekday and weekend rates, see [How to Price a Rental With Weekday and Weekend Day Rates](https://tidytabs.blogspot.com/2026/11/how-to-price-rental-with-weekday-and.html); to split a booking by month, see [How to Split a Date Range Across Months by Days](https://tidytabs.blogspot.com/2026/11/how-to-split-date-range-across-months.html).

## Example

The seasons, rates and dates are made up. Seasons: Off-season from 01/06 at $129.00 (key 106), Spring from 04/01 at $159.00 (401), Peak from 06/15 at $219.00 (615), Shoulder from 09/02 at $169.00 (902), Off-season from 11/01 at $129.00 (1101) and Holiday from 12/15 at $249.00 (1215).

| Date | Key | Season | Rate |
|---|---|---|---|
| 01/02/2026 | 102 | Holiday | $249.00 |
| 01/05/2026 | 105 | Holiday | $249.00 |
| 01/06/2026 | 106 | Off-season | $129.00 |
| 04/01/2026 | 401 | Spring | $159.00 |
| 06/14/2026 | 614 | Spring | $159.00 |
| 06/15/2026 | 615 | Peak | $219.00 |
| 09/02/2026 | 902 | Shoulder | $169.00 |
| 12/14/2026 | 1214 | Off-season | $129.00 |
| 12/15/2026 | 1215 | Holiday | $249.00 |
| 12/31/2026 | 1231 | Holiday | $249.00 |
| 01/01/2027 | 101 | Holiday | $249.00 |
| 07/04/2027 | 704 | Peak | $219.00 |

Key 102 is smaller than the first start, 106, so 01/02/2026 wraps to Holiday, the last row of the list. Key 614 sits between 401 and 615, so 06/14/2026 is still Spring. The template lists 5 more dates, such as 02/29/2028 (Off-season).

The summary in **G4:H9** shows 17 dates, $3,203.00 as the total of the rates, 0 rows showing Check, the Seasons tab CHECK OK, 0 dates typed below row 504, and CHECK OK.

## Troubleshooting

### Dates from 01/01 to 01/05 show #N/A
MATCH with 1 finds no start at or before key 101, because the first start is 106. In LibreOffice Calc, plain MATCH and LOOKUP both gave #N/A for those dates. Keep the IF that sends keys below `MIN(starts)` to the last season.

### A between-two-dates test says "not holiday" for every date
A test such as `AND(date>=DATE(YEAR(date),12,15),date<=DATE(YEAR(date),1,5))` is never true, because no date is after Dec 15 and before Jan 5 of the same year. In LibreOffice Calc it said "not holiday" for all 17 sample dates. Compare month-day keys instead.

### Dates get the wrong season
The Seasons list is not sorted, and MATCH with 1 needs sorted starts. In a copy of the template with the Seasons guard removed, an unsorted list gave Off-season for 06/14 and 07/04, Shoulder for 12/31 and #N/A for 01/06. The template itself never shows those: the start that is out of order shows Check: sort by start, and the Dates tab stays blank with Check: fix the Seasons tab until you sort the rows.

### A date shows Check: date is text
The date was typed or pasted as text, such as "12/20/2026", "July 4" or a single space. Retype it as a date.

### A season row shows Check: day not valid for the month
The start day does not exist in that month, such as 4/31. Use 4/30 or 5/1. February 29 works as a start: 02/29/2028 then gets that season, while 02/28/2027 stays in the season before it.

## Template

[Download the .xlsx](templates/tidy-tabs-season-from-date.xlsx)

The file has three tabs. **Seasons** holds up to 12 season starts in rows 5 to 16, with Start month, Start day, Season name and Rate $ in **A:D**, the key in **E** and a Status in **F** that flags bad months and days, missing names or rates, repeated starts and a list that is not sorted. **H4:I7** gives the CHECK.

**Dates** holds up to 500 dates in rows 5 to 504. You type **A**; **B:E** hold the key, season, rate and Status, and **G4:H9** sums the rates and checks the tab. **How to use** explains each formula. Both data tabs print Landscape on US Letter.

Yellow cells are inputs, gray cells are formulas, and every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
