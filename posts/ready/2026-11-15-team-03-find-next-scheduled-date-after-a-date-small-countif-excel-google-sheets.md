---
title: How to Find the Next Date on a Schedule in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 25 workbooks and 5,650 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Pull the next 5 markets or deadlines from an unsorted list with SMALL and COUNTIF, show days until each, and keep two same-day events apart. Free .xlsx template.
image_prompts: A collapsible garden wagon parked by a small shop's back door at dawn, loaded with a folded market table, a rolled canvas tablecloth and a folding stool, soft early light, realistic photo, no screens, no readable text, no logos
image_alt: Garden wagon loaded with a folded market table, rolled tablecloth and stool
threads: Which market is next when your schedule isn't sorted?\n=SMALL(A5:A104,COUNTIF(A5:A104,"<"&B1)+1) finds it, and +2, +3 give the ones after.\nIf B1 holds a time like 9:00 AM, wrap it in INT or today's market gets skipped.
---
Return the next date on or after the date in **B1** with `=SMALL(A5:A104,COUNTIF(A5:A104,"<"&B1)+1)`, and change the +1 to +2, +3 and so on for the dates after it, even when the list is not sorted. In the made-up sample with an as-of date of 11/14/2026, the next markets are the Farmers market on 11/14/2026 (0 days), the Holiday market on 11/21/2026 (7 days) and two events on 11/28/2026; the download adds a helper column so both same-day events show by name, and INT so an as-of date with a time does not skip today.

COUNTIF counts the dates before your as-of date, and SMALL takes the one right after them. That is why the list can stay in the order you typed it.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (25 workbooks, 5,650 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither was opened for this post; COUNTIF with `"<"&INT(B1)`, SMALL on a helper column with decimals and how a text date is read are the open questions there.

![Unsorted market dates marked Past or Upcoming beside the next five events and days until](images/2026-11-15-team-03-find-next-scheduled-date-after-a-date-small-countif-excel-google-sheets/next-date-on-schedule-template.png) *As of 11/14/2026, an unsorted list of seven market dates gives the next five in order, the two on 11/28 each with its own name (ZIP, helper and summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has two tabs: **Schedule** and **How to use**.

1. On the **Schedule** tab, type the as-of date in **B1**. The sample uses 11/14/2026. Type a real date, not text.
2. List your events in **A5:C104**: date in A, event name in B and location ZIP in C. Any order works.
3. In **D5**, build a sort key that keeps two events on one day apart: `=IF(ISNUMBER(A5),INT(A5)+ROW()/100000,"")`. The 11/28/2026 event on row 5 gets a slightly smaller key than the one on row 10.
4. In **G5:G9**, type 1 to 5. These say which upcoming date each row shows.
5. In **H5**, find the next upcoming key: `=IFERROR(SMALL($D$5:$D$104,COUNTIF($D$5:$D$104,"<"&INT($B$1))+G5),"")`.
6. In **I5**, show the date: `=IF(H5="","None scheduled",INT(H5))`.
7. In **J5**, pull the event name with the key, not the date: `=IF(H5="","",INDEX($B$5:$B$104,MATCH(H5,$D$5:$D$104,0)))`.
8. In **K5**, count the days until it: `=IF(H5="","",I5-INT($B$1))`.
9. Fill **H5:K5** down to row 9.
10. For the last event before the as-of date, use `=IFERROR(LARGE($A$5:$A$104,COUNTIF($A$5:$A$104,">="&INT($B$1))+1),"None")`.

The `"<"` in step 5 includes events on the as-of date itself. With `"<="` you get only later dates: in the sample, 11/21/2026 instead of 11/14/2026.

The template wraps these formulas so a blank or text as-of date shows None scheduled and marks every Status as Check, and the event name stays blank when the name cell is blank.

The template never uses TODAY, so its results stay fixed and you can check them against this post. In your own copy you can type `=TODAY()` in **B1**; the results then change every day.

A next date per market name (the next Farmers market only, say) needs a separate list for that market. This file does not do it.

Help pages: Microsoft's [SMALL function](https://support.microsoft.com/en-us/office/small-function-17da8222-7c82-42b2-961b-14c45384df07) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34), and Google's [SMALL](https://support.google.com/docs/answer/3094050) and [COUNTIF](https://support.google.com/docs/answer/3093480). They describe the functions, not our test.

## Example

Sample data is fictional: 7 made-up markets and events, typed in no order, with an as-of date of 11/14/2026.

| Row | Date | Event | Location |
|---|---|---|---|
| 5 | 11/28/2026 | Small Biz Saturday | 02108 |
| 6 | 11/07/2026 | Fall craft show | 07030 |
| 7 | 12/05/2026 | Craft fair | 01060 |
| 8 | 11/14/2026 | Farmers market | 02108 |
| 9 | 10/24/2026 | Harvest market | 06510 |
| 10 | 11/28/2026 | Pop-up | 07030 |
| 11 | 11/21/2026 | Holiday market | 02139 |

The next block in **G4:K9** lists them in date order:

| No. | Date | Event | Days until |
|---|---|---|---|
| 1 | 11/14/2026 | Farmers market | 0 |
| 2 | 11/21/2026 | Holiday market | 7 |
| 3 | 11/28/2026 | Small Biz Saturday | 14 |
| 4 | 11/28/2026 | Pop-up | 14 |
| 5 | 12/05/2026 | Craft fair | 21 |

Both 11/28/2026 events show by name, in row order. The Farmers market date is green because it is 0 days away.

The summary in **M4:N9** shows Last one before 11/07/2026 (Fall craft show), 5 events this month, 5 events on or after the as-of date, 0 text dates and 0 rows showing Check. Status in column E reads Upcoming for 5 rows and Past for 2.

Other as-of dates from the tests: 11/14/2026 9:00 AM gives the same list. As of 12/28/2026, a schedule with 12/31/2026, 01/02/2027 and 01/09/2027 shows 3, 5 and 12 days, so the year change is fine. As of 01/15/2027 with the sample, all five slots show None scheduled and the last one before is the 12/05/2026 Craft fair.

## Troubleshooting

### Two events on one day show the same name
INDEX and MATCH on the date alone find the first match. In LibreOffice, both 11/28/2026 rows returned Small Biz Saturday. Match on the sort key in column D instead, which is different for every row.

### Today's market is skipped
The as-of date has a time. With **B1** as 11/14/2026 9:00 AM and no INT, the formula returned 11/21/2026, because the 11/14 market (midnight) counts as before 9:00 AM. Wrap the as-of date in INT, as the template does.

### A late-evening event lands on the next day
A key of `A5+ROW()/100000` on an 11:59:59 PM event on row 104 rounds past midnight into 11/15/2026. Take `INT(A5)` first, then add the row part. The template's key keeps it on 11/14/2026. The row part works up to 99,999 rows; the file has 100.

### An event never shows
Its date is stored as text, not a real date. SMALL and COUNTIF skip it. The template counts it under Text dates (not counted) and marks its Status as Check; retype it as a date.

### Every slot says None scheduled
**B1** is blank or holds text. Every Status shows Check too. Type a real date in **B1**.

## Template

[Download the .xlsx](templates/tidy-tabs-next-date-on-schedule.xlsx)

**Schedule** has the as-of date in **B1**, inputs in **A:C** (Date, Event, Location as ZIP) for rows 5 to 104, the Sort key in **D**, Status in **E** (Past, Upcoming, Check), the next five in **G4:K9** (No., Key, Date, Event, Days until) and the summary in **M4:N9**. **How to use** explains each part. It prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Upcoming is green, Check is red, and a date 0 days away is green. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
