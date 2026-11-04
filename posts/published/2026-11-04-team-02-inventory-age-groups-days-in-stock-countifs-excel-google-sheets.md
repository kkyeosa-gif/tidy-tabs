---
threads_url: https://www.threads.com/@tin_ylab/post/DfU124YIJsI
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-group-inventory-by-days-in-stock.html
title: How to Group Inventory by Days in Stock in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 16 workbooks and 5,328 comparisons with 0 mismatches. Checked: an empty list, items at 0, 1, 30, 31, 60, 61, 90, 91, 180, 181, 365 and 700 days, a received date after the as-of date, a date typed as text, an as-of date typed as text, an as-of date of 02/28/2027, custom group starts, group starts that do not rise, text units, a blank cost, half-cent values, 100 items, and four random lists. Excel and Google Sheets were not used.
search_description: Group stock by days on the shelf in Excel and Google Sheets with COUNTIFS and SUMIFS. Free template shows items, units and value per age group.
image_prompts: A wooden storage shelf with a few dusty cardboard boxes pushed to the back and a neat row of fresh boxes at the front, strips of plain masking tape on each box, soft daylight, no screens, no readable text, no logos
image_alt: Storage shelf with old dusty boxes at the back and fresh boxes in front
threads: Which stock has been sitting the longest?\nSubtract the date received from an as-of date, then COUNTIFS(days,">="&31,days,"<"&61) counts the 31 to 60 day group. SUMIFS gives units and dollars for the same group.\nIn the sample shelf, 2 items worth $93.60 are 181 days or older.
---
Subtract the date received from an as-of date with `=$B$2-B6`, then count the items in each age group with `=COUNTIFS(Stock!$E$6:$E$105,">="&A5,Stock!$E$6:$E$105,"<"&A6)`. Use SUMIFS the same way for units and stock value.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIFS, SUMIFS, INDEX and MATCH exist in current Excel and Google Sheets, but neither app was opened for this post.

![Summary table of five age groups with items, units, stock value and share of value](images/2026-11-04-team-02-inventory-age-groups-days-in-stock-countifs-excel-google-sheets/inventory-age-groups-by-days-in-stock-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Stock** has the as-of date in **B2** and one item per row from row 6: item in **A**, date received in **B**, units on hand in **C**, unit cost in **D**. **Summary** has the age groups in rows 5 to 9.

1. In **Stock B2**, type the as-of date, such as 11/04/2026.
2. In **E6**, type `=$B$2-B6`. This is the days in stock. Fill it down.
3. In **G6**, type `=ROUND(C6*D6,2)` for the stock value of that row. Fill it down.
4. On **Summary**, type the first day of each group in **A5:A9**: 0, 31, 61, 91 and 181.
5. In **Summary C5**, type `=COUNTIFS(Stock!$E$6:$E$105,">="&A5,Stock!$E$6:$E$105,"<"&A6)`. This counts items from day 0 up to, but not including, day 31. Fill it down to row 8.
6. In **C9**, the last group has no upper limit: `=COUNTIFS(Stock!$E$6:$E$105,">="&A9)`.
7. In **D5**, type `=SUMIFS(Stock!$C$6:$C$105,Stock!$E$6:$E$105,">="&A5,Stock!$E$6:$E$105,"<"&A6)` for units. In **E5**, use the same formula with `Stock!$G$6:$G$105` for the value. Fill both down and change row 9 the same way as C9.

To show each item's group on the Stock tab, put `=INDEX(Summary!$B$5:$B$9,MATCH(E6,Summary!$A$5:$A$9,1))` in **F6**. MATCH with 1 finds the last group start that is not larger than the days in stock.

The as-of date is typed, so the groups stay put when you reopen the file. Type `=TODAY()` in **B2** if you want them to move each day.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842), [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/03/2026; the pages were not read for this post.

## Example

Sample data is fictional: a gift shop with 16 items, as of 11/04/2026.

| Age group | Items | Units | Stock value | Share of value |
|---|---|---|---|---|
| 0 to 30 days | 5 | 254 | $438.70 | 41.7% |
| 31 to 60 days | 2 | 115 | $135.25 | 12.8% |
| 61 to 90 days | 2 | 67 | $135.95 | 12.9% |
| 91 to 180 days | 5 | 120 | $249.70 | 23.7% |
| 181 days or more | 2 | 15 | $93.60 | 8.9% |

The oldest item is the linen napkin set, received 01/12/2026 and 296 days in stock. The two items over 180 days are worth $93.60 at cost, which is 8.9% of the $1,053.20 total.

This shows how long stock has sat. It does not say whether it is selling, and the value is cost, not what you could sell it for.

## Troubleshooting

### An item at exactly 30 or 31 days lands in the wrong group
The groups use "from this day up to, but not including, the next start". In the test, an item at 30 days was in 0 to 30, 31 in 31 to 60, and the same at 60/61, 90/91 and 180/181.

### A row shows Check
The date received is blank, typed as text, or later than the as-of date. The units or cost may also be missing. In the test, a future date and a text date each showed Check on that row, and the **CHECK** cell said Check.

### Every row shows Check
The as-of date in **B2** is text. In the test, an as-of date typed as text turned every item to Check. Retype the date so it is right-aligned.

### CHECK says Check but no row shows Check
The group starts must begin at 0 and rise. In the test, typing 20 in **A7** under a start of 31 made **I9** say Check.

### A restocked item shows one date
The sheet uses one received date per row. List each lot on its own row, such as the same candle received in September and again in October.

## Template

[Download the .xlsx](templates/tidy-tabs-inventory-age-buckets.xlsx)

The file has three tabs. **Stock** has the as-of date in **B2** and 100 rows from row 6 (yellow inputs in **A:D**, formulas in **E:G**). **Summary** has the five yellow group starts in **A5:A9**, the group table in **A4:F10** and a summary in **H5:I9** with the oldest item and a **CHECK** cell. **How to use** has notes on each formula.

**CHECK** says OK when no row shows Check, the starts rise from 0, every item is in a group and the group values add up to the stock values.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
