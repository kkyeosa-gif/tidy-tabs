---
threads_url: https://www.threads.com/@tin_ylab/post/Dfht0-eIJRg
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-convert-text-like-2h-30m-into.html
title: How to Convert Text Like 2h 30m Into Minutes in Excel and Google Sheets
labels: excel-formulas, small-business, time-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 11 workbooks and 800 rows. Checked: the sample, an empty sheet, 48 typed entries (capital H and M, no spaces, extra spaces, 1.5h, 90m, 30 min, blank, n/a, 2:30, 5, 2h 30, 0m, -30m, h30m, abc and more), two random sets of 100 rows, numbers typed into the cells, a rate of 0, a text rate, a blank rate, tab characters and non-breaking spaces. Excel and Google Sheets were not used. Not run: Excel's handling of a non-breaking space, a non-US language setting, more than 100 entries.
search_description: Turn typed durations like 2h 30m or 45m into minutes and decimal hours with LEFT, MID and FIND in Excel and Google Sheets. Free template.
image_prompts: A spiral notebook with handwritten time scribbles blurred out, a small analog kitchen timer and a pencil on a freelancer's wooden desk, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Kitchen timer, pencil and notebook on a wooden desk
threads: Clients type "2h 30m", "45m" and "1.5h" in the same column?\nStrip the spaces, read the number before the h, then =ROUND(D4*60+E4,0).\nIn the sample, 2h 30m is 150 minutes and 3H 5M is 185, which bills $200.42 at $65.00 (made-up entries).
---
Strip the spaces and lowercase the entry with `=LOWER(SUBSTITUTE(TRIM(B4)," ",""))`, take the hours before the h with `=IF(ISNUMBER(FIND("h",C4)),VALUE(LEFT(C4,FIND("h",C4)-1)),0)`, and get total minutes with `=ROUND(D4*60+E4,0)`. The entry "2h 30m" becomes 150 minutes, and "45m" becomes 45 (E4 reads the number before the m, shown below).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (11 workbooks, 800 rows). The formulas use only IF, ISNUMBER, FIND, VALUE, LEFT, MID, LOWER, SUBSTITUTE, TRIM and ROUND, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Table of typed durations with cleaned text, minutes, decimal hours and amounts for twelve sample entries](images/2026-11-09-team-01-convert-text-hours-minutes-1h-30m-to-minutes-find-mid-excel-google-sheets/text-duration-to-minutes-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. Put the typed entries in column **B** from row 4, with headers in row 3. Type your hourly rate in **B1**.

1. In **C4**, type `=LOWER(SUBSTITUTE(TRIM(B4)," ",""))`. This removes every space and makes the letters lowercase, so " 3H 5M " becomes `3h5m`.
2. In **D4**, type `=IF(ISNUMBER(FIND("h",C4)),VALUE(LEFT(C4,FIND("h",C4)-1)),0)`. This reads everything before the h as the hours, or gives 0 when there is no h.
3. In **E4**, type the minutes formula:

   `=IF(ISNUMBER(FIND("m",C4)),VALUE(MID(C4,IF(ISNUMBER(FIND("h",C4)),FIND("h",C4)+1,1),FIND("m",C4)-IF(ISNUMBER(FIND("h",C4)),FIND("h",C4)+1,1))),0)`

   It starts reading just after the h. When there is no h, as in `45m`, it starts at character 1. It stops at the m.
4. In **F4**, type `=ROUND(D4*60+E4,0)`. The ROUND turns a decimal such as 1.5h into whole minutes.
5. In **G4**, type `=F4/60` for decimal hours, and in **H4** type `=ROUND(G4*$B$1,2)` for the dollar amount.
6. Fill **C4:H4** down as far as your entries go.

Type entries as text. If a cell turns `1h` into something else, format column **B** as Plain text first.

The template wraps the core in checks. Total minutes shows Check, instead of a wrong number, when an entry cannot be read, has a negative part or ends in a digit with no unit, such as `2h 30`. The totals skip those rows.

Help pages: Microsoft's [FIND function](https://support.microsoft.com/en-us/office/find-find-b-function-c7912941-af2a-4bdf-a553-d0d89b0a0628) and [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332), and Google's [SUBSTITUTE](https://support.google.com/docs/answer/3094215). They describe the functions, not our test.

## Example

Sample data is fictional: 12 rows with dates from 11/02/2026 to 11/13/2026, one blank, and a rate of $65.00.

| Typed | Minutes | Decimal hours | Amount at $65.00 |
|---|---|---|---|
| 2h 30m | 150 | 2.50 | $162.50 |
| 45m | 45 | 0.75 | $48.75 |
| 1h | 60 | 1.00 | $65.00 |
| 1h15m | 75 | 1.25 | $81.25 |
| 3H 5M | 185 | 3.08 | $200.42 |
| " 90m " | 90 | 1.50 | $97.50 |
| 2h 0m | 120 | 2.00 | $130.00 |
| 1.5h | 90 | 1.50 | $97.50 |
| (blank) | blank | blank | blank |
| 30 min | 30 | 0.50 | $32.50 |
| n/a | Check | 0.00 | $0.00 |
| 0h 20m | 20 | 0.33 | $21.67 |
| Total | 865 | 14.42 | $937.09 |

For "2h 30m", C4 is `2h30m`, D4 is 2 and E4 is 30, so F4 is 2 times 60 plus 30, which is 150. For "45m", D4 is 0 because there is no h, and E4 reads the 45 from character 1 up to the m.

Each amount is rounded to cents before it is added, so $937.09 is the sum of the rounded rows. The 11 entries add up to 865 minutes, or 14.42 hours. The "n/a" row shows Check and is not counted.

## Troubleshooting

### Hours or Minutes shows #VALUE!
The entry has an order or wording the formulas cannot read. In the test, `h30m`, `30m2h`, `1 hour 30 minutes`, `M`, `m30`, `1.5.5h` and `1,5h` all did this. Total minutes and Check say Check on those rows, so the #VALUE! is the sheet pointing at the entry. Retype it as `1h30m`.

### 2:30, n/a or abc gives 0 minutes
There is no h or m for the formulas to find, so both parts are 0. The Check column flags a row that comes to 0 minutes. A colon time such as `2:30` is not supported, so retype it as `2h 30m`.

### "2h 30" or a plain 5 is not counted
An entry that ends in a digit has no unit for the last number. The template shows Check instead of quietly using 0 or dropping the 30. Add the m. A cell holding only a number such as 5 is treated the same way.

### Some words are read and some are not
In the test, `2 hours` gave 120, `30mm` gave 30 and `1hh` gave 60, because text after the h or m is ignored when nothing else follows. `1 hour 30 minutes` was not read and shows Check. Do not count on words working. Stick to h and m.

### A pasted entry still shows Check
The paste may carry a non-breaking space, which looks like a space but is a different character. In LibreOffice, entries with one were read correctly (150, 150 and 60 for `2h 30m` variants) because its VALUE ignores it. Excel was not run, so that is unconfirmed there. If a paste shows Check, wrap B4 as `SUBSTITUTE(B4,UNICHAR(160)," ")` in the formula in C4 (Excel 2013 and later, Google Sheets). A tab character is not removed and shows Check.

## Template

[Download the .xlsx](templates/tidy-tabs-text-duration-to-minutes.xlsx)

The file has two tabs. **Entries** has your hourly rate in **B1**, 100 slots from row 4 to row 103, and 12 fictional sample entries. Columns **A** to **I** hold the date, the entry as typed, Cleaned, Hours, Minutes, Total minutes, Decimal hours, Amount and Check. A totals block at the right counts entries, total minutes, total decimal hours, the amount, entries showing Check, and whether the rate is a number. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. Hours and Minutes show 0 as faint text on blank rows. The date column is optional and no formula uses it. The page prints on US Letter, landscape.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
