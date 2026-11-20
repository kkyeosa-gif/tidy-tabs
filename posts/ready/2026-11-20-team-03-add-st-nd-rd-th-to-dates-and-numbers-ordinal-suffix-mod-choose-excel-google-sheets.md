---
title: How to Add st, nd, rd and th to Dates and Numbers in Excel and Google Sheets
labels: excel-formulas, dates, text-functions, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 36 workbooks and 21,888 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Add 1st, 2nd, 3rd and th to numbers and dates with MOD and CHOOSE, with 11th, 12th and 13th handled. Get "Friday, November 20th, 2026". Free template.
image_prompts: A small bakery pickup counter with a stack of white pastry boxes tied with red and white baker's twine, a little wooden A-frame chalkboard sign with smudged illegible chalk writing, and a metal clipboard of order slips lying beside the boxes, morning light, realistic photo, no screens, no readable text, no logos
image_alt: Twine-tied bakery boxes stacked on a pickup counter beside a small chalkboard sign
threads: 11st, 12nd and 13rd is what you get when a formula only checks the last digit.\nTest MOD(n,100) for 11 to 13 first, then let CHOOSE(MOD(n,10)+1,...) pick the rest.\nJoin DAY() to TEXT() for "Friday, November 20th, 2026".
---
Build the suffix with a formula: `=n&IF(AND(MOD(n,100)>=11,MOD(n,100)<=13),"th",CHOOSE(MOD(n,10)+1,"th","st","nd","rd","th","th","th","th","th","th"))`, where the 11 to 13 test keeps 11th, 12th and 13th from turning into 11st, 12nd and 13rd. For a date, use `DAY(date)` as n and join it to `TEXT(date,"dddd, mmmm ")` and `TEXT(date,", yyyy")` to get "Friday, November 20th, 2026", and the download marks text, 0, negatives and decimals as Check.

The result is text for a market sign, pickup reminder or thank-you card. Keep the real date beside it for sorting and math.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (36 workbooks, 21,888 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (TEXT, DAY, MOD, CHOOSE, IF, AND) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Twelve pickup dates shown in long, short and day-only formats with ordinal endings](images/2026-11-20-team-03-add-st-nd-rd-th-to-dates-and-numbers-ordinal-suffix-mod-choose-excel-google-sheets/ordinal-date-suffix-template.png) *Twelve sample dates with the sign text "Pickup on": 1st, 2nd and 3rd, 11th, 12th and 13th (not 11st, 12nd, 13rd), 21st to 23rd and Thursday, December 31st (the Status column and the check summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Both template tabs start their data in row 5.

### Add the suffix to a number

1. On the **Numbers** tab, type a whole number of 1 or more in **A5** and down.
2. In **B5**, type `=A5&IF(AND(MOD(A5,100)>=11,MOD(A5,100)<=13),"th",CHOOSE(MOD(A5,10)+1,"th","st","nd","rd","th","th","th","th","th","th"))`.
3. Fill **B5** down. 1 becomes 1st, 22 becomes 22nd, and 113 becomes 113th.

`MOD(A5,10)` is the last digit. CHOOSE picks by position, so adding 1 maps 0 to th, 1 to st, 2 to nd, 3 to rd and 4 to 9 to th. See Microsoft's [CHOOSE function](https://support.microsoft.com/en-us/office/choose-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc) page.

The AND test runs first. `MOD(A5,100)` is the last two digits, so 11, 12, 13, 111, 112 and 113 all get th. The template's **B5** also checks the Status in **C** and stays blank for a bad entry.

### Add the suffix to a date

1. On the **Dates** tab, type real dates in **A5** and down as MM/DD/YYYY, such as 11/20/2026.
2. For the day alone, type `=DAY(A5)&IF(AND(DAY(A5)>=11,DAY(A5)<=13),"th",CHOOSE(MOD(DAY(A5),10)+1,"th","st","nd","rd","th","th","th","th","th","th"))` in **D5**. For 11/20/2026 it gives 20th. A day never passes 31, so DAY itself is tested for 11 to 13.
3. For the long form, type `=TEXT(A5,"dddd, mmmm ")&DAY(A5)&IF(AND(DAY(A5)>=11,DAY(A5)<=13),"th",CHOOSE(MOD(DAY(A5),10)+1,"th","st","nd","rd","th","th","th","th","th","th"))&TEXT(A5,", yyyy")` in **B5**. For 11/20/2026 it gives "Friday, November 20th, 2026".
4. For the short form, type `=TEXT(A5,"mmm ")&D5` in **C5** (Nov 20th).
5. For a sign line, type your lead-in text in **B1** (the sample says "Pickup on"). **E5** puts it in front of the weekday, month and day; for 11/20/2026 that reads "Pickup on Friday, November 20th". TRIM drops the extra space when **B1** is blank or has a double space.
6. Fill down to row 104.

The format codes "dddd" (full weekday), "mmmm" (full month) and "mmm" (short month) are described on Microsoft's [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c) page and Google's [TEXT function](https://support.google.com/docs/answer/3094139) page. In LibreOffice, `TEXT(A5,"mmmm d")` gave "November 20" with no suffix, so the day comes from DAY.

A time part is ignored: in LibreOffice, 11/20/2026 6:45 PM still gave "Friday, November 20th, 2026".

Day and month names follow the app's language setting. The template was checked in English only; other languages were not tested.

To work out which date the next market falls on, see [finding the 2nd Saturday of a month](https://tidytabs.blogspot.com/2026/10/how-to-find-date-of-2nd-saturday-or-any.html).

## Example

The market and pickup dates are fictional.

| Date | Long | Short |
|---|---|---|
| 11/01/2026 | Sunday, November 1st, 2026 | Nov 1st |
| 11/02/2026 | Monday, November 2nd, 2026 | Nov 2nd |
| 11/03/2026 | Tuesday, November 3rd, 2026 | Nov 3rd |
| 11/11/2026 | Wednesday, November 11th, 2026 | Nov 11th |
| 11/12/2026 | Thursday, November 12th, 2026 | Nov 12th |
| 11/13/2026 | Friday, November 13th, 2026 | Nov 13th |
| 11/20/2026 | Friday, November 20th, 2026 | Nov 20th |
| 11/21/2026 | Saturday, November 21st, 2026 | Nov 21st |
| 11/22/2026 | Sunday, November 22nd, 2026 | Nov 22nd |
| 11/23/2026 | Monday, November 23rd, 2026 | Nov 23rd |
| 12/31/2026 | Thursday, December 31st, 2026 | Dec 31st |
| 01/01/2027 | Friday, January 1st, 2027 | Jan 1st |

With **B1** set to "Pickup on", **E** reads "Pickup on Saturday, November 21st" for 11/21/2026. The summary in **H4:I7** shows Dates listed 12, Rows showing Check 0, Typed below row 104 0 and CHECK OK.

The **Numbers** tab shows 1st, 2nd, 3rd, 4th, 11th, 12th, 13th, 21st, 22nd, 23rd, 101st, 111th, 112th and 113th. Every day of five months (including Tuesday, February 29th, 2028) and the numbers 1 to 1000 matched the Python model.

## Troubleshooting

### The sign says "Market on 46346"
The date was joined straight to text, as in `="Market on "&A1` with 11/20/2026 in **A1**. Joining shows the date's serial number. Wrap the date in TEXT and add the suffix from DAY, as column **E** does.

### The number shows 0th, -1th or 2.5nd
In LibreOffice, the formula without a guard gave 0th, -1th and 2.5nd. The template shows Check: 0 or less or Check: decimal in **C** and leaves **B** blank.

### A number typed as text still got a suffix
The text "21" gave 21st in the unguarded LibreOffice formula, because MOD turned it into a number. The template shows Check: text; retype it as a plain number. Excel and Google Sheets were not tested on this.

### Status shows Check: not a date
The date in **A** is text, such as "11/20/2026" pasted as text, or a number like 0, -5 or 0.5. Retype it as a real date in MM/DD/YYYY. If dates came in from a CSV in the wrong order, see [fixing dates that flip from MM/DD to DD/MM](https://tidytabs.blogspot.com/2026/09/how-to-fix-dates-that-flip-from-mmdd-to.html).

### CHECK shows Check with no red rows
The tab is empty, or something was typed on row 105 (Dates) or 55 (Numbers), where no formula reads it. Move it up into the yellow rows.

## Template

[Download the .xlsx](templates/tidy-tabs-ordinal-dates.xlsx)

**Dates** has the sign text in **B1** and rows 5 to 104 with Date (input), Long, Short, Day only, Sign line and Status (OK or Check: not a date), plus a summary in **H4:I7**: Dates listed, Rows showing Check, Typed below row 104 and CHECK. **Numbers** has rows 5 to 54 with Number, With suffix and Status (OK, Check: text, Check: 0 or less, Check: decimal), and the same summary in **E4:F7**. **How to use** explains each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Check statuses and a CHECK that is not OK turn red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
