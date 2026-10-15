---
threads_url: https://www.threads.com/@tin_ylab/post/DehPFoclFrW
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-count-market-saturdays-in-month.html
title: How to Count Market Saturdays in a Month With NETWORKDAYS.INTL in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the Saturday, Sunday and Friday counts for all 12 months of 2026 matched Python's calendar. Edge cases run on a copy: a month starting on a Saturday, a reversed date range, a one day range, a different picked day and a date typed as text. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Help.
search_description: Count Saturdays in a month: =NETWORKDAYS.INTL(A2,EOMONTH(A2,0),"1111101"). October 2026 has 5. Free .xlsx with all 12 months and booth cost.
image_prompts: Wall calendar page with several Saturday squares marked by small round stickers, a folded canopy tent bag and a cash box on a garage floor, soft daylight, no screens, no readable text, no logos
image_alt: Wall calendar with stickers on Saturdays beside a folded tent bag and cash box
threads: Booth fee is per Saturday and your market runs every one. How many is that this month?\n=NETWORKDAYS.INTL(A2,EOMONTH(A2,0),"1111101") counts them. October 2026 has 5.\nThe seven digits are Monday to Sunday, 0 means count that day. Change one digit for any other weekday.
---
Count one weekday in a month with `=NETWORKDAYS.INTL(A2,EOMONTH(A2,0),"1111101")`, where A2 is the first day of the month. The seven digits stand for Monday to Sunday, and the single 0 is the day to count. For 10/01/2026 it returns 5 Saturdays.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every count below was recalculated and checked against Python's calendar. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Twelve months of 2026 with first day, last day, Saturday and Sunday counts and booth cost per month](images/2026-10-14-team-05-count-market-saturdays-in-month-networkdays-intl-excel-google-sheets/count-market-saturdays-by-month-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

NETWORKDAYS.INTL normally counts workdays. The third argument is a string that says which weekdays to skip. Mark every day with 1 except the one you want, and it counts only that day.

1. Type the first day of a month in **A2**, for example 10/01/2026. Use a real date, not text.
2. In **B2**, type `=EOMONTH(A2,0)`. This returns the last day of that month.
3. In **C2**, type `=NETWORKDAYS.INTL(A2,B2,"1111101")` for Saturdays. The sixth digit is 0, so only Saturday counts.
4. In **D2**, type `=NETWORKDAYS.INTL(A2,B2,"1111110")` for Sundays. The seventh digit is 0.
5. Fill rows 2 to 13 with the first day of each month.
6. In **I2**, type your booth fee per market day, such as 25. In **F2**, type `=C2*$I$2` for the booth cost of the month.

The string always runs Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday. For Fridays use `"1111011"`. For Mondays use `"0111111"`.

The start date counts too. If the month starts on a Saturday, that day is included.

To pick the day from a dropdown, type Mon to Sun in **K2:K8**, put the choice in **I3**, and build the string in **I4** with `=REPT("1",MATCH(I3,$K$2:$K$8,0)-1)&"0"&REPT("1",7-MATCH(I3,$K$2:$K$8,0))`. Then use `$I$4` as the third argument.

### In Excel

The formulas work as typed. Microsoft describes the weekend string on its [NETWORKDAYS.INTL function](https://support.microsoft.com/en-us/office/networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28) page and the month end on its [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628) page. Format dates from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Google's [NETWORKDAYS help page](https://support.google.com/docs/answer/3092979) covers the simpler sibling function, and the weekend string works the same way as in Microsoft's description above. Format dates with **Format > Number > Date**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. The booth fee is $25.00 per market day and the market runs every Saturday.

| First day | Last day | Saturdays | Sundays | Booth cost (Saturdays) |
|---|---|---|---|---|
| 01/01/2026 | 01/31/2026 | 5 | 4 | $125.00 |
| 02/01/2026 | 02/28/2026 | 4 | 4 | $100.00 |
| 08/01/2026 | 08/31/2026 | 5 | 5 | $125.00 |
| 09/01/2026 | 09/30/2026 | 4 | 4 | $100.00 |
| 10/01/2026 | 10/31/2026 | 5 | 4 | $125.00 |
| 11/01/2026 | 11/30/2026 | 4 | 5 | $100.00 |

For all 12 months of 2026, Saturdays come to 5, 4, 4, 4, 5, 4, 4, 5, 4, 5, 4, 4, which is 52. The booth cost for the year is 52 x $25.00 = $1,300.00. In **I9**, `=IF(COUNT(A2:A25)=0,"",IF(I6=NETWORKDAYS.INTL(MIN(A2:A25),EOMONTH(MAX(A2:A25),0),"1111101"),"OK","CHECK"))` compares the twelve monthly counts with one count over the whole year and shows OK.

August 2026 starts on a Saturday, so it has 5 Saturdays (08/01, 08/08, 08/15, 08/22, 08/29).

This counts calendar days only. It does not know your market's holidays, rain dates or the days it actually opens. Your market sets the real fees and dates.

## Troubleshooting

### The count is negative
The start date is after the end date. In the recalculated file, `=NETWORKDAYS.INTL(DATE(2026,10,31),DATE(2026,10,1),"1111101")` returned -5. Swap the two dates.

### It counts the wrong day
The string is read left to right as Monday to Sunday, not Sunday to Saturday. `"1111101"` counts Saturdays and `"0111111"` counts Mondays. Count the position of the 0 to check.

### It returns 0 or an error
A date typed as text can break it. In the recalculated file, LibreOffice converted the text "01/01/2026" and still returned 5, but that depends on the app and settings, so type real dates. Check that the string has exactly seven digits.

### The month total is off by one
Both the first and last dates count. A one day range on Saturday 10/03/2026 returned 1, and on Sunday 10/04/2026 returned 0. If your range should stop the day before, subtract a day from the end date.

### I need holidays removed
This template does not remove holidays. NETWORKDAYS.INTL has an optional fourth argument for a holiday list, but it was not tested here.

## Template

[Download the .xlsx](templates/tidy-tabs-count-weekdays-in-month.xlsx)

The file has two tabs. **Months** holds the 12 months of 2026 in A2:A13, formulas ready through row 25, a booth fee cell in I2, a day dropdown in I3 and the totals in I6:I9. **How to use** has short fill-in notes.

The dates and fees are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
