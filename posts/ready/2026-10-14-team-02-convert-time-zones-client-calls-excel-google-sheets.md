---
title: How to Convert Time Zones for Client Calls in Excel and Google Sheets
labels: excel-formulas, small-business, dates-and-times
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and 8 calls were checked against Python datetime math, plus a call that crosses midnight in both directions, a 0 hour difference, a blank time, a time typed as text and an unknown zone. Excel and Google Sheets were not opened; the links are Microsoft Support, Google Docs Help and NIST.
search_description: Convert a client call time with =C2+D2/24, where D2 is the hour difference. 4:30 PM Eastern is 1:30 PM Pacific. Free .xlsx with a zone dropdown.
image_prompts: Wooden desk with a small analog wall clock propped beside a paper desk calendar with a pencil resting on it, soft window light, no screens, no readable text, no logos
image_alt: Analog clock leaning against a paper desk calendar with a pencil
threads: Your 4:30 PM Eastern call is 1:30 PM for the client in Pacific.\nAdd the hours divided by 24: =C2+D2/24, where D2 is -3. It moves the date too, so a 1:00 AM call becomes 10:00 PM the night before.\nThe template flags Next day and Previous day so nobody dials in a day early.
---
Add the hour difference divided by 24 with `=C2+D2/24`, where C2 holds the call date and time and D2 holds the difference, such as -3 from Eastern to Pacific. A call at 10/14/2026 4:30 PM Eastern shows as 10/14/2026 1:30 PM Pacific.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Client call list with client zone, hours difference, client time and a day column that flags calls landing on the previous day](images/2026-10-14-team-02-convert-time-zones-client-calls-excel-google-sheets/time-zone-converter-client-calls-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

Excel and Sheets store a date and time as one number. The whole part is the day and the fraction is the time. One hour is 1/24 of a day, so adding hours divided by 24 moves the date and the time together.

1. On a tab named **Offsets**, type the headers Zone and Hours from Eastern in row 1.
2. In **A2:B5**, type Eastern 0, Central -1, Mountain -2 and Pacific -3.
3. On a tab named **Calls**, type these headers in row 1: Client, Client zone, Your call time, Hours difference, Client time, Day.
4. In **I2**, type your own zone, for example Eastern.
5. Type each client in column **A**, their zone in column **B**, and your call date and time in column **C**, such as 10/14/2026 4:30 PM.
6. In **D2**, type `=IF(B2="","",IFERROR(VLOOKUP(B2,Offsets!$A$2:$B$5,2,FALSE)-VLOOKUP($I$2,Offsets!$A$2:$B$5,2,FALSE),"Check zone"))` and fill it down. This is the client's offset minus yours.
7. In **E2**, type `=IF(OR(C2="",NOT(ISNUMBER(D2))),"",IF(ISNUMBER(C2),C2+D2/24,"Check time"))` and fill it down. The core of it is `=C2+D2/24`.
8. Format column **C** and column **E** as `mm/dd/yyyy h:mm AM/PM`.
9. In **F2**, type `=IF(NOT(ISNUMBER(E2)),"",IF(INT(E2)>INT(C2),"Next day",IF(INT(E2)<INT(C2),"Previous day","Same day")))` and fill it down. INT drops the time and keeps the date, so F2 compares the two dates.

If you type a fixed number instead of the lookup, the short version is `=C2+(-3)/24` for Eastern to Pacific.

### In Excel

The formulas work as typed. Microsoft describes the lookup on its [VLOOKUP function](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1) page. Set the custom format from the **Format Cells** dialog (select the cells and press Ctrl+1).

### In Google Sheets

The same formulas should work as typed. Google describes the lookup on its [VLOOKUP help page](https://support.google.com/docs/answer/3093318). Set the format from **Format > Number > Custom date and time**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. Your zone in I2 is Eastern.

| Client | Client zone | Your call time | Hours difference | Client time | Day |
|---|---|---|---|---|---|
| Maple Street Bakery | Eastern | 10/14/2026 9:00 AM | 0 | 10/14/2026 9:00 AM | Same day |
| Harbor Yoga Studio | Central | 10/14/2026 4:30 PM | -1 | 10/14/2026 3:30 PM | Same day |
| Birch Design Co | Mountain | 10/15/2026 10:00 AM | -2 | 10/15/2026 8:00 AM | Same day |
| Lantern Print Shop | Pacific | 10/14/2026 4:30 PM | -3 | 10/14/2026 1:30 PM | Same day |
| Cedar Row Realty | Pacific | 10/15/2026 1:00 AM | -3 | 10/14/2026 10:00 PM | Previous day |
| Willow Pet Care | Central | 10/15/2026 12:15 AM | -1 | 10/14/2026 11:15 PM | Previous day |
| Granite Fitness | Mountain | 10/16/2026 2:45 PM | -2 | 10/16/2026 12:45 PM | Same day |
| Juniper Books | Pacific | 10/16/2026 6:00 PM | -3 | 10/16/2026 3:00 PM | Same day |

All 8 results matched Python date math. **I4** counts the calls that land on a different day (`=COUNTIF(F2:F20,"Next day")+COUNTIF(F2:F20,"Previous day")`) and shows 2.

## Troubleshooting

### The time shows as a decimal like 0.5625
The cell is formatted as a number. Format it as `mm/dd/yyyy h:mm AM/PM`. The value is right: 0.5625 is 1:30 PM.

### The client time says Check time
The call time in column C is text, not a real date and time. In the recalculated file, "10/14/2026 4:30 PM" typed as text in **C13** gave Check time in E13. Retype it as a date and time so it right-aligns. Some apps convert text dates on their own, so do not rely on that.

### The hours go the wrong way
The sign is reversed. Offsets are hours from Eastern, and D2 is the client's offset minus yours. With your zone set to Pacific in **I2**, an Eastern client shows 3 and a call at 10/14/2026 10:30 PM becomes 10/15/2026 1:30 AM, flagged Next day. Set **I2** to the zone you are sitting in.

### The date changes at midnight
That is correct. A call at 10/15/2026 1:00 AM Eastern is 10/14/2026 10:00 PM Pacific, and the Day column says Previous day. Check these before you confirm a time.

### The zone shows Check zone
The name in column B is not on the Offsets tab. In the recalculated file, "Atlantic" gave Check zone in D14 and left E and F blank. Pick from the dropdown or add a row to the Offsets table and extend the lookup range.

## Limits

These are fixed hour offsets. Most of the US changes clocks on the same dates, so the gap between these four zones stays the same through the year. This sheet does not handle Arizona (most of it), Hawaii or other places that do not change clocks, or any zone outside the four. For the current rules, see the [NIST daylight saving time page](https://www.nist.gov/pml/time-and-frequency-division/popular-links/daylight-saving-time-dst).

## Template

[Download the .xlsx](templates/tidy-tabs-time-zone-converter.xlsx)

The file has three tabs. **Calls** holds 8 sample calls in A2:C9, formulas ready for rows 2 to 20, and your zone, a day-change count and a call count in H2:I5. **Offsets** holds the four zones in A2:B5. **How to use** has short fill-in notes. Zone cells use a dropdown.

The clients and times are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
