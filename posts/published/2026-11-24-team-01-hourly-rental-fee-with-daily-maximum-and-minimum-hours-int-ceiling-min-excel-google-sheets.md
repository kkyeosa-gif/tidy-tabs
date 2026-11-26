---
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-charge-hourly-rental-rate-with.html
title: How to Charge an Hourly Rental Rate With a Daily Maximum in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 36 workbooks and 43,488 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The raw-hours rounding error was measured in LibreOffice Calc only (3,000 random bookings).
search_description: Bill rentals by the hour with a cap per 24 hours and a minimum: count minutes, split into 24-hour blocks with INT, round up with CEILING, cap with MIN.
image_prompts: A small camera gear rental counter with an open hard-shell equipment case holding a camera lens in foam, a folded light stand leaning beside it and a plain wall clock without numbers above the counter, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Open equipment case with a camera lens and light stand at a rental counter
threads: Charge each full 24-hour block at the daily max, then cap only the hours left over.\nCount whole minutes first, split them with =INT(minutes/1440) and round the rest up with CEILING.\nAt $35.00 an hour and $200.00 per 24 hours, two days and 2.5 hours comes to $487.50, not $600.00.
---
Count whole minutes first with `=ROUND((end-start)*1440,0)`, split them into full 24-hour blocks with `=INT(minutes/1440)`, round the minutes left over up to your billing step with `=CEILING(left,step)/60`, and charge `=MAX(minimum_charge,blocks*daily_max+MIN(hours*rate,daily_max))`. In the made-up bookings at $35.00 an hour and $200.00 per 24 hours, 11/06/2026 9:00 AM to 11/08/2026 11:15 AM comes to $487.50 (two full days plus 2.5 hours), while the common `MIN(hours*rate,daily_max*ROUNDUP(hours/24,0))` formula charges $600.00.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (36 workbooks, 43,488 cell comparisons, 0 mismatches), plus the edge cases below. The raw-hours rounding error in Troubleshooting was measured in LibreOffice Calc only (3,000 random bookings). Excel and Google Sheets were not opened for this post.

![Nine sample rental bookings with minutes, full 24-hour blocks, charge and Capped status](images/2026-11-24-team-01-hourly-rental-fee-with-daily-maximum-and-minimum-hours-int-ceiling-min-excel-google-sheets/hourly-rental-daily-max-template.png) *Bookings tab with nine fictional bookings at $35.00 an hour and a $200.00 maximum per 24 hours: Tri-County Wedding Expo, 11/06/2026 9:00 AM to 11/08/2026 11:15 AM, comes to $487.50 (2 full blocks plus 2.5 hours), five bookings are Capped and the one-hour booking pays the 2-hour Minimum of $70.00. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The rate, maximum, billing step, grace minutes and minimum below are sample settings, not advice on what to charge. The sheet does not handle deposits, damage, cancellation fees or lodging tax.

The cap here is per 24 hours from the start time, not per calendar day. Charging by calendar day is a different rule; for that, see [How to Price a Rental With Weekday and Weekend Day Rates](https://tidytabs.blogspot.com/2026/11/how-to-price-rental-with-weekday-and.html).

The steps are the same in Excel and Google Sheets.

1. On the **Bookings** tab, type your settings in **B1:B5**: Hourly rate ($35.00), Maximum per 24 hours ($200.00), Billing step in minutes (30), Grace minutes after full 24 hours (10) and Minimum hours (2).
2. From row 8, type one booking per row: Customer in **A**, Start date-time in **B** and End date-time in **C**, such as 11/02/2026 9:00 AM.
3. In **D8**, count whole minutes with `=ROUND((C8-B8)*1440,0)`. There are 1440 minutes in a day, and ROUND drops the tiny decimal the time math leaves. The template wraps this in a check: `=IF(OR(NOT(ISNUMBER(B8)),NOT(ISNUMBER(C8)),$L$7<>"OK"),"",IF(C8<=B8,"",ROUND((C8-B8)*1440,0)))`.
4. In **E8**, count full 24-hour blocks with `=IF(D8="","",INT(D8/1440))`. See the [INT help from Microsoft](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef).
5. In **F8**, get the minutes left over with `=IF(D8="","",D8-E8*1440)`.
6. In **G8**, round those minutes up to the step and turn them into hours with `=IF(D8="","",IF(F8<=$B$4,0,CEILING(F8,$B$3)/60))`. Up to 10 minutes past a full block bills nothing extra. See the [CEILING help from Microsoft](https://support.microsoft.com/en-us/office/ceiling-function-0a5cd7c8-0720-4f0a-bd2c-c943e510899f).
7. In **H8**, charge each full block at the maximum, cap the hours left over at the maximum, and apply the minimum: `=IF(D8="","",ROUND(MAX($B$5*$B$1,E8*$B$2+MIN(G8*$B$1,$B$2)),2))`.
8. Fill **D8:I8** down to row 207. **I** shows Capped, Minimum, OK or a Check message.

## Example

The sample data is made up. Settings: $35.00 an hour, $200.00 per 24 hours, 30-minute step, 10 grace minutes, 2-hour minimum.

| Customer | Start | End | Minutes | Blocks | Left over | Hours billed | Charge | Status |
|---|---|---|---|---|---|---|---|---|
| Rivera Family Photos | 11/02/2026 9:00 AM | 11/02/2026 1:30 PM | 270 | 0 | 270 | 4.5 | $157.50 | OK |
| Northside Bakery Co. | 11/03/2026 9:00 AM | 11/03/2026 1:40 PM | 280 | 0 | 280 | 5.0 | $175.00 | OK |
| Lena Park Headshots | 11/04/2026 8:00 AM | 11/04/2026 5:00 PM | 540 | 0 | 540 | 9.0 | $200.00 | Capped |
| Oak & Ember Catering | 11/05/2026 6:00 PM | 11/06/2026 10:00 AM | 960 | 0 | 960 | 16.0 | $200.00 | Capped |
| Tri-County Wedding Expo | 11/06/2026 9:00 AM | 11/08/2026 11:15 AM | 3,015 | 2 | 135 | 2.5 | $487.50 | Capped |
| J. Alvarez Studio | 11/09/2026 10:00 AM | 11/09/2026 11:00 AM | 60 | 0 | 60 | 1.0 | $70.00 | Minimum |
| Hillside Pop-up Market | 11/10/2026 9:00 AM | 11/11/2026 9:00 AM | 1,440 | 1 | 0 | 0.0 | $200.00 | Capped |
| Maya Chen Portraits | 11/12/2026 7:15 AM | 11/12/2026 1:25 PM | 370 | 0 | 370 | 6.5 | $200.00 | Capped |
| Green Fork Meal Prep | 11/13/2026 9:10 AM | 11/13/2026 1:40 PM | 270 | 0 | 270 | 4.5 | $157.50 | OK |

The 9-hour headshot booking would be $315.00 at the hourly rate; the cap makes it $200.00. The 1-hour booking pays the $70.00 minimum. The summary in **K7:L14** shows 9 bookings, $1,847.50 total, 5 capped, 1 at the minimum, 0 rows showing Check and CHECK OK.

Edge cases with the same settings, each starting 11/02/2026 9:00 AM:

| Booking length | Charge |
|---|---|
| Exactly 24 hours | $200.00 (Capped) |
| 24 hours 10 minutes | $200.00 (grace) |
| 24 hours 11 minutes | $217.50 |
| 5 minutes | $70.00 (Minimum) |

A booking from 12/31/2026 10:00 PM to 01/02/2027 1:00 AM is one block plus 3 hours, $305.00.

## Troubleshooting

### A 1-hour booking bills 1.5 hours
The formula rounds raw hours. In LibreOffice Calc, both `=ROUNDUP((end-start)*24*2,0)/2` and `=CEILING((end-start)*24,0.5)` billed 1.5 hours for 10:00 AM to 11:00 AM, because the hours read back as 1.00000000046556. Of 3,000 random bookings, ROUNDUP on raw hours was wrong on 12 and CEILING on 13; rounding to whole minutes first was wrong on 0.

### A booking just over two days costs three full days
The web formula `MIN(hours*rate,daily_max*ROUNDUP(hours/24,0))` caps per started day, so 11/06/2026 9:00 AM to 11/08/2026 11:15 AM comes to $600.00. Splitting into full blocks plus the minutes left over gives $487.50.

### Charge is blank and Status says Check: settings (B1:B5)
A setting in **B1:B5** is off: the step is 0 or text, the grace is 1440 or more, or the minimum hours times the rate is above the maximum. Fix it and every row fills in.

### Charge is blank and Status says Check
The end is not after the start (Check: end is not after start), or the start or end is blank or typed as text (Check: start or end blank or text). Retype it as a date and time.

## Template

[Download the .xlsx](templates/tidy-tabs-hourly-rental-daily-max.xlsx)

The file has two tabs. **Bookings** holds the settings in **B1:B5** and up to 200 bookings in rows 8 to 207: you type **A:C**, and **D:I** hold the formulas above. The summary in **K7:L14** checks the settings, totals the charges, counts Capped, Minimum and Check rows and ends in a CHECK cell. **How to use** explains each formula. Bookings prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Capped turns light blue, Minimum light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
