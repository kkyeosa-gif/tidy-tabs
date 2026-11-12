---
threads_url: https://www.threads.com/@tin_ylab/post/Dfp-jQJFHY6
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-price-rental-with-weekday-and.html
title: How to Price a Rental With Weekday and Weekend Day Rates in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 16 workbooks and 16,112 comparisons with 0 mismatches. Checked: every pickup weekday from 11/01/2026 to 11/14/2026 with 0, 1, 2, 3, 6, 7, 8 and 15 days, holidays on a weekday, on a Saturday (07/04/2026, 11/07/2026) and listed twice, rentals over 12/25/2026 and 01/01/2027, a return before the pickup, a text pickup date, a pickup with a time, no return date, rates 0/0, 45/-1, "abc", blank and 45.50/65.25, no holidays, and random sets of 200 rentals. Probes: a holiday typed as text ("11/27/2026") and the word "Black Friday" in the holiday list were ignored with no error; =D5-C5+1 gives 4 for a Friday to Monday rental. Excel and Google Sheets were not used. Not run: text in the holiday list in Excel or Sheets, other weekend days (NETWORKDAYS.INTL), a non-US language setting, more rows than the template holds.
search_description: Split each rental into weekday and weekend or holiday days with NETWORKDAYS, then price each part at its own day rate. Free Excel template.
image_prompts: A folded white canopy tent in its carry bag leaning on a stack of folding chairs by an open garage door, a paper wall calendar with its weekend columns shaded in pencil on the wall beside it, dates unreadable, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Folded canopy tent bag leaning on folding chairs beside a shaded paper wall calendar
threads: 3 days billed, 2 of them weekend, $175.00. That is a Friday to Monday tent rental at $45.00 a weekday and $65.00 a weekend day.\nThe trick: days billed minus NETWORKDAYS is the weekend and holiday count.\nType holidays as real dates. In LibreOffice a holiday typed as text was ignored and billed at the weekday rate (made-up rentals).
---
Count the weekend and holiday days in a rental with `=E5-NETWORKDAYS(C5,C5+E5-1,Holidays!$A$4:$A$23)`, where C5 is the pickup date and E5 the days billed (`=MAX(1,D5-C5)`), then price it with `=ROUND(G5*$B$1+F5*$B$2,2)`. With made-up rates of $45.00 a weekday and $65.00 a weekend day, a Friday pickup and Monday return is 3 days, 2 of them weekend, for $175.00. The pickup day is billed and the return day is not.

The billing rule, the rates and the holiday list are settings in the file, not rules you have to follow. The file does not decide what to charge.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (16 workbooks, 16,112 comparisons). The formulas use only IF, AND, OR, NOT, ISNUMBER, INT, MAX, ROUND, TRIM, SUM, SUMPRODUCT, COUNT, COUNTIF and NETWORKDAYS, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Rentals tab with seven sample rentals, days billed, weekend days and prices](images/2026-11-12-team-03-rental-price-weekday-weekend-day-rates-networkdays-excel-google-sheets/rental-weekday-weekend-rate-template.png) *Seven rentals priced at $45.00 a weekday and $65.00 a weekend or holiday day. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Rentals** tab, a **Holidays** tab and a **How to use** tab.

1. On **Rentals**, type the weekday day rate in **B1** (for example `45`) and the weekend or holiday day rate in **B2** (for example `65`).
2. On **Holidays**, type each holiday as a real date in **A4:A23** and its name in **B4:B23**. Dates on this list are billed at the weekend rate.
3. Back on **Rentals**, enter one rental per row from row 5: customer in **A**, item in **B**, pickup date in **C**, return date in **D**, as MM/DD/YYYY.
4. In **E5**, type `=MAX(1,D5-C5)` for the days billed. A same-day return counts as 1 day.
5. In **F5**, type `=E5-NETWORKDAYS(C5,C5+E5-1,Holidays!$A$4:$A$23)` for the weekend and holiday days.
6. In **G5**, type `=E5-F5` for the weekday days.
7. In **H5**, type `=ROUND(G5*$B$1+F5*$B$2,2)` for the price.
8. Fill **E5:H5** down to the last rental. The template holds rows 5 to 204.

How it works: `C5+E5-1` is the last billed day, so NETWORKDAYS looks at exactly the days you charge for. NETWORKDAYS counts Monday to Friday days that are not on the holiday list. Every billed day it skips is a Saturday, a Sunday or a listed holiday.

The weekend in NETWORKDAYS is Saturday and Sunday only. A holiday that falls on a Saturday or Sunday is not counted twice; that was checked with 07/04/2026 and 11/07/2026, both Saturdays. If your weekend is different days, Microsoft's [NETWORKDAYS.INTL function](https://support.microsoft.com/en-us/office/networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28) takes a weekend setting, but the template does not use it.

The template wraps these formulas so blank rows stay blank and bad dates or rates show Check. Column **I** shows OK, Same-day return or Check, and Check is filled red.

Help page: Microsoft's [NETWORKDAYS function](https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7). It describes the function, not our test.

## Example

Sample data is fictional: 7 event-equipment rentals at $45.00 a weekday and $65.00 a weekend or holiday day, with Thanksgiving Day (11/26/2026), Christmas Day (12/25/2026) and New Year's Day (01/01/2027) on the Holidays tab. The weekday names are added here; the file shows dates only.

| Customer | Pickup | Return | Days billed | Weekend and holiday | Weekday | Price | Status |
|---|---|---|---|---|---|---|---|
| Lopez family | 11/06/2026 (Fri) | 11/09/2026 (Mon) | 3 | 2 | 1 | $175.00 | OK |
| Harbor Dental | 11/10/2026 (Tue) | 11/12/2026 (Thu) | 2 | 0 | 2 | $90.00 | OK |
| Jordan Reed | 11/14/2026 (Sat) | 11/14/2026 (Sat) | 1 | 1 | 0 | $65.00 | Same-day return |
| Cedar Print Shop | 11/16/2026 (Mon) | 11/23/2026 (Mon) | 7 | 2 | 5 | $355.00 | OK |
| Lakeside Yoga | 11/25/2026 (Wed) | 11/30/2026 (Mon) | 5 | 3 | 2 | $285.00 | OK |
| Maria Lopez | 12/05/2026 (Sat) | 12/08/2026 (Tue) | 3 | 2 | 1 | $175.00 | OK |
| Pine & Co. Realty | 12/24/2026 (Thu) | 12/28/2026 (Mon) | 4 | 3 | 1 | $240.00 | OK |

The Lakeside Yoga patio heater is billed Wednesday to Sunday. Thanksgiving, Saturday and Sunday are the 3 weekend and holiday days, so the price is 2 x $45.00 + 3 x $65.00 = $285.00. The Pine & Co. Realty tent picks up Christmas Eve, so Christmas Day, Saturday and Sunday are the 3 days at the higher rate.

The summary block in **K4:L10** shows 7 rentals listed, 25 days billed, 13 weekend and holiday days, 12 weekday days, a total of $1,385.00, 0 rows showing Check and a CHECK cell that says OK.

## Troubleshooting

### Status shows Check
A date is not a plain date. The return date is before the pickup (11/20/2026 to 11/18/2026), the pickup was typed as text, the pickup has a time in it (11/20/2026 9:30 AM), or the return date is blank. Retype the dates as MM/DD/YYYY with no time.

### A holiday is billed at the weekday rate
The holiday was typed as text in the Holidays list. In LibreOffice, "11/27/2026" stored as text and the word "Black Friday" were both ignored with no error, so that day stayed at the weekday rate. How Excel and Google Sheets treat text in the holiday range was not tested. Retype the holiday as a real date.

### One day too few or too many
This file bills the pickup day and not the return day, so Friday to Monday is 3 days. If you bill the return day too, change E5 to `=D5-C5+1`, which gives 4 for the Friday to Monday rental. The NETWORKDAYS end in F5 is `C5+E5-1`, so it follows the new count.

### Price shows Check on every row
A rate in B1 or B2 is blank, text or negative. Type a number of 0 or more in both cells.

## Template

[Download the .xlsx](templates/tidy-tabs-rental-weekend-rate.xlsx)

The file has three tabs. **Rentals** has the rates in **B1:B2**, rentals in **A5:D204**, the formula columns **E:I** and a summary with a CHECK cell in **K4:L10**. **Holidays** holds up to 20 dates in **A4:A23** with three sample holidays. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
