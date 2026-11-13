---
title: How to Split a Date Range Across Months by Days in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 15 workbooks and 52,965 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Count the days in each month between two dates with MIN, MAX and EOMONTH, then split the dollar amount by days so the months add up to the cent.
image_prompts: An empty rented stall at an indoor makers market between seasons, bare pegboard walls, a hand truck loaded with folded display shelves and a padlock with key hanging on the stall's wire gate, soft overhead light, realistic photo, no screens, no readable text, no logos
image_alt: Empty indoor market stall with bare pegboard, a hand truck and a padlocked gate
threads: $100.00 over 92 days, rounded month by month, comes to $100.01.\nCount each month's days with =MAX(0,MIN($C5,EOMONTH(G$4,0))-MAX($B5,G$4)+1), then round the running share and subtract what earlier months already got.\n$33.70 + $32.60 + $33.70 = $100.00.
---
Count how many days of a booking fall in each month with `=MAX(0,MIN($C5,EOMONTH(G$4,0))-MAX($B5,G$4)+1)`, where B5 and C5 are the start and end dates and G4 is the first day of the month, then copy it across a row of months. In the made-up sample, a booth season from 10/15/2026 to 01/14/2027 gets 17, 30, 31 and 14 days, and its $1,200.00 splits into $221.74, $391.30, $404.35 and $182.61, which add up to the cent.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (15 workbooks, 52,965 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened; those steps follow Microsoft's and Google's help pages.

![Booking amounts spread over October to March columns by days, with an Adds up check](images/2026-11-14-team-04-split-booking-days-and-amount-across-months-eomonth-min-max-excel-google-sheets/split-amount-across-months-template.png) *Six jobs split by days into month columns, shown through Mar 2027 (later months are hidden here, not in the file); an end date before the start shows Check. PDF export from LibreOffice Calc 24.2.*

## Steps

The formulas are the same in Excel and Google Sheets. The idea: for each month, the days a range covers are its earlier end minus its later start plus one, and never less than zero.

1. In **B1**, type the first month of your grid, such as 10/01/2026.
2. In **G4**, type `=DATE(YEAR($B$1),MONTH($B$1),1)`. This moves any date in B1 to the 1st of its month.
3. In **H4**, type `=EDATE(G4,1)` and fill it right to **R4** for 12 months. Format **G4:R4** with the custom format `mmm yyyy`.
4. In row 5, type the job in **A5**, the start date in **B5**, the end date in **C5** and the dollar amount in **D5**.
5. In **E5**, count the days in the whole range: `=C5-B5+1`. The +1 counts both the start and the end day.
6. In **G5**, type `=MAX(0,MIN($C5,EOMONTH(G$4,0))-MAX($B5,G$4)+1)` and fill it right to **R5**. `EOMONTH(G$4,0)` is the last day of that month. MIN picks the earlier end, MAX the later start, and the outer MAX turns months the job misses into 0.
7. In **F5**, count days outside the grid: `=E5-SUM(G5:R5)`. Above 0 means the job runs past your 12 months.
8. Fill **E5:R5** down to the last job.

Help pages: Microsoft's [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628) and [EDATE function](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5), and Google's [EOMONTH](https://support.google.com/docs/answer/3093044).

### Split the dollars by days

On a second tab (**Amounts** in the template), put the dollars next to the days.

1. In the first month, round the share to cents: `=ROUND($D5*Days!G5/$E5,2)`.
2. From the second month on, round the running share and subtract what earlier months already got: `=ROUND($D5*SUM(Days!$G5:H5)/$E5,2)-SUM($G5:G5)`. Fill it right to **R5**.

The running share is what makes the months add up. If you round each month on its own, $100.00 over 92 days from October to December 2026 gives $33.70 + $32.61 + $33.70 = $100.01. The running share gives $33.70 + $32.60 + $33.70 = $100.00.

This is a planning view of days and dollars per month. It does not say how to book income or expenses; ask your bookkeeper about that.

## Example

The jobs and amounts are fictional. The grid starts at Oct 2026 (B1 = 10/01/2026); the table shows the first five months.

| Job | Start | End | Amount | Days | Oct 2026 | Nov 2026 | Dec 2026 | Jan 2027 | Feb 2027 | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Booth season | 10/15/2026 | 01/14/2027 | $1,200.00 | 92 | 17 | 30 | 31 | 14 | 0 | OK |
| Studio sublet | 12/20/2026 | 02/19/2027 | $750.00 | 62 | 0 | 0 | 12 | 31 | 19 | OK |
| Tent rental | 11/14/2026 | 11/14/2026 | $85.00 | 1 | 0 | 1 | 0 | 0 | 0 | OK |
| Site care retainer | 01/01/2027 | 12/31/2027 | $1,800.00 | 365 | 0 | 0 | 0 | 31 | 28 | 92 days outside |
| Window display | 12/05/2026 | 12/01/2026 | $300.00 | -3 | | | | | | Check |
| Pop-up shop | 02/01/2027 | 03/31/2027 | $590.00 | 59 | 0 | 0 | 0 | 0 | 28 | OK |

On the **Amounts** tab, the dollars come out like this:

- Booth season: $221.74, $391.30, $404.35, $182.61 (total $1,200.00).
- Studio sublet: $145.16 in December, $375.00 in January, $229.84 in February (total $750.00).
- Tent rental: $85.00 in November.
- Pop-up shop: $280.00 in February, $310.00 in March.
- Site care retainer: $152.88 in January, $138.08 in February, and so on through September, $1,346.30 in the grid and $453.70 outside it.

The summary on the **Days** tab shows 6 jobs, 1 row showing Check, 1 job with days outside the grid, $3,971.30 in the grid and $453.70 outside. CHECK shows Check on purpose because of the Window display row; fix its dates or delete the row and it turns OK.

## Troubleshooting

### The months add up to $100.01
Each month was rounded on its own. Use the running share formula above.

### A month is off by one day
The +1 is missing. 10/15/2026 to 01/14/2027 is 92 days counting both ends; without the +1 it is 91.

### The day count is negative
The end date is before the start date. The Window display row (12/05/2026 to 12/01/2026) gives -3, and the template shows Check. Swap or fix the dates.

### The months total less than the amount
The job runs past the grid. The Site care retainer ends 12/31/2027, but the grid stops at Sep 2027, so 92 days and $453.70 sit outside. Status flags it instead of dropping the days. Move **B1** later or read the outside column.

### Every row shows Check
**B1** is text or blank. Type a real date; any day of the month works, since the grid starts on the 1st.

## Template

[Download the .xlsx](templates/tidy-tabs-split-dates-across-months.xlsx)

The file has three tabs. **Days** has the grid start in **B1**, jobs in rows 5 to 104 with Job, Start, End and Amount as inputs in **A:D**, then Days, Days outside the grid, 12 month columns in **G:R** and Status in **S**, with totals in row 105 and the summary in **U4:V9**. **Amounts** repeats the jobs with the dollars per month, Outside the grid ($), In the grid ($) and an Adds up check. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Status turns red for Check and orange for days outside the grid. Both tabs print Landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
