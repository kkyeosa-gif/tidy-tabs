---
threads_url: https://www.threads.com/@tin_ylab/post/Df8blk2I-Wr
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-electricity-cost-of.html
title: How to Calculate the Electricity Cost of Each Batch From Watts in Excel and Google Sheets
labels: excel-formulas, small-business, pricing, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 17 workbooks and 2,669 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Price the power each batch uses with =watts/1000*hours*24*on_percent*rate, run time typed as h:mm. Cost per batch, unit and month. Free template.
image_prompts: A stainless steel wax melting pot on a wooden workbench in a home candle studio, its cord running to a small plug-in electricity usage meter with a blank display at a wall outlet, and a row of empty amber glass candle jars with cotton wicks standing beside it, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Wax melting pot plugged into a small power meter beside empty amber candle jars
threads: Your sheet says the wax melter costs $0.028125 a batch. Cheap power, or bad math?\nBad math. A run time typed as 2:30 is part of a day, so multiply by 24.\n=watts/1000*hours*24*on_percent*rate gives 3.75 kWh and $0.68 at $0.18 per kWh.
---
Find what one run costs with `=watts/1000*hours*24*on_percent*rate`, where hours is the run time typed as h:mm, so the *24 turns the time into hours. In the made-up sample at a rate of $0.18 per kWh, a 1,500 W wax melter running 2:30 uses 3.75 kWh and costs $0.68 per batch, or $5.40 a month at 8 batches, while leaving out the *24 gives $0.028125 instead.

Take the rate from your own power bill: it is the price per kWh. The watts come from the label on the device, which is the most it draws, so the result is an estimate, not a meter reading.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (17 workbooks, 2,669 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (IF, ISNUMBER, ROUND, SUMPRODUCT) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Four workshop devices with watts, run time and percent heating, plus cost per batch](images/2026-11-16-team-02-electricity-cost-per-batch-from-watts-hours-kwh-rate-excel-google-sheets/electricity-cost-per-batch-template.png) *Four fictional devices at $0.18 per kWh: the 1,500 W wax melter running 2:30 uses 3.75 kWh and costs $0.68 a batch, $5.40 a month (the month summary is hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Equipment** tab, one row per device starting in row 5.

1. Type the rate per kWh from your bill in **B1** (the sample uses $0.18).
2. Type the device name in **A5** and the watts from its label in **B5** (1500, not "1,500 W" and not 1.5).
3. Type the run time per batch in **C5** as h:mm: 2:30 for two and a half hours. Format the column as **[h]:mm** so a full day shows as 24:00.
4. Type the percent of that time it is actually heating in **D5** (70% for a kiln that cycles off). Leave it blank for a device that runs the whole time; the template counts blank as 100%.
5. Type batches per month in **E5** and units per batch in **F5**.
6. In **G5**, type `=B5/1000*C5*24*D5` for kWh per batch. Dividing by 1,000 turns watts into kilowatts. A time such as 2:30 is stored as part of a day (0.104), so *24 turns it into 2.5 hours (more in [how to convert clock times to decimal hours](https://tidytabs.blogspot.com/2026/10/how-to-convert-clock-times-to-decimal.html)).
7. In **H5**, type `=ROUND(G5*$B$1,2)` for the cost per batch.
8. In **I5**, type `=G5*$B$1/F5` for the cost per unit, and in **J5**, `=G5*$B$1*E5` for the cost per month. Both use the unrounded cost, so the month total is not thrown off by rounding each batch.
9. Fill **G5:J5** down and add the month with `=SUM(J5:J34)`.

The template wraps these in checks. G5 stays blank until the row has a name and number inputs, and treats a blank D5 as 100%. **K5** shows a Status, and **N5** recomputes the month as `=SUMPRODUCT(G5:G34,E5:E34)*B1` so the CHECK in **N10** can confirm it matches the sum of column J.

Help pages: Microsoft's [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) and [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e), and Google's [SUMPRODUCT](https://support.google.com/docs/answer/3094294).

## Example

The equipment, run times and rate are fictional.

| Equipment | Watts | Run time | Percent heating | Batches per month | Units per batch | kWh per batch | Cost per batch | Cost per unit | Cost per month | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Wax melter | 1,500 | 2:30 | 100% | 8 | 24 | 3.75 | $0.68 | $0.03 | $5.40 | OK |
| Heat gun | 1,200 | 0:20 | 100% | 8 | 24 | 0.40 | $0.07 | $0.00 | $0.58 | OK |
| Test kiln | 1,800 | 6:00 | 70% | 4 | 30 | 7.56 | $1.36 | $0.05 | $5.44 | OK |
| Drying fan | 60 | 24:00 | 100% | 30 | | 1.44 | $0.26 | | $7.78 | OK |

The cells hold more decimals than they show. The heat gun costs $0.003 a unit, the kiln $0.04536, and the month costs are $5.40, $0.576, $5.4432 and $7.776.

The summary in **N4:N10** shows a total of $19.20 a month (19.1952 in the cell), the same 19.1952 from SUMPRODUCT, 106.64 kWh a month, 4 rows listed, 0 rows showing Check, 0 typed below row 34 and CHECK OK.

## Troubleshooting

### The cost is far too small
You left out the *24. The wax melter then comes out at $0.028125 instead of $0.675, because 2:30 is 0.104 of a day, not 2.5 hours. Multiply the time by 24.

### The cost is 24 times too big
You typed 2.5 instead of 2:30. A plain 2.5 in a time cell means 2.5 days, so the melter shows 90 kWh and $16.20 a batch. The template flags it as Check: time. Type 2:30.

### A one-hour run costs as much as a full day
You typed 1 instead of 1:00. A plain 1 is one whole day, which shows as 24:00, and exactly 24:00 is allowed (the drying fan runs that long), so the Status stays OK. Only more than 24:00 shows Check: time. Type one hour as 1:00.

### Check: watts on a row with a number
The watts are under 10, which is usually kilowatts typed instead of watts: 1.5 instead of 1500. Text such as "1,500 W" shows Check: watts too. Type the plain number of watts from the label.

### 24:00 shows as 0:00
The cell uses plain h:mm. In the LibreOffice PDF export, a full day showed as 0:00 with h:mm (and 30:00 as 6:00) but as 24:00 with [h]:mm, which is why the template formats C5:C34 as [h]:mm. The post on [summing hours over 24](https://tidytabs.blogspot.com/2026/10/how-to-sum-hours-over-24-in-excel-and.html) covers the format.

## Template

[Download the .xlsx](templates/tidy-tabs-electricity-cost-per-batch.xlsx)

The file has two tabs. **Equipment** has the rate in **B1** and rows 5 to 34: Equipment, Watts (label), Run time per batch (h:mm), Percent of time heating, Batches per month and Units per batch (inputs in **A:F**), then kWh per batch, Cost per batch, Cost per unit, Cost per month and Status (formulas in **G:K**). The summary in **M4:N10** holds the month total, the SUMPRODUCT total, kWh per month, rows listed, rows showing Check, entries typed below row 34 and the CHECK. **How to use** explains each column and the Status words (Check: watts, time, percent, batches, units, rate, name). The Equipment tab prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. A Status starting with Check turns red, and so does the CHECK cell when it is not OK. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
