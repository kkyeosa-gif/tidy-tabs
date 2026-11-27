---
title: How to Total Unbilled Hours by Client in Excel and Google Sheets
labels: excel-formulas, freelancers, time-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 222 workbooks and 167,740 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets. The SUMIFS blank-criterion result and the CHAR(160) difference were seen in LibreOffice Calc only.
search_description: Total unbilled hours and dollars per client from a time log with a blank Invoice # column, find the oldest entry, and catch cells that only look blank.
image_prompts: A freelancer's desk edge with a stack of printed weekly timesheets held by a black binder clip, a small spiral notebook open to handwritten hour tallies, and a pad of blank carbon invoice forms beside a pen, late afternoon window light, realistic photo, no screens, no readable text, no logos
image_alt: Clipped printed timesheets, a notebook of hour tallies and blank carbon invoice forms
threads: How many hours did you work and never invoice?\nAny entry with a blank Invoice # is unbilled, and TRIM(...)="" also catches the cells that only hold spaces.\nIn the sample log that adds up to 20.75 hours and $1,245.00, the oldest from 10/09/2026.
---
Total each client's unbilled time with `=SUMPRODUCT(('Time log'!$B$5:$B$1004=A5)*(TRIM('Time log'!$D$5:$D$1004)="")*('Time log'!$G$5:$G$1004="OK"),'Time log'!$C$5:$C$1004)`, which counts an entry as unbilled when its Invoice # cell is empty or holds only spaces (the Status test leaves out rows with a typo; the hours go in as a separate argument so a text entry cannot break the total). In the made-up log, 20.75 hours worth $1,245.00 are not billed yet, and the oldest is Pine & Co. Realty's entry from 10/09/2026.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (222 workbooks, 167,740 cell comparisons, 0 mismatches), plus the edge cases below. The SUMIFS blank-criterion result and the CHAR(160) difference below were seen in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Five clients with unbilled hours, dollars, oldest entry dates and overdue flags](images/2026-11-28-team-03-unbilled-hours-and-amount-by-client-blank-invoice-number-sumifs-sumproduct-excel-google-sheets/unbilled-hours-by-client-template.png) *Clients tab of a made-up freelancer's time log as of 11/27/2026: per client, the hours and dollars from entries with a blank invoice number, the oldest unbilled date and how many are older than 30 days (orange), with totals of 20.75 hours and $1,245.00 in green and the checks. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

You type the invoice number next to each entry once you bill it, so any entry with a blank Invoice # is unbilled. When to invoice and which tax year the income belongs to are not covered here.

The steps are the same in Excel and Google Sheets.

1. On the **Clients** tab, type the as-of date in **B1** (for example 11/27/2026) and the age that counts as overdue in **B2** (for example 30 days).
2. From row 5, type each client's name in column **A** and hourly rate in column **B**. Each client goes on one row only.
3. On the **Time log** tab, type one entry per row from row 5: Date in **A**, Client in **B**, Hours in **C**. Leave **D** (Invoice #) empty until the entry is billed, then type the invoice number.
4. Check column **G** (Status). It shows OK, a Check message, or "After the as-of date (not counted)". Only OK rows count.
5. Column **E** prices each OK entry with `=IF(G5<>"OK","",ROUND(C5*INDEX(Clients!$B$5:$B$54,MATCH(B5,Clients!$A$5:$A$54,0)),2))`.
6. Column **F** (Looks blank?) flags an Invoice # that holds only spaces as "Yes: spaces only (not billed)" and one that holds a non-breaking space as "Check: non-breaking space".
7. Back on **Clients**, column **C** totals unbilled hours with the SUMPRODUCT above, and **D** uses the same formula with the Amount column `'Time log'!$E$5:$E$1004` to total unbilled dollars. `TRIM(...)=""` is what treats a space-only cell as blank. See the SUMPRODUCT help from [Microsoft](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) or [Google](https://support.google.com/docs/answer/3094294), and the TRIM help from [Microsoft](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) or [Google](https://support.google.com/docs/answer/3094140).
8. Column **E** shows the oldest unbilled date: Time log column **H** holds each unbilled entry's age, and **E5** subtracts the largest from the as-of date with `=IF(OR(H5<>"OK",N(G5)=0),"",$B$1-SUMPRODUCT(MAX(('Time log'!$B$5:$B$1004=A5)*'Time log'!$H$5:$H$1004)))`.
9. Column **F** counts entries older than **B2** days (exactly 30 is not older than 30) and **G** counts unbilled entries.
10. Read the totals in **J4:K15**. **CHECK** in **K15** is OK only when the settings are valid, no Invoice # holds a non-breaking space, no client row shows Check, the Time log CHECK is OK, and the per-client dollars add up to the unbilled dollars in the whole log.

Anything else in Invoice # counts as billed, including "-" or "N/A", so leave the cell empty until you bill.

Related: [round billable time up to 15 minutes](https://tidytabs.blogspot.com/2026/10/how-to-round-billable-time-up-to.html), [prepaid retainer hours](https://tidytabs.blogspot.com/2026/10/how-to-track-prepaid-retainer-hours.html) and the [Freelance Project Tracker](https://tidytabs.blogspot.com/2026/09/freelance-project-tracker-in-google.html).

## Example

The log is made up: a freelance designer's 16 entries from 09/14/2026 to 11/24/2026 for five clients, as of 11/27/2026 with a 30-day flag. Rates are $55.00 (Pine & Co. Realty), $60.00 (Lakeside Yoga), $75.00 (Harbor Dental), $50.00 (Maple Street Bakery) and $65.00 (Cedar Hill Florist).

| Client | Unbilled hours | Unbilled $ | Oldest unbilled | Older than 30 days | Unbilled entries |
|---|---|---|---|---|---|
| Pine & Co. Realty | 8.25 | $453.75 | 10/09/2026 | 1 | 3 |
| Lakeside Yoga | 6.00 | $360.00 | 10/28/2026 | 0 | 3 |
| Harbor Dental | 4.25 | $318.75 | 10/21/2026 | 1 | 2 |
| Maple Street Bakery | 2.25 | $112.50 | 11/05/2026 | 0 | 2 |
| Cedar Hill Florist | 0.00 | $0.00 | | 0 | 0 |

Of 36.50 logged hours, 20.75 hours in 10 entries are unbilled, worth **$1,245.00**. Two entries are older than 30 days: Pine & Co. Realty's at 49 days and Harbor Dental's at 37. Lakeside Yoga's 10/28 entry is exactly 30 days old, so it is not flagged.

Two of Lakeside Yoga's Invoice # cells only look empty: the 10/28 cell holds one space, and the 11/24 cell holds the formula `=""`. Both count as unbilled, and **K10** shows 1 cell with only spaces.

## Troubleshooting

### A SUMIFS with "" for blank invoice numbers comes up short
One cell holds a space. In LibreOffice Calc, `=SUMIFS(C:C,B:B,"Lakeside Yoga",D:D,"")` typed on the sample Time log gave 4.00 hours instead of 6.00: it counted the empty cell and the `=""` cell but skipped the 2-hour entry with a space. `D:D,"="` gave 3.00, counting only the truly empty cell. These are LibreOffice results; see the SUMIFS help from [Microsoft](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) or [Google](https://support.google.com/docs/answer/3238496) for your app, or use the TRIM version, which gave 6.00.

### Looks blank? says "Check: non-breaking space" and CHECK is red
The cell holds a pasted non-breaking space, which TRIM left in place in LibreOffice, so the entry counts as billed. Clear the cell. In LibreOffice under the C.UTF-8 locale, `=CODE(CHAR(160))` gave 239 instead of 160, so the template's check uses a typed non-breaking space inside SUBSTITUTE. More in [How to Remove Non-Breaking Spaces](https://tidytabs.blogspot.com/2026/10/how-to-remove-non-breaking-spaces-in.html).

### An entry is missing from a client's total
Check its Status. "Check: client not on the Clients tab" means a leading or trailing space, such as "Lakeside Yoga "; capital letters do not matter. "After the as-of date (not counted)" means its date is later than **B1**.

### Both rows of a client say "Check: client listed twice"
The name is on two Clients rows (capital letters do not matter, so "lakeside yoga" clashes with "Lakeside Yoga"). Delete one row.

## Template

[Download the .xlsx](templates/tidy-tabs-unbilled-hours-by-client.xlsx)

**Time log** takes 1,000 entries (rows 5 to 1004) with its own CHECK in **J4:K9**, which also catches rows typed below row 1004. **Clients** holds the as-of date and age in **B1:B2**, up to 50 clients (rows 5 to 54) and the summary in **J4:K15**. **How to use** explains each formula. Data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs, gray cells are formulas, unbilled totals are light green, old entries and space-only cells are light orange, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
