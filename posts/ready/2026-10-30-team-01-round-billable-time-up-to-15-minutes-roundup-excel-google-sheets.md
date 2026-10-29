---
title: How to Round Billable Time Up to the Nearest 15 Minutes in Excel and Google Sheets
labels: excel-formulas, freelancers, time-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python model (20 workbooks, 907 cell comparisons, 0 mismatches). Checked: blocks of 1, 6, 15, 30 and 60 minutes, 0 and 1 minute, exact multiples, seconds in the times, a start after the end, a missing start or end, text times, and invalid settings. Excel and Google Sheets were not used.
search_description: Round billable time up to the nearest 15 minutes in Excel and Google Sheets with ROUNDUP. Free template with a block size you set and billed hours.
image_prompts: A wind-up kitchen timer, a paper invoice pad with a pencil, and a small stack of manila client folders on a wooden freelancer's desk, soft window light, no screens, no readable text, no logos
image_alt: Wind-up kitchen timer beside an invoice pad and pencil in front of client folders
threads: Billing in 15 minute blocks and your hours never match the clock?\nMinutes worked x 1440 gives whole minutes, then =ROUNDUP(F5/15,0)*15 snaps them to the next block.\n46 minutes becomes 60, 45 stays 45, and one cell holds the block size so 6 or 30 works too.
---
Round the minutes worked up to the next block with `=ROUNDUP(F5/Settings!$B$1,0)*Settings!$B$1`, where F5 is the minutes worked and Settings B1 is the block size, such as 15. With a 15 minute block, 46 minutes becomes 60 and 45 stays 45.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. ROUNDUP, ROUND, IF and SUM are old functions that exist in Excel and Google Sheets, but neither app was opened for this post.

![Time log showing minutes worked rounded up to billed minutes and hours for fictional entries](images/2026-10-30-team-01-round-billable-time-up-to-15-minutes-roundup-excel-google-sheets/billable-time-rounding-template.png) *caption: made from the template with fictional sample data, PDF render from LibreOffice Calc 24.2.7.2*

## Steps

The trick is to turn the clock times into whole minutes first. Then round that count up to a multiple of the block size. This post shows the arithmetic only. It does not say what a client agreement allows, so check your own terms.

The template layout: Start is in **D**, End is in **E**, and the first entry is on row 5. The block size is in **Settings!B1** and the hourly rate is in **Settings!B2**.

### In Excel

1. On a Settings sheet, type `15` in **B1** (block size in minutes) and `85` in **B2** (hourly rate).
2. On the log sheet, type Start times in **D5** and End times in **E5**, for example `9:02 AM` and `9:49 AM`.
3. In **F5**, type `=ROUND((E5-D5)*1440,0)`. A time is a fraction of a day, and a day has 1,440 minutes. ROUND removes any tiny leftover decimals.
4. In **G5**, type `=ROUNDUP(F5/Settings!$B$1,0)*Settings!$B$1`. This divides by the block, rounds up to a whole number of blocks, and multiplies back.
5. In **H5**, type `=G5/60` for billed hours. In **I5**, type `=ROUND(H5*Settings!$B$2,2)` for the amount.
6. In **J5**, type `=G5-F5`. This is the extra minutes the rounding added.
7. Select **F5:J5** and drag the fill handle down.
8. Select **D5:E5** and press **Ctrl+1** to format the times as **Time**. Format **I** as **Currency**.

Microsoft's help pages: [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) and [CEILING function](https://support.microsoft.com/en-us/office/ceiling-function-0a5cd7c8-0720-4f0a-bd2c-c943e510899f). CEILING can round up to a multiple in one step, but ROUNDUP with a divide and multiply is easy to read and matches the template. These pages describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in the Excel steps, typed exactly as written.
2. Format D and E with **Format > Number > Time**, and I with **Format > Number > Currency**.

Google's help page: [ROUNDUP](https://support.google.com/docs/answer/3093443). That page describes the function, not our test.

### Change the block size

Type `6` in **Settings!B1** for six-minute blocks, or `30` for half hours. Nothing else changes, because every formula reads the block from that cell. A block of `1` bills every minute with no rounding up.

## Example

The sample data is fictional: a freelance bookkeeper, four invented clients, a 15 minute block and a $85.00 hourly rate. This table shows 6 of the 15 entries.

| Start | End | Minutes worked | Billed minutes | Billed hours | Amount | Extra minutes |
|---|---|---|---|---|---|---|
| 9:02 AM | 9:49 AM | 47 | 60 | 1.00 | $85.00 | 13 |
| 10:00 AM | 10:45 AM | 45 | 45 | 0.75 | $63.75 | 0 |
| 10:00 AM | 10:46 AM | 46 | 60 | 1.00 | $85.00 | 14 |
| 1:15 PM | 1:15 PM | 0 | 0 | 0.00 | $0.00 | 0 |
| 8:30 AM | 9:31 AM | 61 | 75 | 1.25 | $106.25 | 14 |
| 2:00 PM | 2:01 PM | 1 | 15 | 0.25 | $21.25 | 14 |

An exact multiple stays put, 0 minutes stays 0, and 1 minute becomes a full block. The extra minutes are always less than one block.

The summary sits in **L5:M11**. For all 15 entries it shows 799 actual minutes, 885 billed minutes, 86 extra minutes and a total of $1,253.75. **M11** says OK.

## Troubleshooting

### A row says Check times
The End is before the Start, one of the two is empty, or a time was typed as text. The template leaves that row out of every total, and **M10** counts such rows. Retype the time as a real time, such as `9:00 AM`. A shift that runs past midnight is not handled by this template.

### The billed minutes look one block too high when there are seconds
The template rounds the difference to the nearest whole minute first, then rounds up to the block. In the check, 14 minutes 29 seconds counted as 14, and 14 minutes 31 seconds counted as 15. A client agreement may count seconds another way, so decide that before you rely on the numbers.

### Settings shows a message instead of OK
**Settings!B3** accepts a block that is a whole number from 1 to 480 and a rate of 0 or more. A block of 0, 2.5, 481 or the text "fifteen" makes column G show Check settings, and the billed numbers stay blank. Fix B1 or B2 and the rows fill in again.

### M11 says CHECK
The check fails when the settings are invalid, a billed number is not a whole number of blocks, or something is typed below row 64. The template has 60 entry rows. In the test, a date typed in A65 turned **M11** to CHECK. Use the free rows, or insert rows inside the table.

### The amount is off by a cent from your invoice
The amount is rounded to cents on each row, not on the total. If your invoicing tool rounds the total instead, the two can differ by a cent or two.

## Template

[Download the .xlsx](templates/tidy-tabs-billable-time-rounding.xlsx)

The file has three tabs. **Time log** has 60 entry slots in **A5:E64** with 15 fictional entries filled in, formulas in **F** to **J**, and the summary in **L5:M11**. **Settings** holds the yellow block size and hourly rate, plus an OK check. **How to use** has short notes. Yellow cells are yours to change and gray cells are formulas.

It prints on US Letter, landscape. To open it in Google Sheets, go to **File > Import > Upload** and select the file. The formulas were checked in LibreOffice Calc 24.2.7.2 only.
