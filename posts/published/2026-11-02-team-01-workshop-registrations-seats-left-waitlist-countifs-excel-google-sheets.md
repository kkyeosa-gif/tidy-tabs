---
threads_url: https://www.threads.com/@tin_ylab/post/DfPYdu0nFo0
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-track-workshop-registrations-and.html
title: How to Track Workshop Registrations and Seats Left in Excel and Google Sheets
labels: excel-formulas, small-business, templates
tested_in: LibreOffice Calc 24.2.7.2 (Linux). Not tested in Excel or Google Sheets. The template was recalculated in 16 edited copies and compared with an independent Python model (9,517 cell comparisons across all five templates built that day, 0 mismatches). Checked: sign-ups equal to and one over the seat count, zero sign-ups, different capitalization, a trailing space in a workshop name, a misspelled workshop, Paid typed with a trailing space, 0 seats, seats typed as text, a duplicate workshop row and blank rows in Signups. The PDF export pages are US Letter.
search_description: Count paid sign-ups, seats left and a waitlist per workshop with COUNTIFS and MAX in Excel and Google Sheets. Free template with a check cell.
image_prompts: A wooden craft table with a half-finished embroidery hoop, a small stack of blank paper name cards and a pencil, soft daylight from a window, no screens, no readable text, no logos
image_alt: Embroidery hoop, blank paper name cards and a pencil on a wooden workshop table
threads: Class says 8 seats and the sign-ups keep coming?\nCOUNTIFS counts the paid ones per workshop, MAX(0,C5-D5) gives seats left without going negative.\nA second count of all sign-ups minus the seats is your waitlist.
---
Count paid sign-ups for a workshop with `=COUNTIFS(Signups!$B$5:$B$204,A5,Signups!$D$5:$D$204,"Paid")`, then show the seats left with `=MAX(0,C5-D5)`. Everyone who signs up beyond the seat count goes on the waitlist, which you get with `=MAX(0,E5-C5)` from a separate count of all sign-ups.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIFS, COUNTIF and MAX are in current versions of Excel and Google Sheets, but neither app was opened for this post.

![Workshop sheet with paid sign-ups, seats left, waitlist count and status for four workshops](images/2026-11-02-team-01-workshop-registrations-seats-left-waitlist-countifs-excel-google-sheets/workshop-seats-left-waitlist-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The template has two tabs. **Workshops** lists each workshop once (name in A, date in B, seats in C, starting in row 5). **Signups** has one row per person, with the header in row 4: name in A, workshop in B, email in C, payment in D.

1. On **Workshops**, type each workshop in **A5:C14**: the name, the date as MM/DD/YYYY, and the number of seats.
2. On **Signups**, type one row per person in **A5:D204**. In column B, use the workshop name exactly as it appears on **Workshops**. In column D, use Paid or Unpaid.
3. In **D5** on **Workshops**, type the paid count and fill down to row 14:

   `=IF(A5="","",COUNTIFS(Signups!$B$5:$B$204,A5,Signups!$D$5:$D$204,"Paid"))`

4. In **E5**, count every sign-up for that workshop, paid or not:

   `=IF(A5="","",COUNTIF(Signups!$B$5:$B$204,A5))`

5. In **F5**, type the seats left:

   `=IF(A5="","",IF(ISNUMBER(C5),MAX(0,C5-D5),"Check seats"))`

6. In **G5**, type the waitlist:

   `=IF(A5="","",IF(ISNUMBER(C5),MAX(0,E5-C5),"Check seats"))`

7. In **H5**, type the status:

   `=IF(A5="","",IF(NOT(ISNUMBER(F5)),"Check seats",IF(F5=0,"Full",IF(F5<=2,"Almost full","Open"))))`

`COUNTIFS` counts the rows where every pair matches, here the workshop name and the word Paid. `MAX(0,...)` stops the seats left from going below zero when a workshop is oversold.

The waitlist counts all sign-ups beyond the seat count, paid or unpaid. An unpaid sign-up does not use a seat, but it still counts toward the waitlist. If you want a different rule, change column E to count only the rows you want.

If a seat count is blank or typed as text, F, G and H show Check seats instead of a wrong number.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-fa83b8e4a0d8) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Neither link was link-checked.

## Example

Sample data is fictional: made up workshops, people and example.com emails.

| Workshop | Date | Seats | Paid | All sign-ups | Seats left | Waitlist | Status |
|---|---|---|---|---|---|---|---|
| Candle Pouring 101 | 11/14/2026 | 8 | 7 | 8 | 1 | 0 | Almost full |
| Intro to Block Printing | 11/21/2026 | 10 | 5 | 6 | 5 | 0 | Open |
| Embroidery Basics | 12/05/2026 | 6 | 6 | 7 | 0 | 1 | Full |
| Wreath Making | 12/12/2026 | 12 | 6 | 7 | 6 | 0 | Open |

Embroidery Basics has 6 seats and 6 paid sign-ups, so it is Full. The seventh person puts it at one on the waitlist. Wreath Making also has 7 sign-ups, but only 6 are paid and it has 12 seats, so nobody waits.

The summary in **J1:K7** adds it up: 4 workshops, 36 seats, 24 paid, 12 seats left and 1 person on the waitlist.

This sheet counts people and seats. It says nothing about refunds, cancellations or waivers, so keep those terms wherever you already state them.

## Troubleshooting

### A sign-up is not counted anywhere
The workshop name in **Signups** does not match **Workshops**. A trailing space (`Wreath Making `) or a typo (`Candel Pouring 101`) is skipped by every count. Cell **K6** counts these rows, and the check cell **K7** shows Check. In the test, two trailing-space rows made K6 read 2. Retype the name or pick it from the dropdown.

### A paid person shows as unpaid, and the check cell says OK
Payment typed as `Paid ` with a trailing space is not counted as paid, and **K7** does not warn about it. Pick Paid from the dropdown in column D instead of typing it. Capitalization does not matter: `paid` and `PAID` are counted.

### Seats left shows 0 but the check cell says Check
More people have paid than there are seats. In the test, 9 paid on 8 seats gave 0 left, a waitlist of 1 and Check in **K7**. Raise the seat count, or move someone to another workshop.

### Seats left shows Check seats
The seat count in column C is blank or typed as text. Type it as a plain number such as 8.

### The totals are too high
A workshop listed twice on **Workshops** makes both rows count every matching sign-up. Keep each workshop to one row.

## Template

[Download the .xlsx](templates/tidy-tabs-workshop-registrations.xlsx)

The file has three tabs. **Workshops** has 4 fictional workshops in yellow cells **A5:C14** (10 slots), with gray formula columns for paid, all sign-ups, seats left, waitlist and status, plus the summary in **J1:K7**. **Signups** has 28 fictional rows in yellow cells **A5:D204** (200 slots), with dropdowns for the workshop and for Paid or Unpaid. **How to use** has notes on each formula.

The **CHECK** cell, **K7**, says OK, or Check if a sign-up names a workshop that is not on the list, a seat count is not a number, or a workshop has more paid sign-ups than seats. Fix every Check before you trust a count.

Limits: 10 workshops and 200 sign-ups. In the test, 200 sign-ups and 10 workshops matched. More was not tested. COUNTIF treats `*` and `?` in a workshop name as wildcards, and that was not tested. **Workshops** prints landscape on US Letter. To open the file in Google Sheets, go to **File > Import > Upload** and select it.
