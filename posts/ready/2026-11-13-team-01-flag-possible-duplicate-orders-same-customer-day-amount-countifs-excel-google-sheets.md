---
title: How to Flag Possible Duplicate Orders by Customer, Date and Amount in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 12 workbooks and 12,084 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Flag orders with the same customer, day and amount using COUNTIFS and INT, and number the copies so the first order stays clear. Free .xlsx template.
image_prompts: Two identical kraft mailer boxes side by side on a home packing table, each with the same folded packing slip tucked under a strip of tape, a roll of packing tape between them, address labels blank and unreadable.
image_alt: Two identical kraft mailer boxes side by side on a packing table
threads: Same customer, same $86.50, one minute apart, and a COUNTIFS on the date and time column still says 1.\nCount on the day instead: =INT(B4) in a helper column, then COUNTIFS on customer, day and amount. Above 1 is a possible duplicate.\nA flag means look first. Customers do order the same thing twice.
---
Flag orders that match another order's customer, day and amount with `=COUNTIFS($C$4:$C$203,C4,$E$4:$E$203,E4,$D$4:$D$203,D4)`, where C is the customer, D the amount and E the order day from `=INT(B4)`; any result above 1 is a possible duplicate. A second `=COUNTIFS($C$4:C4,C4,$E$4:E4,E4,$D$4:D4,D4)` numbers the copies so the first order stays unflagged, and in the made-up sample Harbor Dental's two $86.50 orders one minute apart are marked First of 2 and Possible duplicate.

A flag means "look at this," not "this is a duplicate." A customer can really order the same thing twice in one day, so compare the order numbers and items first.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and recalculated and every result was compared with a Python model (12 workbooks, 12,084 comparisons, 0 mismatches). The functions also exist in Excel and Google Sheets, but neither was opened for this post. The open question there is the `"="&` criterion with `~` escapes inside COUNTIFS.

![Order list with Possible duplicate flags on repeat orders from the same customer and day](images/2026-11-13-team-01-flag-possible-duplicate-orders-same-customer-day-amount-countifs-excel-google-sheets/possible-duplicate-orders-template.png) *Fifteen orders with five possible duplicates in four groups, a repeated order number and a text date. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has an **Orders** tab and a **How to use** tab.

1. Paste your orders into **A4:D203**: order number in A, order date and time in B, customer in C and amount in D.
2. In **E4**, type `=INT(B4)`. INT drops the time, so an order at 9:14 AM and one at 9:15 AM land on the same day.
3. In **F4**, type `=COUNTIFS($C$4:$C$203,C4,$E$4:$E$203,E4,$D$4:$D$203,D4)`. A 1 means the order is the only one like it.
4. In **G4**, type `=COUNTIFS($C$4:C4,C4,$E$4:E4,E4,$D$4:D4,D4)`. Only the first row is locked, so the range grows as you copy down. The first order of a group gets 1, the next gets 2.
5. In **H4**, type `=IF(F4=1,"",IF(G4=1,"First of "&F4,"Possible duplicate"))`.
6. Select **E4:H4** and drag the fill handle down to your last order.

Count on the Order day in column E, not on column B, where the times almost never match.

The day comes from the date and time as they are in your export. If the export is in UTC, two orders late in the evening US time can fall on different UTC days. The file does not convert time zones; see [How to Convert Time Zones for Client Calls](https://tidytabs.blogspot.com/2026/10/how-to-convert-time-zones-for-client.html) to shift the times first.

The template wraps these core formulas. **E4** shows Check when the date is text, the customer is blank, or the amount is text or not in whole cents. **F4** and **G4** match the customer as plain text, so a name with `*`, `?` or `~` in it is not read as a wildcard. **I4** counts how many times each order number appears, which catches a row pasted twice.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef), and Google's [COUNTIFS](https://support.google.com/docs/answer/3256550).

## Example

Sample data is fictional: 15 made-up orders from made-up customers. Here are seven of them.

| Order no. | Order date and time | Customer | Amount | Same cust/day/amt | Copy no. | Flag |
|---|---|---|---|---|---|---|
| #1001 | 11/02/2026 9:14 AM | Harbor Dental | $86.50 | 2 | 1 | First of 2 |
| #1002 | 11/02/2026 9:15 AM | Harbor Dental | $86.50 | 2 | 2 | Possible duplicate |
| #1004 | 11/03/2026 10:05 AM | Jordan Reed | $18.00 | 2 | 1 | First of 2 |
| #1006 | 11/03/2026 6:10 PM | jordan reed | $18.00 | 2 | 2 | Possible duplicate |
| #1009 | 11/05/2026 2:15 PM | Pine & Co. Realty | $250.00 | 1 | 1 | |
| #1010 | 11/05/2026 2:50 PM | Pine & Co. Realty | $25.00 | 1 | 1 | |
| #1014 | 11/07/2026 12:30 PM | Harbor Dental | $86.50 | 1 | 1 | |

Harbor Dental's third $86.50 order is on a different day, so it is not flagged. Pine & Co. Realty ordered twice in one afternoon, but $250.00 and $25.00 are different amounts, so neither row is flagged. Upper and lower case do not matter: jordan reed matches Jordan Reed.

The summary block in **K3:L9** shows 15 orders listed, 5 possible duplicates, 4 groups with more than one order, $292.00 in possible duplicates, 2 rows with a repeated order number, 1 row showing Check and 5 orders with no flag.

## Troubleshooting

### Two different customers are counted as one
A plain COUNTIFS reads `*` and `?` in a name as wildcards. In LibreOffice, the core formula with "A*B Studio" returned 2 because the `*` also matched "AxB Studio". The template's wrapped version returns 1.

### Orders minutes apart are not flagged
The formula counts column B (date and time) instead of the Order day. Counting on B returned 1 for Harbor Dental #1001, where the file returns 2. Point the date part of COUNTIFS at column E.

### An obvious duplicate is not flagged
Extra spaces make a different name. "Harbor Dental " with a trailing space was not flagged next to "Harbor Dental". Clean the names first with [How to Clean Up a Customer List With TRIM and PROPER](https://tidytabs.blogspot.com/2026/10/how-to-clean-up-customer-list-with-trim.html). Amounts also have to match to the cent.

### A row shows Check
The date is text (in the sample, #1015 has 11/08/2026 typed as text), the customer is blank, or the amount is text or has fractions of a cent, such as 12.345. Retype the date or fix the amount.

### The same order number shows First of 2 and Possible duplicate
The row was pasted twice. In the sample, #1007 for Cedar Print Shop appears on two rows, and **Times this order no. is listed** shows 2 in red on both.

## Template

[Download the .xlsx](templates/tidy-tabs-possible-duplicate-orders.xlsx)

The file has two tabs. **Orders** holds your orders in **A4:D203** (Order no., Order date and time, Customer, Amount), the formula columns **E:I** (Order day, Same customer, day and amount, Copy no., Flag, Times this order no. is listed) and the summary in **K3:L9**. **How to use** explains each column. It prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. The Flag cell is red for Check and Possible duplicate, and the order number count is red when it is above 1. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
