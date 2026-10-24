---
threads_url: https://www.threads.com/@tin_ylab/post/De4n-1wGjzH
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-track-prepaid-retainer-hours.html
title: How to Track Prepaid Retainer Hours Remaining for Each Client in Excel and Google Sheets
labels: excel-formulas, freelancers, time-tracking
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 29 edited variants (a client with no log rows, hours used up exactly, hours bought 0, a trailing space in a Log or Clients name, capital letters, text hours, a typo client, a duplicate client row, the 20 percent boundary, negative hours, Log rows 200 and 201). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Track prepaid retainer hours remaining for each client in Excel and Google Sheets. Free template with a time log, hours left and a low hours flag.
image_prompts: A brass hourglass sand timer, a stack of three closed linen notebooks and a sharpened pencil on a wooden home office desk, soft window light, no screens, no readable text, no logos
image_alt: Brass hourglass beside three closed linen notebooks and a pencil on a desk
threads: Which client is about to use up their prepaid hours?\nSUMIFS adds the Log hours for each client, then subtract that from the hours bought.\nIn the sample, Pine Hill Yoga is over by 2.25 hours and Lakeview Dental is flagged Low with 3.5 left. The clients are made up.
---
Find a client's hours left with `=B2-SUMIFS(Log!$C$2:$C$200,Log!$B$2:$B$200,A2)`, where B2 is the hours the client prepaid and A2 is the client name. The Log sheet holds one row per entry, with the client in column B and the hours in column C.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Client list with hours bought, hours used and hours left, one client flagged Over](images/2026-10-24-team-01-client-retainer-hours-remaining-sumifs-excel-google-sheets/retainer-hours-remaining-by-client-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: a retainer is a block of hours the client paid for. Add up the hours you logged for that client, then subtract the total from the hours bought.

1. Make a sheet named **Log** with the headings Date, Client, Hours and Task in **A1:D1**.
2. Add one row per entry from row 2. Type the client name exactly the same way each time, and the hours as a plain number such as 4.25.
3. Make a sheet named **Clients**. Type the client names in **A2:A6** and the hours each one prepaid in **B2:B6**.
4. In **C2**, type `=SUMIFS(Log!$C$2:$C$200,Log!$B$2:$B$200,A2)` and fill it down. This is hours used.
5. In **D2**, type `=B2-C2` and fill it down. This is hours left.

You can skip C2 and put everything in one cell: `=B2-SUMIFS(Log!$C$2:$C$200,Log!$B$2:$B$200,A2)`. The template uses two cells so you can see the hours used.

How the formula reads:

- `Log!$C$2:$C$200` is the hours to add up.
- `Log!$B$2:$B$200` is the column of client names to check.
- `A2` is the name to match. `SUMIFS` adds only the hours on rows where the client equals A2.
- The dollar signs lock the two ranges so they stay put when you fill down.

Both apps document the function on their help pages: [SUMIFS in Excel](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) and [SUMIFS in Google Sheets](https://support.google.com/docs/answer/3238496). Those pages describe the function, not our test.

### What the template adds

The built formulas are longer than the bare ones. **C2** is `=IF(A2="","",SUMIFS(Log!$C$2:$C$200,Log!$B$2:$B$200,A2))` and **D2** is `=IF(OR(A2="",NOT(ISNUMBER(B2))),"",B2-C2)`. Both leave the row blank when the client name is blank, and D2 also stays blank when the hours bought is not a number. The counting core is the same.

**F2** gives a status. It says "Over by" and the number of hours when you have logged more than you sold, "Low" when 20% or less of the hours bought are left, and "OK" otherwise. The 20% line is a rule in the template, not a standard. Change it in the formula if you want a different warning point.

### In Excel

The formulas work as typed. SUMIFS, IF and SUM are old functions, and no newer function is needed.

### In Google Sheets

The same formulas should work as typed.

## Example

Sample data is fictional: five clients of a freelance writer, 17 log entries and 59 hours in total.

| Client | Hours bought | Hours used | Hours left | Status |
|---|---|---|---|---|
| Birch Street Bakery | 20.00 | 12.75 | 7.25 | OK |
| Lakeview Dental Office | 30.00 | 26.50 | 3.50 | Low |
| Pine Hill Yoga | 10.00 | 12.25 | -2.25 | Over by 2.25 hours |
| Cedar & Co. Realty | 15.00 | 7.50 | 7.50 | OK |
| Oak Lane Florist | 12.00 | 0.00 | 12.00 | OK |

Birch Street Bakery has three Log rows: 5.5, 4.25 and 3 hours. They add to 12.75, and 20 minus 12.75 leaves 7.25.

Lakeview is Low because 3.50 hours left is at or below 6 hours, which is 20% of 30. Oak Lane Florist has no Log rows yet, so it shows 0.00 used and all 12.00 hours left.

The summary in **H2:I6** shows 87.00 hours bought, 59.00 hours used, 59.00 hours in the Log, 2 clients Low or Over, and OK in the check cell.

This is arithmetic on the hours you type. It does not price the hours, create an invoice or say what your contract requires.

## Troubleshooting

### Hours left is wrong for one client
Check the spelling of that client in the Log. A trailing space or a typo makes a separate name, and those hours are not added to the client in your list. In our LibreOffice runs, "Pine Hill Yoga " with a trailing space on one row left that row out, and **I6** showed CHECK. Capital letters do not matter: "pine hill YOGA" was counted as the same client.

### Hours are counted twice
The same client is listed on two rows of the Clients sheet. Each row adds the full Log total for that name, so the hours used add up to more than the Log. I3 then no longer matches I4, and I6 shows CHECK. Delete one of the rows.

### The hours I typed are ignored
The hours were typed as text, such as "2h". SUMIFS skips text without an error, so the total looks fine but is short. Retype the cell as a plain number. I6 shows CHECK when it finds one.

### The status says Over by
More hours were logged than the client bought. That is the arithmetic, not an error. A negative hours entry, such as -1.5 for a refund, is subtracted from the total.

### New entries are not counted
The formulas read Log rows 2 to 200. In our runs, row 200 was read and row 201 was not. Extend the range in the SUMIFS if you need more rows.

## Template

[Download the .xlsx](templates/tidy-tabs-retainer-hours.xlsx)

The file has three tabs. **Clients** holds the yellow inputs: client names in **A2:A13** and hours bought in **B2:B13**, for up to 12 clients. Gray cells are formulas: hours used in **C**, hours left in **D**, percent used in **E**, status in **F**, and a summary with the check in **H2:I6**. Over rows turn red and Low rows turn orange. **Log** holds the yellow entries in **A2:D200**. **How to use** has short notes.

The clients and hours are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, and the red and orange formatting was seen in a PDF export, not in a spreadsheet window. The template is US Letter, landscape on the Clients tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
