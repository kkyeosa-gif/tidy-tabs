---
threads_url: https://www.threads.com/@tin_ylab/post/DfP6AVXiVqF
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-show-your-3-biggest-unpaid.html
title: How to Show Your 3 Biggest Unpaid Invoices in Excel and Google Sheets
labels: excel-formulas, invoicing, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model across 19 workbooks and 4,123 comparisons, with 0 mismatches. Checked: fewer than 3 open invoices, all invoices paid, a tie for first place, a tie across ranks 3 and 4, status typed as paid and PAID, a status with a trailing space, an amount typed as text, a blank amount, a negative amount, a missing invoice number, a blank row in the middle and 100 random rows. Excel and Google Sheets were not used.
search_description: List your top 3 unpaid invoices with LARGE, INDEX and MATCH in Excel and Google Sheets, ties handled. Free invoice template with a check cell.
image_prompts: A stack of paper invoices on a wooden desk with a red pencil lying across the top sheet and a rubber stamp beside it, soft window light, no screens, no readable text, no logos
image_alt: Paper invoices with a red pencil across the top and a rubber stamp beside them
threads: Which unpaid invoice should you chase first?\nLARGE finds the biggest open amount, then INDEX and MATCH bring back the invoice number.\nTwo invoices for $1,250.00? A helper column subtracts a tiny row-based amount so neither gets listed twice.
---
Find the largest open amount with `=LARGE(Invoices!$G$5:$G$104,1)`, change the 1 to 2 and 3 for the next two, then bring back the invoice number with `=INDEX(Invoices!$A$5:$A$104,MATCH(B5,Invoices!$G$5:$G$104,0))`. Column G is the open amount minus a tiny row-based number, so two invoices for the same amount are not mixed up.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every result below was recalculated and compared with a Python model. LARGE, INDEX, MATCH and IFERROR are in every current version of Excel and Google Sheets, but neither app was opened for this post.

![Invoice list beside a Top 3 table of largest open invoices and their share](images/2026-11-02-team-03-three-biggest-unpaid-invoices-large-match-excel-google-sheets/top-3-unpaid-invoices-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Invoices** has a header in row 4 and one invoice per row from row 5 to row 104: invoice number in A, client in B, amount in C, due date in D, status in E.

1. On **Invoices**, type one invoice per row in **A5:E104**. Use Paid, Sent or Overdue for the status.
2. In **F5**, type `=IF(OR(A5="",E5="Paid",NOT(ISNUMBER(C5))),"",C5)`. This is the open amount. It stays blank for paid invoices, empty rows and amounts typed as text.
3. In **G5**, type `=IF(F5="","",F5-ROW()/1000000)`. This is the tie-safe key. Fill both F and G down to row 104.
4. On **Top 3**, type 1, 2 and 3 in **A5:A7**.
5. In **B5**, type `=IFERROR(LARGE(Invoices!$G$5:$G$104,A5),"")` and fill down to row 7.
6. In **C5**, type `=IF(B5="","",INDEX(Invoices!$A$5:$A$104,MATCH(B5,Invoices!$G$5:$G$104,0)))`.
7. In **D5**, use the same formula with `$B` in place of `$A` to get the client.
8. In **E5**, use the same formula with `Invoices!$F$5:$F$104` to get the amount. Fill C through E down to row 7.

`LARGE` returns the nth biggest number in a range, but not whose it is. `MATCH` finds the row where that number sits, with the `0` meaning exact match. `INDEX` then returns the invoice number, client or amount from that same row.

## Why the helper column

Say two open invoices are both $1,250.00. `LARGE` on the plain amount returns $1,250.00 for rank 2 and again for rank 3. `MATCH` finds the first $1,250.00 both times, so one invoice would show twice and the other never.

Column G fixes that. It subtracts the row number divided by 1,000,000, so every key is a little different. The invoice higher in the list gets the larger key and ranks first.

You never see that key in the result. The amount on **Top 3** comes from column F, so it is the exact amount you typed.

Help pages: Microsoft's [LARGE function](https://support.microsoft.com/en-us/office/large-function-3af0af19-1190-42bb-bb8b-01672ec00a64), [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The links were not link-checked.

## Example

Sample data is fictional: made up clients and invoices for a small service business.

The sheet has 18 invoices. Six are Paid, so 12 are open, and they add up to $9,845.75. The top 3:

| Rank | Invoice # | Client | Amount |
|---|---|---|---|
| 1 | INV-2045 | Bluebird Bakery | $2,400.00 |
| 2 | INV-2042 | Pine & Co. Realty | $1,250.00 |
| 3 | INV-2047 | Oakline Studio | $1,250.00 |

Ranks 2 and 3 are a tie, and both invoices show up.

Next to the table, the template adds a summary:

| Item | Cell | Result |
|---|---|---|
| Open invoices | H4 | 12 |
| Total open | H5 | $9,845.75 |
| Top 3 total | H6 | $4,900.00 |
| Top 3 share of open | H7 | 49.8% |
| CHECK | H8 | OK |

The share formula in **H7** is `=IF(H5<=0,"",ROUND(H6/H5,4))`. Here three invoices are about half of what you are owed. This is arithmetic on your own numbers, not collections advice.

## Troubleshooting

### Fewer than 3 rows fill in
That is expected when fewer than 3 invoices are open. In the tests, two open invoices left row 3 blank and the share read 100%. One open invoice left rows 2 and 3 blank. With every invoice paid, all three rows were blank and the total was $0.00.

### An invoice is missing from the list
Check the amount. A number typed as text, such as `640` that is left-aligned, is left out of the open amount, and **CHECK** still says OK, so nothing warns you. Retype it as a number. A blank amount or a missing invoice number also leaves the row out. The missing invoice number does turn **CHECK** to Check.

### A paid invoice still appears
The status must be exactly Paid. In the tests, `paid` and `PAID` counted as Paid, because the comparison ignores capitalization. A status with a trailing space, `Paid `, counted as open. A blank status is open too.

### A credit shows up in the top 3
A negative amount counts as open. In one test with four open invoices ($700.00, $100.00, a credit of $10.00 and a credit of $250.00), third place was the $10.00 credit. The share can also go over 100% (146.3% in that test), because credits lower the total. If you do not want credits ranked, enter them on a separate sheet.

### Three invoices tie for third place
Only the first one in the list is shown. In the test, three $300.00 invoices competed for third and the first, INV-3, was listed.

## Template

[Download the .xlsx](templates/tidy-tabs-top-unpaid-invoices.xlsx)

The file has three tabs. **Invoices** has 18 fictional invoices in yellow cells **A5:E104** (100 slots), with the open amount and tie-safe key in grey formula cells. **Top 3** has the ranks in yellow cells **A5:A7**, the invoice number, client and amount, and the summary in **G4:H8**. **How to use** has notes on each formula.

The **CHECK** cell, **H8**, says OK when the open amounts match a separate `SUMIF` over every row that is not Paid, and the three amounts are in order. Otherwise it says Check. It does not catch an amount typed as text.

Limits: 100 invoices and 3 ranks. The key subtracts 0.000001 per row, so amounts that differ by less than a cent could rank in the wrong order. That was not tested, and neither were amounts above about $1 billion. **Top 3** prints landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
