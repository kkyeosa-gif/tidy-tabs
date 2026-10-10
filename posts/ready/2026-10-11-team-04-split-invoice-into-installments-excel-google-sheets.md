---
title: How to Split an Invoice Into Equal Monthly Payments in Excel and Google Sheets
labels: excel-formulas, invoicing, freelancers
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the payment amounts and due dates were checked against Python math and the final balance was $0.00. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Split an invoice into equal monthly payments with =ROUND(B1/B2,2) and a last payment of =B1-B4*(B2-1) so cents add up exactly. Free .xlsx template.
image_prompts: Freelancer's wooden desk with a paper invoice face down, a small stack of coins and a wall calendar page with a few days circled in pencil, soft window light, no screens, no readable text, no logos
image_alt: Face down paper invoice beside a coin stack and a calendar with penciled circles
threads: $2,500.00 invoice, 6 payments, and the cents refuse to add up?\nRound the regular payment with =ROUND(B1/B2,2), then make the last one =B1-B4*(B2-1).\nFive payments of $416.67 and a last one of $416.65 total exactly $2,500.00.
---
Divide with `=ROUND(B1/B2,2)` for the regular payment, then make the last payment `=B1-B4*(B2-1)` so the cents left over land there and the payments add up to the invoice exactly. A $2,500.00 invoice in 6 payments is five payments of $416.67 and a last one of $416.65.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python math. The Excel and Google Sheets steps use the same functions and link to their help pages. Neither Excel nor Google Sheets was opened for this post.

![Payment plan splitting a 2,500 dollar invoice into six monthly payments](images/2026-10-11-team-04-split-invoice-into-installments-excel-google-sheets/invoice-installments-payment-plan-template.png) *Render of the Payment plan tab of the template, a PDF export from LibreOffice Calc 24.2 using fictional sample data*

## Steps

Plain division fails because most invoices do not split into whole cents. $2,500.00 divided by 6 is $416.666..., and six payments of $416.67 would add up to $2,500.02. Rounding every payment and letting the last one absorb the difference fixes that.

1. In **A1:A3**, type these labels: Invoice total, Number of payments, First payment date.
2. In **B1**, type the invoice total, for example 2500. In **B2**, type the number of payments, for example 6. In **B3**, type the first due date, for example 11/01/2026.
3. In **A4** and **A5**, type the labels Regular payment and Last payment.
4. In **B4**, type `=ROUND(B1/B2,2)`. This is the invoice divided by the number of payments, rounded to cents.
5. In **B5**, type `=B1-B4*(B2-1)`. This is the invoice minus all the regular payments, so it holds the leftover cents.
6. In row 7, type these headers in A to D: Payment #, Due date, Amount, Balance after.
7. In **A8:A19**, type the numbers 1 to 12.
8. In **B8**, type `=IF(A8>$B$2,"",EDATE($B$3,A8-1))`. Each row moves the first date forward one month.
9. In **C8**, type `=IF(A8>$B$2,"",IF(A8=$B$2,$B$5,$B$4))`. It shows the last payment on the final row and the regular payment on every row before it.
10. In **D8**, type `=IF(A8>$B$2,"",$B$1-SUM($C$8:C8))`. This is what is still owed after that payment.
11. Select **B8:D8** and drag the fill handle down to row 19. Rows past your payment count show blank.
12. In **C21**, type `=SUM(C8:C19)`. In **C22**, type `=ROUND(C21-B1,2)`. C22 should read $0.00.
13. Format B1, B4, B5, C8:D19 and C21:C22 as currency with two decimals, and B3 and B8:B19 as dates.

The `$` signs in `$B$2`, `$B$3`, `$B$4` and `$B$5` keep those inputs fixed when you fill down. `SUM($C$8:C8)` grows by one row each time, so the balance counts every payment so far.

`EDATE` keeps the day of the month. When a month is too short it uses the last day instead, so a first date of 01/31/2026 gives 02/28/2026 and then 03/31/2026.

In the Tidy Tabs template, the three input cells are yellow and cell B2 only accepts a whole number from 1 to 12, because the plan has 12 rows.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) page and its [EDATE function](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) page.

### In Google Sheets

The same formulas should work as typed. Google describes the rounding function on its [ROUND](https://support.google.com/docs/answer/3093440) help page. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a $2,500.00 invoice split into 6 payments starting 11/01/2026.

| Payment # | Due date | Amount | Balance after |
|---|---|---|---|
| 1 | 11/01/2026 | $416.67 | $2,083.33 |
| 2 | 12/01/2026 | $416.67 | $1,666.66 |
| 3 | 01/01/2027 | $416.67 | $1,249.99 |
| 4 | 02/01/2027 | $416.67 | $833.32 |
| 5 | 03/01/2027 | $416.67 | $416.65 |
| 6 | 04/01/2027 | $416.65 | $0.00 |

Check it by hand: $416.67 times 5 is $2,083.35, and $2,500.00 minus $2,083.35 is $416.65. The total of the payments is $2,500.00 and the difference cell shows $0.00.

Other splits checked in the recalculated file:

| Invoice | Payments | Regular payment | Last payment |
|---|---|---|---|
| $1,000.00 | 3 | $333.33 | $333.34 |
| $1,000.00 | 12 | $83.33 | $83.37 |
| $100.00 | 7 | $14.29 | $14.26 |
| $750.00 | 1 | $750.00 | $750.00 |

With a first date of 01/31/2026 and 4 payments, the due dates came out as 01/31/2026, 02/28/2026, 03/31/2026 and 04/30/2026.

## Troubleshooting

### The payments add up to a few cents more or less than the invoice
The last payment is probably typed in as a number or uses `B4` instead of the leftover formula. Put `=B1-B4*(B2-1)` in B5 and make the final row use it. Then C22 should read $0.00.

### The plan breaks when you enter more than 12 payments
The plan has only 12 rows. In a copy of the template, typing 13 into B2 gave a difference of -$100.00 in C22, because only 12 rows exist. Keep B2 between 1 and 12, or fill the formulas down further and widen the sums. The template's validation rule rejects other numbers when you type them, but that rule was not tried by typing in LibreOffice.

### Rows past my payment count still show a number in column A
That is expected. Column A holds the numbers 1 to 12 as plain values, and only columns B to D go blank. Hide the unused rows if you do not want them on a printout.

### Due dates look like numbers such as 46327
Cell B3 or the date column is formatted as General. Select the cells and apply a date format such as MM/DD/YYYY.

### The due date moves from the 31st to the 28th
That is how `EDATE` handles short months. A first date of 01/31/2026 gives 02/28/2026, then 03/31/2026. If your client expects a fixed day such as the 1st, start on the 1st.

## Template

[Download the .xlsx](templates/tidy-tabs-invoice-installments.xlsx)

The file has two tabs. **Payment plan** has three yellow inputs (invoice total in B1, number of payments in B2, first payment date in B3), the regular and last payment in B4 and B5, a 12 row schedule in A8:D19, and a total and difference check in C21 and C22. **How to use** has short fill-in notes. The template prints on one US Letter page wide.

It does arithmetic only. It adds no interest, fees or late charges, so follow the written agreement with your client. The months and amounts are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
