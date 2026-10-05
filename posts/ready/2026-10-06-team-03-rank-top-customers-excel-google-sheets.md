---
title: How to Rank Customers by Sales in Excel and Google Sheets
labels: excel-formulas, freelancers, small-business-bookkeeping
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps follow Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Rank customers by yearly sales with =RANK(B2,$B$2:$B$11,0), then break ties with COUNTIF so every row gets its own number. Free .xlsx template.
image_prompts: A small business owner's stack of paper client folders with colored tabs fanned out on a wooden desk, a brass bell and a handwritten-looking leaderboard card with blurred unreadable lines, a pencil, no screens, no logos
image_alt: Fanned stack of client folders with colored tabs, a brass bell, and a pencil
threads: Two clients tie at $3,920.00 and RANK gives both a 2. There is no rank 3.\nAdd +COUNTIF($B$2:B2,B2)-1 to =RANK(B2,$B$2:$B$11,0) and every row gets its own number.\nThe template also pulls a Top 3 list with INDEX and MATCH.
---
Type `=RANK(B2,$B$2:$B$11,0)` and fill it down to rank each customer from highest total to lowest. Tied totals share a rank, so add `+COUNTIF($B$2:B2,B2)-1` to give every row its own number.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every rank below was recalculated and checked against a hand count. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Customers sheet listing ten clients with total sales, a shared rank, a unique rank, and a Top 3 block](images/2026-10-06-team-03-rank-top-customers-excel-google-sheets/rank-top-customers-template.png) *LibreOffice Calc 24.2 PDF export of the Customers sheet with sample data (fake).*

## Steps

The template keeps one row per customer with the yearly total already typed in. Nothing here needs SUMIFS or a second tab.

1. In row 1, use these headers: Customer, Total sales, Rank (ties share), Unique rank (ties broken).
2. Type each customer name in column A and the total in column B as a plain number, such as 4850. Format column B as currency.
3. In **C2**, type `=RANK(B2,$B$2:$B$11,0)` and fill it down to row 11. The 0 ranks the biggest total as 1. Use 1 instead of 0 to rank the smallest total as 1.
4. Look at the result. Customers with the same total get the same rank, and the next rank is skipped. Two customers at rank 2 means nobody has rank 3.
5. In **D2**, type `=C2+COUNTIF($B$2:B2,B2)-1` and fill it down. The COUNTIF range starts at **$B$2** and ends at the current row, so it counts how many times this total has appeared so far. The first customer with a total keeps the rank. The next one with the same total gets rank + 1.
6. For a short leaderboard, type 1, 2, and 3 in **F2:F4**. In **G2**, type `=INDEX($A$2:$A$11,MATCH(F2,$D$2:$D$11,0))` and in **H2**, type `=INDEX($B$2:$B$11,MATCH(F2,$D$2:$D$11,0))`. Fill both down. Use the unique rank in MATCH so it never searches for a rank that was skipped.
7. Keep the **$** signs on the ranges. They stop the list from sliding as you fill down. If you add customers, extend $B$2:$B$11 in every formula.

### In Excel

The formulas work as typed. Microsoft documents the function on its [RANK function](https://support.microsoft.com/en-us/office/rank-function-6a2fc49d-1831-4a03-9d8c-c279cf99f723) page and the tie-break helper on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) page.

### In Google Sheets

The same formulas work as typed. Google documents the tie-break helper on its [COUNTIF](https://support.google.com/docs/answer/3093480) page. I did not find a RANK page to link, and RANK was not tried in Sheets for this post.

## Example

Sample data below is fictional. It is ten small-business clients of a freelancer, with yearly totals typed in.

| Customer | Total sales | Rank (ties share) | Unique rank (ties broken) |
|---|---|---|---|
| Harbor Coffee Co. | $4,850.00 | 1 | 1 |
| Maple Street Bakery | $3,920.00 | 2 | 2 |
| Oak & Ember Candles | $3,920.00 | 2 | 3 |
| Riverside Yoga | $2,780.00 | 5 | 5 |
| Juniper Florist | $2,150.00 | 6 | 6 |
| Bluebird Books | $2,150.00 | 6 | 7 |
| Cedar Hardware | $1,640.00 | 8 | 8 |
| Lakeview Dental Office | $3,100.00 | 4 | 4 |
| Pine Grove Garden | $980.00 | 10 | 10 |
| North End Tailors | $1,275.00 | 9 | 9 |

The rows are not sorted, and RANK does not need them to be. Lakeview Dental Office sits in row 9 and still gets rank 4, because $3,100.00 is the fourth largest total.

There are two deliberate ties. Maple Street Bakery and Oak & Ember Candles both total $3,920.00, so RANK gives each a 2 and skips 3. The unique rank turns that into 2 and 3. Juniper Florist and Bluebird Books tie at $2,150.00, so RANK gives both a 6 and the unique rank gives 6 and 7.

The Top 3 block then shows:

| Rank | Customer | Total sales |
|---|---|---|
| 1 | Harbor Coffee Co. | $4,850.00 |
| 2 | Maple Street Bakery | $3,920.00 |
| 3 | Oak & Ember Candles | $3,920.00 |

## Troubleshooting

### Two customers share a rank and the next number is missing
That is how RANK works with tied totals. Use the unique rank column in **D**, or add `+COUNTIF($B$2:B2,B2)-1` to your RANK formula, so each row gets its own number.

### The Top 3 block shows #N/A
MATCH is looking for a rank that does not exist. This happens when it searches the shared rank column and the rank was skipped, such as 3 in the sample. Point MATCH at the unique rank column, as in step 6. A #N/A from a typed rank outside 1 to 10 was not reproduced for this post.

### Ranks change after I add a customer
The ranges end at row 11. A customer in row 12 is not counted, and its own rank formula would show the wrong number. Extend $B$2:$B$11 to cover the new row in every RANK, COUNTIF, and INDEX formula.

### Every unique rank is wrong after I fill down
Check the **$** signs. The COUNTIF range must be `$B$2:B2` with the first row locked and the second row free. If both are locked, every row counts all ten totals. If neither is locked, the range slides and misses earlier rows. The sample uses the form shown in step 5.

## Template

[Download the .xlsx](templates/tidy-tabs-rank-top-customers.xlsx)

The file has two tabs. **Customers** holds the ten sample clients, the shared rank, the unique rank, and the Top 3 block in **F1:H4**. **How to use** has short fill-in notes.

The customers and totals are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).
