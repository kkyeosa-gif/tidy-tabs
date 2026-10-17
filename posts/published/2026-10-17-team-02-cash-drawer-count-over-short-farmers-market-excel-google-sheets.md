---
threads_url: https://www.threads.com/@tin_ylab/post/Dem7AUcDyhK
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-count-cash-drawer-and-find-over.html
title: How to Count a Cash Drawer and Find Over or Short in Excel and Google Sheets
labels: excel-formulas, small-business, farmers-market
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and checked against Python Decimal arithmetic: the sample drawer, three more fictional drawers (balanced, one cent short, one cent over), blank counts, a float of $0.00, and counts typed as text, negative or fractional. The one cent over case was also run without ROUND. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Count a cash drawer and find over or short with =ROUND(D12-F4,2) in Excel and Google Sheets. Free cash count sheet for a market booth.
image_prompts: Open metal cash box on a folding table at a farmers market, loose bills sorted by size in the tray, a paper-wrapped roll of coins beside a jar with a cloth lid, soft morning light, no screens, no readable text, no logos
image_alt: Open metal cash box with sorted bills and a coin roll on a folding table
threads: Closing the cash box at the market and the count is off?\n=ROUND(D12-F4,2) compares what you counted with the float plus cash sales, and the sample drawer comes out $2.50 short.\nROUND matters: one cent over can show up as 0.00999999999999091 without it.
---
Multiply each denomination by its count, add the amounts with `=SUM(D2:D11)`, then find over or short with `=ROUND(D12-F4,2)`, where F4 is the starting float plus cash sales. A positive result is over, a negative result is short, and 0 is balanced.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Cash count table with amounts per bill and coin, the expected cash and a Short status](images/2026-10-17-team-02-cash-drawer-count-over-short-farmers-market-excel-google-sheets/cash-drawer-count-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The sheet has one row per bill or coin. The value is in **B**, the count you type is in **C**, and the amount is in **D**.

1. On a tab named **Count**, type these headers in row 1: Denomination, Value, Count, Amount.
2. In **A2:A11** list $100 bill, $50 bill, $20 bill, $10 bill, $5 bill, $1 bill, Quarter, Dime, Nickel, Penny.
3. In **B2:B11** type the values: 100, 50, 20, 10, 5, 1, 0.25, 0.10, 0.05, 0.01. Format the column as currency.
4. In **D2**, type `=IF(C2="","",IF(NOT(ISNUMBER(C2)),"Check",IF(AND(C2>=0,C2=INT(C2)),B2*C2,"Check")))`. It multiplies value by count, and shows Check for a count that is text, negative or a fraction.
5. Drag the fill handle from **D2** down to **D11**.
6. In **A12** type Total counted. In **C12**, type `=SUM(C2:C11)` for the number of bills and coins. In **D12**, type `=SUM(D2:D11)` for the cash you counted.
7. In **E1** type Closing numbers, then in **E2:E3** label two inputs, Starting float and Cash sales for the day. Type the float in **F2** and your cash sales from your own records in **F3**.
8. In **F4**, type `=F2+F3`. This is the cash you expect to find.
9. In **F5**, type `=ROUND(D12-F4,2)`.
10. In **F6**, type `=IF(COUNTIF(D2:D11,"Check")>0,"Check counts",IF(F5=0,"Balanced",IF(F5>0,"Over","Short")))`.
11. In **E4:E8** add the labels Expected cash, Over (+) or short (-), Status, Check: SUMPRODUCT total and Check: both totals agree.
12. For a second opinion, type `=SUMPRODUCT(B2:B11,C2:C11)` in **F7** and `=IF(ROUND(F7-D12,2)=0,"OK","CHECK")` in **F8**. Both totals should agree.

### In Excel

The formulas work as typed. Microsoft describes them on its [SUM function](https://support.microsoft.com/en-us/office/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) pages. To format currency, select the cells and use the Number Format box on the **Home** tab, then choose **Currency**. Microsoft lists the options on its [Available number formats in Excel](https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2) page.

### In Google Sheets

The same formulas should work as typed. For currency, the usual menu path is **Format > Number > Currency**. Neither the formulas nor that menu path were opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is one market booth's cash box at closing, with a $150.00 float and $161.85 in cash sales.

| Denomination | Value | Count | Amount |
|---|---|---|---|
| $20 bill | $20.00 | 6 | $120.00 |
| $10 bill | $10.00 | 9 | $90.00 |
| $5 bill | $5.00 | 14 | $70.00 |
| $1 bill | $1.00 | 23 | $23.00 |
| Quarter | $0.25 | 18 | $4.50 |
| Dime | $0.10 | 12 | $1.20 |
| Nickel | $0.05 | 9 | $0.45 |
| Penny | $0.01 | 20 | $0.20 |

The $100 and $50 rows are 0 and left out of the table. The counted total in **D12** is $309.35 across 111 bills and coins. Expected cash in **F4** is $311.85, so **F5** shows -2.50 and **F6** says Short. **F7** shows 309.35 and **F8** says OK.

If cash sales had been $159.35, the same drawer would show 0 and Balanced.

Three more fictional drawers matched the Python figures. A float of $200.00 with $329.10 in sales and $529.10 counted is Balanced. A float of $100.00 with $220.52 in sales and $320.51 counted is Short by 0.01. The same counts with $220.50 in sales is Over by 0.01.

## Troubleshooting

### One cent over shows a tiny odd number
Without ROUND, `=D12-F4` returned 0.00999999999999091 for the one cent over drawer in LibreOffice Calc, and a test like `=0` came out false. Decimals such as 0.10 are stored in binary, so small errors appear. The ROUND in **F5** returned 0.01. The balanced drawers tried here came out exactly 0 even without ROUND, so the problem is not universal. Excel was not run.

### A count typed as text shows Check
In the recalculated copy, typing six as a word in the $20 row put Check in **D**, and **F6** said Check counts. The total in **D12** skipped that row, so it fell from $309.35 to $189.35, which is $120.00 low, until the count was fixed. Retype the count as a number.

### A negative or fractional count shows Check
A count of -6 or 6.5 gets Check in column **D**. In both cases **F8** showed CHECK, because SUMPRODUCT still includes those counts while SUM of the amounts does not. Fix the count and both totals match again.

### Blank counts look like a missing drawer
A blank count is treated as none. With every count blank, a $150.00 float and no sales, **D12** was 0 and **F5** showed -150 and Short. That is correct arithmetic, so check that you typed the counts.

### Float or sales typed as text
Text in **F2** or **F3** makes **F4** an error, and the template does not guard against it. Retype the amount as a plain number.

## Limits

This is counts and arithmetic only, not accounting or bank reconciliation advice. It handles one drawer per file, and rolled coins and checks are not covered.

## Template

[Download the .xlsx](templates/tidy-tabs-cash-drawer-count.xlsx)

The file has two tabs. **Count** has the denomination table, counts in yellow input cells in **C2:C11**, the float in **F2**, cash sales in **F3**, formulas in **D**, **C12**, **D12** and **F4** to **F8**, and highlighting on **F6** for Over, Short or Check counts. **How to use** has short fill-in notes. Both tabs print portrait on US Letter.

The counts and sales are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file. To count another drawer, save a copy or clear column **C** and **F3**.
