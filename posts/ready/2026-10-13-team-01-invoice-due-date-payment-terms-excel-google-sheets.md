---
title: How to Calculate an Invoice Due Date From Payment Terms in Excel and Google Sheets
labels: excel-formulas, invoicing, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python date math, including an unknown terms name, lowercase terms, a missing date and a moved as-of date. Excel and Google Sheets were not opened; the two help links are from Microsoft and Google.
search_description: Add the terms days to the invoice date: =C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE). 09/01/2026 on Net 30 is due 10/01/2026. Free .xlsx with status.
image_prompts: Stack of paper invoices in a wire tray beside a desk calendar with a few days circled and a rubber date stamp on an ink pad, soft window light, no screens, no readable text, no logos
image_alt: Paper invoices in a wire tray next to a circled desk calendar and a date stamp
threads: Invoice says Net 30 and you are counting days on a calendar?\nAdd the terms to the invoice date: =C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE).\n09/01/2026 on Net 30 is due 10/01/2026, and a status column flags what is Past due.
---
Add the days in the terms to the invoice date with `=C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE)`, where C2 is the invoice date and J2:K6 lists each terms name next to its number of days. An invoice dated 09/01/2026 on Net 30 terms is due 10/01/2026.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python date math. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Invoice due dates in a table with Net 30 and Net 15 terms, a Due date column and Past due status flags](images/2026-10-13-team-01-invoice-due-date-payment-terms-excel-google-sheets/invoice-due-date-payment-terms-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: keep a small table that turns each terms name into a number of days. Look the name up, add the days to the invoice date, then compare the due date to today.

1. In row 1, type these headers: Invoice in **A**, a detail of your choice in **B** (such as client), Invoice date in **C**, Terms in **D**, Days in **E**, Due date in **F**, Days until due in **G**, Status in **H**.
2. In **J1:K1**, type Terms and Days. In **J2:K6**, list each terms name and its days: Due on receipt 0, Net 15 15, Net 30 30, Net 45 45, Net 60 60.
3. In **J8**, type As-of date. In **K8**, type a real date such as 10/13/2026. Type `=TODAY()` instead if you want it to update every day.
4. Enter each invoice in rows 2 and down. Type the invoice date in **C** as a real date, and pick or type the terms name in **D**, spelled like the list in J2:J6.
5. In **E2**, type `=IF(D2="","",IFERROR(VLOOKUP(D2,$J$2:$K$6,2,FALSE),"Check terms"))`. This returns the days for the terms, or a warning if the name is not in the table.
6. In **F2**, type `=IF(OR(C2="",NOT(ISNUMBER(E2))),"",C2+E2)`. This is the due date. Format it as a date.
7. In **G2**, type `=IF(F2="","",F2-$K$8)`. Negative means the invoice is already late.
8. In **H2**, type `=IF(G2="","",IF(G2<0,"Past due",IF(G2=0,"Due today","Open")))`.
9. Select **E2:H2** and fill the formulas down as far as you have invoices. The formulas return blanks on empty rows, so you can fill past the last invoice.
10. To count what is outstanding, type `=COUNTIF(H2:H20,"Open")+COUNTIF(H2:H20,"Due today")` in **K10**. In **K11**, type `=COUNTIF(H2:H20,"Past due")`.

The `FALSE` at the end of VLOOKUP forces an exact match on the terms name. The `$` signs keep the terms table fixed when you fill down.

### In Excel

The formulas work as typed. Microsoft explains the arguments on its [VLOOKUP function](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1) page. To add a terms dropdown in **D**, use **Data > Data Validation**, allow List, and point the source at **J2:J6**.

### In Google Sheets

The same formulas should work as typed. Google documents the lookup on its [VLOOKUP help page](https://support.google.com/docs/answer/3093318). For a dropdown in **D**, use **Data > Data validation** and choose Dropdown (from a range) with **J2:J6**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. The as-of date in K8 is typed as 10/13/2026 so the numbers stay put.

| Invoice | Invoice date | Terms | Due date | Days until due | Status |
|---|---|---|---|---|---|
| INV-1001 | 09/01/2026 | Net 30 | 10/01/2026 | -12 | Past due |
| INV-1002 | 09/15/2026 | Net 15 | 09/30/2026 | -13 | Past due |
| INV-1003 | 09/28/2026 | Net 45 | 11/12/2026 | 30 | Open |
| INV-1004 | 10/05/2026 | Net 30 | 11/04/2026 | 22 | Open |
| INV-1005 | 10/13/2026 | Due on receipt | 10/13/2026 | 0 | Due today |
| INV-1006 | 10/01/2026 | Net 60 | 11/30/2026 | 48 | Open |
| INV-1007 | 08/20/2026 | Net 60 | 10/19/2026 | 6 | Open |
| INV-1008 | 10/10/2026 | Net 15 | 10/25/2026 | 12 | Open |

K10 shows 6 invoices still outstanding and K11 shows 2 past due. All eight due dates matched a Python check.

Change K8 to 11/01/2026 and INV-1001, 1002, 1005, 1007 and 1008 all turn Past due.

## Troubleshooting

### The Due date column is blank and Days says Check terms
The terms name is not in **J2:J6**. In the recalculated file, "Net 20" gave Check terms in **E** and blanks in **F**, **G** and **H**. Add a Net 20 row to the table (extend the lookup range if you run past row 6) or fix the spelling.

### The due date is blank but the terms are fine
The invoice date in **C** is empty. In the recalculated file, a missing date left F, G and H blank on purpose. Type the date.

### The due date is a month off
Dates typed as text can be misread. In the recalculated file, a text "10/01/2026" was converted and gave 10/31/2026 for Net 30. Whether text converts depends on the app and your regional settings, so enter real dates and check that they right-align in the cell.

### Net 30 from 01/31/2026 lands on 03/02/2026
That is correct. The formula adds days, not months. January 31 plus 30 days is March 2, because February has 28 days in 2026.

### A weekend or holiday due date looks wrong
The due date is not moved for weekends or holidays. `WORKDAY` can do that, but your contract decides whether it applies. This sheet also has no early-pay discount terms such as 2/10 net 30 and no end-of-month terms. It does date arithmetic only, so check the terms you actually agreed to.

## Template

[Download the .xlsx](templates/tidy-tabs-invoice-due-dates.xlsx)

The file has two tabs. **Due dates** is set to print landscape on one US Letter page. It holds the eight sample invoices in A2:D9, the terms table in J2:K6, the as-of date in K8, the formulas in E2:H20 and the counts in K10:K11. Terms in column D have a dropdown for rows 2 to 20, and Past due cells turn red. **How to use** has short fill-in notes.

The invoices are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
