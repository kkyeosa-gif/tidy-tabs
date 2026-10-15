---
title: How to Find Rows With Missing Information in a Customer List in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against a Python blank count, including a space-only cell, a cell holding ="", an all-blank row in the middle and an empty list. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Find missing info in a customer list: =COUNTBLANK(A2:E2), then =IF(F2=0,"Complete","Missing"). Free .xlsx template.
image_prompts: Index card box with a few cards pulled halfway out beside a paper address label sheet with several blank spaces and a pen on a wooden counter, soft window light, no screens, no readable text, no logos
image_alt: Index card box with cards pulled out beside a half-empty label sheet
threads: Missing emails, phones and ZIPs add up quietly.\n=COUNTBLANK(A2:E2) counts the empty cells in each row, and 8 of the 12 sample rows come out Missing.\nA cell holding only a space counts as filled, so it hides a gap.
---
Count the empty cells in each row with `=COUNTBLANK(A2:E2)`, then flag the row with `=IF(F2=0,"Complete","Missing")`. A cell that holds only a space does not count as blank.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Customer table with Complete and Missing flags and the first empty field named](images/2026-10-16-team-05-find-incomplete-customer-rows-countblank-excel-google-sheets/incomplete-customer-rows-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The sheet has one customer per row: name in **A**, email in **B**, phone in **C**, street in **D** and ZIP in **E**.

1. On a tab named **Customers**, type these headers in row 1: Name, Email, Phone, Street, ZIP, Blanks, Status, First missing.
2. Format column **E** as plain text before you type ZIP codes, so 02108 keeps its leading zero.
3. In **F2**, type `=IF(COUNTA(A2:E2)=0,"",COUNTBLANK(A2:E2))`. This counts the empty cells in the row, and leaves a row with no data at all blank.
4. In **G2**, type `=IF(F2="","",IF(F2=0,"Complete","Missing"))`.
5. In **H2**, type `=IF(OR(F2="",F2=0),"",IF(A2="","Name",IF(B2="","Email",IF(C2="","Phone",IF(D2="","Street","ZIP")))))`. This names the first empty field, checking left to right.
6. Select **F2:H2** and drag the fill handle down to row 30.
7. For a summary, put `=COUNT(F2:F30)` in **K2** (rows in use), `=COUNTIF(G2:G30,"Complete")` in **K3** and `=COUNTIF(G2:G30,"Missing")` in **K4**.
8. To count customers with no ZIP, put `=SUMPRODUCT((F2:F30<>"")*(E2:E30=""))` in **K5**.

Then filter column **G** to Missing and work down the list.

### In Excel

The formulas work as typed. Microsoft describes them on its [COUNTBLANK function](https://support.microsoft.com/en-us/office/countblank-function-6a92d772-675c-4bee-b346-24af6bd3ac22) and [COUNTA function](https://support.microsoft.com/en-us/office/counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509) pages. To filter, use **Data > Filter**.

### In Google Sheets

The same formulas should work as typed, and **Data > Create a filter** gives you the same Missing view. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small shop's customer list.

| Name | Email | Phone | Street | ZIP | Blanks | Status | First missing |
|---|---|---|---|---|---|---|---|
| Ava Thompson | has email | has phone | has street | 02108 | 0 | Complete | |
| Ben Ortiz | | has phone | has street | 07030 | 1 | Missing | Email |
| Carla Nguyen | has email | | has street | 02210 | 1 | Missing | Phone |
| Dev Patel | has email | has phone | | has ZIP | 1 | Missing | Street |
| Erin Walsh | has email | has phone | has street | | 1 | Missing | ZIP |
| Jon Baker | | | | | 4 | Missing | Email |

The table shows 6 of the 12 sample rows, with the field contents shortened. The full list gives K2 = 12, K3 = 4 and K4 = 8. K5 shows 2 customers with no ZIP, Erin Walsh and Jon Baker.

A row with no name but an email, such as the one for hal.k, shows 1 blank and Name as the first missing field.

## Troubleshooting

### A row with a hidden space shows Complete
A cell that holds only a space is not empty, so COUNTBLANK skips it. In the recalculated copy, a single space in B3 turned that row from 1 blank to 0 and Complete, and the summary moved to K3 = 5 and K4 = 7. Select the cell and press **Delete** to clear it.

### A formula that returns "" counts as blank
In LibreOffice Calc 24.2, a cell holding `=""` was counted as blank, so F3 showed 1 and Missing. The check cell in **K6** then showed CHECK, because COUNTA treats a formula cell as filled. How Excel and Google Sheets treat `=""` in COUNTBLANK was not run, so test one cell before you rely on it.

### A fully empty row shows nothing
This is intended. The COUNTA test in column **F** treats a row with no data as unused, so it gets no 5 and no Missing flag. In the recalculated copy, an all-blank row in the middle stayed blank, and the complete row after it still showed 0 and Complete. Type one value in the row and it starts counting.

### The ZIP lost its leading zero
A ZIP typed into a number cell becomes 7030, not 07030. The template formats column **E** as text, so type or paste ZIP codes there as text.

### Every blank looks fine but the data is wrong
COUNTBLANK counts empty cells only. It does not check that an email has an at sign or that a ZIP has five digits. Use it to find gaps, then review what is filled in.

## Limits

The first missing field follows the fixed order Name, Email, Phone, Street, ZIP. Change the order of the IF tests in **H2** if you care about a different field first. The ranges stop at row 30, so extend them for a longer list.

## Template

[Download the .xlsx](templates/tidy-tabs-incomplete-rows.xlsx)

The file has two tabs. **Customers** holds the 12 sample rows in yellow input cells, formulas in **F** to **H** filled to row 30, the summary in **K2** to **K6** and a check cell in **K6**. **How to use** has short fill-in notes. **Customers** prints landscape and **How to use** prints portrait, both on US Letter.

The customers are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
