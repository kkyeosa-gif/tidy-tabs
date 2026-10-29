---
threads_url: https://www.threads.com/@tin_ylab/post/DfDrObZicFi
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-split-list-into-batches-of-10.html
title: How to Split a List Into Batches of 10 and Assign Helpers in Turn in Excel and Google Sheets
labels: excel-formulas, small-business, lists
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and every cell in rows 5 to 204 was compared with Python divmod. Checked: 45 orders in batches of 10 with 3 helpers, batch 7 with 4 helpers, 0, 1, 9, 10, 11, 20, 100, 199 and 200 orders with batch sizes 1, 7 and 10 and 1 to 5 helpers, blank rows in the middle, mixed text and number order numbers, a bad batch size, a gap or duplicate in the helper list, and an order typed below the table. Excel and Google Sheets were not opened.
search_description: Split a list into batches of 10 and assign helpers in turn in Excel and Google Sheets with INT, MOD and INDEX. Free template with batch numbers and a count per helper.
image_prompts: Three small cardboard shipping boxes lined up on a packing table, each with a different colored sticky note on the lid, beside a roll of kraft tape and a tape gun, no screens, no readable text, no logos
image_alt: Three shipping boxes with different colored sticky notes beside a roll of kraft tape
threads: Packing 45 orders with 3 helpers and nobody wants to count by hand?\nNumber the batches with INT((B5-1)/$B$1)+1, then deal orders out with INDEX(names,MOD(B5-1,COUNTA(names))+1).\n45 orders, batches of 10: 5 batches, last one has 5, 15 orders each.
---
Number each order's batch with `=INT((B5-1)/$B$1)+1`, where B5 is the order's item number (1, 2, 3...) and B1 is the batch size. Deal the orders to helpers in turn with `=INDEX(Helpers!$A$4:$A$8,MOD(B5-1,COUNTA(Helpers!$A$4:$A$8))+1)`.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same old functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Orders sheet with forty-five order numbers split into batches of ten with each order assigned in turn to one of three helpers](images/2026-10-28-team-05-split-order-list-into-batches-and-assign-helpers-int-mod-index-excel-google-sheets/order-batches-and-helper-rotation-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: number the orders 1, 2, 3 and so on. Subtract 1 so counting starts at 0, then use two operations. `INT` of the number divided by the batch size tells you which batch. `MOD` (the remainder) tells you the place inside the batch, or which name comes next when you divide by the number of helpers.

### In Excel

1. In **B1**, type the batch size, such as 10. In row 4, type the headings: Order, Item no., Batch, Position, Helper.
2. Type order numbers in **A5** and down.
3. In a second sheet named **Helpers**, type first names in **A4:A8**, one per row, with no empty row between them.
4. In **B5**, type `=IF(A5="","",COUNTA($A$5:A5))`. This counts the orders so far, so an empty row is skipped.
5. In **C5**, type `=INT((B5-1)/$B$1)+1` for the batch.
6. In **D5**, type `=MOD(B5-1,$B$1)+1` for the position in the batch.
7. In **E5**, type `=INDEX(Helpers!$A$4:$A$8,MOD(B5-1,COUNTA(Helpers!$A$4:$A$8))+1)`.
8. Select **B5:E5** and drag the fill handle down the list.

Microsoft's pages: [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) and [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef). These pages describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 7 above, typed exactly as written. Add the **Helpers** sheet with the **+** button at the bottom left.
2. Select **B5:E5**, then drag the fill handle down the list.

Google's help pages: [MOD](https://support.google.com/docs/answer/3093497) and [INDEX](https://support.google.com/docs/answer/3098242). They describe the functions, not our test.

### The version in the template

The template's E5 adds a guard so the cell shows "No helper" when the Helpers list is empty:

`=IF(B5="","",IF(COUNTA(Helpers!$A$4:$A$8)=0,"No helper",INDEX(Helpers!$A$4:$A$8,MOD(B5-1,COUNTA(Helpers!$A$4:$A$8))+1)&""))`

## Example

Sample data is fictional: 45 order numbers (4101 to 4145), a batch size of 10 and three helpers named Ana, Ben and Cora.

| Order | Item no. | Batch | Position | Helper |
|---|---|---|---|---|
| 4101 | 1 | 1 | 1 | Ana |
| 4102 | 2 | 1 | 2 | Ben |
| 4103 | 3 | 1 | 3 | Cora |
| 4104 | 4 | 1 | 4 | Ana |
| 4110 | 10 | 1 | 10 | Ana |
| 4111 | 11 | 2 | 1 | Ben |
| 4112 | 12 | 2 | 2 | Cora |
| 4113 | 13 | 2 | 3 | Ana |
| 4141 | 41 | 5 | 1 | Ben |
| 4145 | 45 | 5 | 5 | Cora |

The Summary sheet shows 5 batches, 5 orders in the last batch (45 minus 4 full batches of 10), and 15 orders each for Ana, Ben and Cora. Its check cell says OK.

The helper rotation does not restart at a new batch. Item 11 goes to Ben because item 10 went to Ana. If you want each batch to start with the same person, use the position column in place of the item number inside `MOD`.

## Troubleshooting

### Batch and Position show "Check batch size"
B1 holds 0, a negative number, a decimal such as 2.5, or text such as "ten". The template flags it in B2. Type a whole number of 1 or more.

### A helper's count looks too high
That name is listed twice in the Helpers sheet. In the test, Ana, Ben, Ana merged the counts for Ana, and the Summary check changed to CHECK. List each name once.

### Some rows have a blank helper
There is an empty row between names in **Helpers!A4:A8**. `COUNTA` counts only filled cells, so the formula reaches past the gap into empty cells. Move the names up so they sit together.

### A row is skipped or an order is not numbered
An order typed below row 204 sits outside the table, so it is not numbered and the check shows CHECK. Empty rows inside the table are skipped on purpose. A cell holding only a space counts as an order.

## Template

[Download the .xlsx](templates/tidy-tabs-batches-and-rotation.xlsx)

The file has four tabs. **Orders** holds the yellow batch size in **B1** and yellow order numbers in **A5:A204**, with 45 fictional rows filled in. **Helpers** holds up to five yellow names in **A4:A8**. Gray cells are formulas: columns **B** to **E** on Orders, and the **Summary** tab with batch count, last batch size, orders per helper and a check. The first row of each batch turns green. **How to use** has short notes.

It splits and deals a list in order. It does not weigh how long an order takes. It was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
