---
title: How to Shade Alternating Groups of Rows by Order Number in Excel and Google Sheets
labels: excel-formulas, small-business, formatting
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The group numbers were recalculated in 10 workbooks (600 rows) and compared with an independent Python model: the sample, empty, one order, one order on 10 lines, 30 single-line orders, blank rows in the middle, unsorted orders, text and number order numbers, capital letters only, and two random sorted sets. The shading was checked by exporting the template to PDF and PNG and looking at the page, and the rule was still in the file after a LibreOffice save and reopen. Excel and Google Sheets were not used. Not run: a row inserted inside the range, 200 rows.
search_description: Shade every line of an order the same color and flip the shade when the order number changes, with a helper column and one conditional formatting rule. Free template.
image_prompts: A stack of kraft packing slips fanned out with every other sheet tinted gray, a roll of twine and a small enamel pin on a wooden packing table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Fanned stack of packing slips beside twine and an enamel pin on a wooden table
threads: Plain banded rows cut an order with 4 lines in half.\nNumber the groups: =IF(A5=A4,D4,MAX($D$4:D4)+1). Then the rule =MOD($D4,2)=0 shades every other order, not every other row.\nSorted by order number first, or the same order gets two shades. Sample orders are made up.
---
Number the groups in a helper column with `=IF(A5="","",IF(A5=A4,D4,MAX($D$4:D4)+1))`, where D4 holds `=IF(A4="","",1)`. Then add a conditional formatting rule on the data range with the custom formula `=MOD($D4,2)=0`. Every line of an order gets the same shade, and the shade flips when the order number changes. Plain banding with `=MOD(ROW(),2)=0` cannot do that. The list must be sorted by order number.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. There the group numbers were recalculated and compared with a Python model, and the shading was checked by rendering the template to PDF and looking at the page. The formulas use only IF, MAX and MOD, and the rules are ordinary custom formula rules, but Excel and Google Sheets were not opened for this post. Their menu paths below come from Microsoft and Google help pages.

![Order lines table with alternating gray and white blocks following the order numbers, next to a plain banded copy](images/2026-11-09-team-04-alternate-row-colors-by-group-order-number-conditional-formatting-excel-google-sheets/shaded-order-groups-vs-plain-banding-template.png) *caption: PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The sheet has the order number in **A**, the item in **B** and the quantity in **C**, with headers in row 3 and data from row 4. Column **D** holds the group number. Sort the list by order number first.

1. In **D4**, type `=IF(A4="","",1)`. The first line is always group 1.
2. In **D5**, type `=IF(A5="","",IF(A5=A4,D4,MAX($D$4:D4)+1))`. The number stays the same while the order number stays the same, and goes up by 1 when it changes.
3. Fill **D5** down to the last row you plan to use (the template goes to row 63). Blank rows stay blank.
4. Select the whole data range, **A4:D63**, including column D.

The plain `=IF(A5=A4,D4,D4+1)` gives the same numbers on the sample, but it also numbers the empty rows below your data. The `MAX` version avoids that.

### In Excel

1. With **A4:D63** selected, choose **Home > Conditional Formatting > New Rule**.
2. Click **Use a formula to determine which cells to format**.
3. Type `=MOD($D4,2)=0` in the formula box. The row number has no dollar sign, so the rule moves down with each row.
4. Click **Format**, open the **Fill** tab, pick a light gray and click **OK**.

Microsoft describes these steps on its [conditional formatting help page](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel). To change the rule later, use **Home > Conditional Formatting > Manage Rules**.

### In Google Sheets

1. With **A4:D63** selected, choose **Format > Conditional formatting**.
2. Under **Format cells if**, pick **Custom formula is**.
3. Type `=MOD($D4,2)=0`.
4. Under **Formatting style**, choose a light gray fill and click **Done**.

Google describes these steps on its [conditional formatting help page](https://support.google.com/docs/answer/78413?hl=en).

### Without a helper column

This rule on **A4:D63** counts the order changes above each row and needs no column D:

`=AND($A4<>"",MOD(SUMPRODUCT(--($A$4:$A4<>$A$3:$A3)),2)=0)`

On the sorted sample with no blank rows it shaded the same rows as the helper rule in LibreOffice. The helper column is easier to read and to check, so the template uses it.

## Example

Sample data is fictional: 9 orders and 24 lines from a small shop. Group numbers run from 1 to 9, and the even groups are shaded.

| Order | Lines | Group no. | Shaded |
|---|---|---|---|
| 5001 | 3 | 1 | No |
| 5002 | 1 | 2 | Yes |
| 5003 | 4 | 3 | No |
| 5004 | 2 | 4 | Yes |
| 5005 | 3 | 5 | No |
| 5006 | 1 | 6 | Yes |
| 5007 | 4 | 7 | No |
| 5008 | 2 | 8 | Yes |
| 5009 | 4 | 9 | No |

Column D for rows 4 to 27 reads 1, 1, 1, 2, 3, 3, 3, 3, 4, 4, 5, 5, 5, 6, 7, 7, 7, 7, 8, 8, 9, 9, 9, 9. Below the last line it stays blank.

In the rendered page, the gray blocks sit on order 5002 (row 7), order 5004 (rows 12 and 13), order 5006 (row 17) and order 5008 (rows 22 and 23). Orders 5001, 5003, 5005, 5007 and 5009 stay white.

The template also has a **Plain banding** tab with the same 24 lines and the rule `=MOD(ROW(),2)=0`. It shades every other row, so it cuts through orders 5001, 5003, 5005, 5007 and 5009.

## Troubleshooting

### One order gets two shades
The list is not sorted by order number, so the same order appears in two places and each place counts as its own group. A test with 5003, 5001, 5001, 5003, 5002, 5003 gave groups 1, 2, 2, 3, 4, 5. Sort by column A, then the groups fix themselves.

### A blank row between lines changes the shading
In the helper version a blank row starts one new group, so the order after it flips shade. The helper-free rule counts a blank row as two changes, so the shades after it can differ from the helper version. Delete blank rows inside the list.

### The blank row itself turns gray
That happens when the rule lacks the `$A4<>""` guard. In the helper version, column D is blank on an empty row, so the rule does not shade it. In the helper-free rule, keep the `AND($A4<>"", ...)` part.

### 5001 and "5001" get different shades
An order number stored as a number and the same number stored as text count as different values. Make the whole column one type. Capital letters do not matter, so "a-17" and "A-17" are the same group.

### Nothing shades in a new row you inserted
We did not test inserting a row inside the range. After you insert one, copy the formula in column D into it and check the rule's range in **Manage Rules** (Excel) or the conditional formatting sidebar (Google Sheets).

## Template

[Download the .xlsx](templates/tidy-tabs-shade-groups.xlsx)

The file has three tabs. **Lines** has 60 slots from row 4 with the 24 fictional sample lines, the **Group no.** formulas in column D and the shading rule on **A4:D63**. **Plain banding** has the same 24 lines with the `MOD(ROW(),2)` rule for contrast. **How to use** lists the formulas and both rules.

The usual yellow input and gray formula fills are left off on purpose, because a fixed fill would hide the shading. Column D is gray text only. The shading comes from the rule, not from cell fills. The sheet prints on 1 page, US Letter.

To open it in Google Sheets, go to **File > Import > Upload** and select the file. Then check **Format > Conditional formatting**, because the rule was only checked in LibreOffice.
